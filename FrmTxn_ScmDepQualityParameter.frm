VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_ScmDepQualityParameter 
   Caption         =   "SCM Depositor Quality Parameter Mapping"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   240
      TabIndex        =   37
      Top             =   4920
      Visible         =   0   'False
      Width           =   14475
      Begin VB.ComboBox CmbAgent 
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
         Left            =   3045
         Sorted          =   -1  'True
         TabIndex        =   68
         Top             =   1260
         Width           =   3735
      End
      Begin VB.ComboBox CmbSExchangeType 
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
         Sorted          =   -1  'True
         TabIndex        =   67
         Top             =   1260
         Width           =   2895
      End
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmTxn_ScmDepQualityParameter.frx":0000
         Left            =   6780
         List            =   "FrmTxn_ScmDepQualityParameter.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   66
         Top             =   1260
         Width           =   2415
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
         Left            =   9285
         TabIndex        =   65
         Top             =   1265
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
         Left            =   10305
         TabIndex        =   64
         Top             =   1265
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
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
         Left            =   12000
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   555
         Width           =   2430
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
         Left            =   3045
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   555
         Width           =   3735
      End
      Begin VB.ComboBox CmbSCxTFNo 
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
         ItemData        =   "FrmTxn_ScmDepQualityParameter.frx":0021
         Left            =   6780
         List            =   "FrmTxn_ScmDepQualityParameter.frx":0023
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   555
         Width           =   1680
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
         Left            =   8475
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   555
         Width           =   3525
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   1590
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   555
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
         Format          =   110100481
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   135
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   555
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
         Format          =   110100481
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   14655
         _ExtentX        =   25850
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
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   7680
         TabIndex        =   63
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1155
         TabIndex        =   54
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   12915
         TabIndex        =   53
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label13 
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
         Height          =   270
         Left            =   4620
         TabIndex        =   52
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "CxTF No"
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
         Left            =   7103
         TabIndex        =   51
         Top             =   255
         Width           =   1035
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
         Left            =   14295
         TabIndex        =   50
         Top             =   1455
         Width           =   90
      End
      Begin VB.Label Label28 
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
         Left            =   11460
         TabIndex        =   49
         Top             =   1320
         Width           =   2415
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
         Left            =   375
         TabIndex        =   47
         Top             =   255
         Width           =   975
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
         Left            =   1995
         TabIndex        =   45
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Height          =   300
         Left            =   9637
         TabIndex        =   43
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Agent"
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
         Left            =   4628
         TabIndex        =   41
         Top             =   960
         Width           =   525
      End
   End
   Begin VB.Frame Frame4 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   57
         Top             =   2880
         Width           =   14775
         Begin VB.CommandButton CmdImport 
            Caption         =   "Import"
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
            TabIndex        =   59
            Top             =   150
            Width           =   1455
         End
         Begin VB.CommandButton CmdExport 
            Caption         =   "Export"
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
            TabIndex        =   58
            Top             =   150
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtValue 
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
         Left            =   3480
         TabIndex        =   36
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   14775
         Begin VB.TextBox TxtMode 
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
            TabIndex        =   55
            Top             =   1680
            Width           =   2535
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
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2160
            Width           =   6120
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2160
            Width           =   6000
         End
         Begin VB.TextBox TxtAgent 
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
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1680
            Width           =   6375
         End
         Begin VB.TextBox TxtCmp 
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
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1200
            Width           =   6375
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1200
            Width           =   6015
         End
         Begin VB.TextBox TxtCxTFDispalyDate 
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
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox TxtID 
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
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   1995
         End
         Begin MSComCtl2.DTPicker TxtDate 
            Height          =   375
            Left            =   4200
            TabIndex        =   16
            Top             =   240
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
            Format          =   110100481
            CurrentDate     =   39884
         End
         Begin VB.ComboBox CmbCxTFNo 
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
            Left            =   8280
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox TxtCxTFNo 
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
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   720
            Width           =   3735
         End
         Begin VB.ComboBox CmbLocation 
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
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   720
            Width           =   6015
         End
         Begin VB.ComboBox CmbStatus 
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
            ItemData        =   "FrmTxn_ScmDepQualityParameter.frx":0025
            Left            =   1440
            List            =   "FrmTxn_ScmDepQualityParameter.frx":002F
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   1680
            Width           =   2775
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
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox TxtLocation 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   720
            Width           =   6015
         End
         Begin VB.Label Label4 
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
            Left            =   120
            TabIndex        =   61
            Top             =   1740
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Mode"
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
            TabIndex        =   56
            Top             =   1740
            Width           =   735
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
            Left            =   7560
            TabIndex        =   35
            Top             =   2220
            Width           =   975
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
            Left            =   120
            TabIndex        =   34
            Top             =   2220
            Width           =   1695
         End
         Begin VB.Label Label29 
            Caption         =   "Agent"
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
            Left            =   7560
            TabIndex        =   27
            Top             =   1740
            Width           =   735
         End
         Begin VB.Label Label11 
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
            Height          =   255
            Left            =   7560
            TabIndex        =   26
            Top             =   1260
            Width           =   615
         End
         Begin VB.Label Label10 
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
            TabIndex        =   24
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "CDTF Date"
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
            TabIndex        =   22
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "CDTF"
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
            Left            =   7560
            TabIndex        =   20
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Date"
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
            Left            =   3600
            TabIndex        =   18
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label1 
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
            TabIndex        =   17
            Top             =   270
            Width           =   1440
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   8310
         Width           =   14790
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
            Left            =   11685
            TabIndex        =   6
            Top             =   165
            Width           =   2955
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
            Left            =   8850
            TabIndex        =   5
            Top             =   165
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
            Left            =   5895
            TabIndex        =   4
            Top             =   165
            Width           =   2955
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
            Left            =   3060
            TabIndex        =   3
            Top             =   165
            Width           =   2835
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
            Left            =   225
            TabIndex        =   13
            Top             =   165
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
            Left            =   11685
            TabIndex        =   12
            Top             =   585
            Width           =   2955
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
            Left            =   8850
            TabIndex        =   10
            Top             =   585
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
            Left            =   5895
            TabIndex        =   9
            Top             =   585
            Width           =   2955
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
            Left            =   3060
            TabIndex        =   8
            Top             =   585
            Width           =   2835
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
            TabIndex        =   7
            Top             =   585
            Width           =   2835
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4665
         Left            =   120
         TabIndex        =   29
         Top             =   3600
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   8229
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
   End
End
Attribute VB_Name = "FrmTxn_ScmDepQualityParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mEmployeeID, mDirectReportID, Edit_Flg, mCommodityID, mParameterID As Variant
Dim mSEmployeeID, mSDirectReportID, mSLocationID, mLocationID, mSMPrefix, mSSMPrefix, mSDesignationID, mCxTFNo, mSCommodityID, mSCxTFNo, mSCMPID, mSExchangeTypeID As Variant
Dim mStatus, mAgentID As Variant
Dim strSplitString() As String
Dim mSelRow, mSelCol As Variant
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim oWS1 As Excel.Worksheet
Dim ErrMsg, ExcelFileName, Path, Status As Variant
Dim mMin, mMax As Variant
Dim FCol As Variant
Dim I As Variant


Public Sub Grid_Head2()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.Cols = 8
MSHFlexGrid2.TextMatrix(1, 0) = "Row ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Truck No"
MSHFlexGrid2.TextMatrix(1, 2) = "Weighbridge Slip No"
MSHFlexGrid2.TextMatrix(1, 3) = "Total Weight"
MSHFlexGrid2.TextMatrix(1, 4) = "Tare Weight"
MSHFlexGrid2.TextMatrix(1, 5) = "Gross Weight"
MSHFlexGrid2.TextMatrix(1, 6) = "No Of Bags"
MSHFlexGrid2.TextMatrix(1, 7) = "Quality Parameters/Remarks"

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 1500
MSHFlexGrid2.ColWidth(4) = 1500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 4000

FCol = MSHFlexGrid2.Cols
End Sub


Private Sub AddCmd_Click()
Edit_Flg = "A"

Call GButtonLock(Me, False)
Call Grid_Head2

Frame1.Enabled = True
Frame2.Enabled = True

TxtID = ""
TxtDate.Value = Date

CmbLocation.Visible = True
TxtLocation.Visible = False
CmbLocation.Text = ""

CmbCxTFNo.Visible = True
TxtCxTFNo.Visible = False
CmbCxTFNo.Text = ""

CmbStatus.Visible = False
TxtStatus.Visible = True
CmbStatus.Text = ""
TxtStatus = "Pending"

mLocationID = ""
mSMPrefix = ""
mCxTFNo = ""
mStatus = "P"

TxtCommodity.Text = ""
TxtCmp.Text = ""
TxtCxTFDispalyDate.Text = ""
TxtMode.Text = ""
TxtAgent.Text = ""
TxtCreated = ""
TxtUpdated = ""

If RsTxn_ScmQualityParameter.RecordCount > 0 Then
   If CmbLocation.Enabled = True Then CmbLocation.SetFocus
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
For I = 1 To RsMst_Agent.RecordCount
    CmbAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbAgent.Text = tmp
End Sub


Private Sub CmbAgent_LostFocus()
If CmbAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection "
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID
End Sub


Private Sub CmbSCmp_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If
If CmbSCmp.Text = "" Then
   Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
Else
   Call TableMst_CMP(" Where LocationID = " & mSLocationID & " And CMPID = " & mSCMPID & "")
End If
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
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

mSCommodityID = RsMst_Commodity!CommodityID
End Sub


Private Sub CmbSCxTFNo_GotFocus()

If CmbSLocation.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbSLocation.SetFocus
   Exit Sub
End If

CmbSCxTFNo.Clear

tmp1 = CmbSCxTFNo.Text
tmp = " Select Distinct CxTF_DetailsID "
tmp = tmp & " from Txn_CxTF_Details where TxnType='D' And SCMFlag='1' And SM_Prefix='" & mSSMPrefix & "'"
tmp = tmp & " And CxTF_DetailsID In (Select CxTF_DetailsID From Txn_ScmQualityParameter) "
Call Tmp1Table

If TmpRs1.RecordCount = 0 Then
   MsgBox "No CxTF No found"
   CmbSCxTFNo.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbSCxTFNo.AddItem TmpRs1!CxTF_DetailsID
    TmpRs1.MoveNext
Next
CmbSCxTFNo.Text = tmp1
End Sub


Private Sub CmbSCxTFNo_LostFocus()
If CmbSCxTFNo.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "CxTF_DetailsID = '" & CmbSCxTFNo.Text & "'"

If TmpRs1.EOF Then
   MsgBox "Invalid CxTFNo selection "
   CmbSCxTFNo.SetFocus
   Exit Sub
End If
mSCxTFNo = TmpRs1.Fields("CxTF_DetailsID")
mSCommodityID = TmpRs1.Fields("CommodityID")
mSCMPID = TmpRs1.Fields("CMPID")
End Sub

Private Sub CmbSExchangeType_GotFocus()

tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   MsgBox "Invalid selection "
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mSExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbSLocation.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSLocation.SetFocus
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
   CmbCxTFNo.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
If tmp <> CmbSLocation.Text Then
   CmbSCxTFNo.Text = ""
End If
mSLocationID = RsMst_Location.Fields("LocationID")
mSSMPrefix = RsMst_Location.Fields("SM_Prefix")

End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "SCM Depositor Quality Parameter Mapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "SCM Depositor Quality Parameter Mapping.xls")
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

Private Sub CmdExport_Click()
If CmbLocation.Text = "" Or CmbCxTFNo.Text = "" Then Exit Sub
If MSHFlexGrid2.TextMatrix(2, 1) = "" Then Exit Sub
Gen_mTimeStamp = GetTimeStamp
FileNm = Gen_mTimeStamp & Gen_UserName & "SCMDepositorQualityParameterMapping-" & TxtID.Text & ".xls"
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "SCMDepositorQualityParameterMapping-" & TxtID.Text & ".xls")
Call Xls_Detail_Rpt_Generate
End Sub
Public Sub Xls_Detail_Rpt_Generate()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "SCMDepositorQualityParameterMapping-" & TxtID.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 0
For I = 1 To MSHFlexGrid2.Rows - 1
    mRow = mRow + 1
    For C = 0 To MSHFlexGrid2.Cols - 2
        If InStr(MSHFlexGrid2.TextMatrix(I, C + 1), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C + 1)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C + 1), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C + 1)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C + 1)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdImport_Click()
If CmbLocation.Text = "" Or CmbCxTFNo.Text = "" Then Exit Sub

If MSHFlexGrid2.TextMatrix(2, 1) = "" Then Exit Sub

On Error GoTo errortrap

For I = 2 To MSHFlexGrid2.Rows - 1
    For iCtr = 8 To MSHFlexGrid2.Cols - 1
        MSHFlexGrid2.TextMatrix(I, iCtr) = ""
    Next
Next
With CommonDialog1
   .InitDir = App.Path & "\Excel\"
   .DialogTitle = "Open Excel File..."
   .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx"
    If FileNm <> "" Then
      .FileName = FileNm
    Else
      .FileName = ""
    End If
   .CancelError = True
   .ShowOpen
End With

ExcelFileName = CommonDialog1.FileTitle
Path = CommonDialog1.FileName
   
If ExcelFileName <> "" Then
Call UploadFile

If Status = 1 Then Exit Sub

'-----------------------------
ErrMsg = ""
Status = 0

'-----------------------------
Call Validation
If Status = 1 Then
   MsgBox ErrMsg, vbExclamation, "Message"
Else
   Call FillData
   oXL.Quit
End If
End If

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub FillData()
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows

ColCount = Val(FCol) + GetMaxCountParameter()
RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
For mRow = 2 To RowCount
 If MSHFlexGrid2.Rows > 1 Then
    For u = 2 To MSHFlexGrid2.Rows - 1
        If Trim(MSHFlexGrid2.TextMatrix(u, 1)) = Trim(oWS.Cells(mRow, C + 1).Value) And Trim(MSHFlexGrid2.TextMatrix(u, 2)) = Trim(oWS.Cells(mRow, C + 2).Value) Then
           For mCol = Val(FCol) To ColCount - 1
               If oWS.Cells(mRow, mCol).Value <> "" Then
                  If LCase(Trim(oWS.Cells(mRow, mCol).Value)) = "na" Then
                     MSHFlexGrid2.TextMatrix(u, mCol) = Trim(oWS.Cells(mRow, mCol).Value)
                  Else
                     If IsNumeric(oWS.Cells(mRow, mCol).Value) And oWS.Cells(mRow, mCol).Value >= 0 Then
                        MSHFlexGrid2.TextMatrix(u, mCol) = oWS.Cells(mRow, mCol).Value
                     End If
                  End If
               End If
           Next
        End If
    Next
 End If
mDataDetailFlag = True
Next
oXL.Quit
MsgBox "Data Imported Successfully ", vbInformation
End Sub

Private Sub Validation()

Set oWS = oWB.Worksheets("Sheet1")
WrkSheet = "Sheet1"
If LCase(Trim(oWS.Cells(1, 1))) <> LCase(MSHFlexGrid2.TextMatrix(1, 1)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A(" & LCase(MSHFlexGrid2.TextMatrix(1, 0)) & ") " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 2))) <> LCase(MSHFlexGrid2.TextMatrix(1, 2)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of B(" & LCase(MSHFlexGrid2.TextMatrix(1, 1)) & ") " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 3))) <> LCase(MSHFlexGrid2.TextMatrix(1, 3)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of C(" & LCase(MSHFlexGrid2.TextMatrix(1, 2)) & ") " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 4))) <> LCase(MSHFlexGrid2.TextMatrix(1, 4)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of D(" & LCase(MSHFlexGrid2.TextMatrix(1, 3)) & ") " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 5))) <> LCase(MSHFlexGrid2.TextMatrix(1, 5)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of E(" & LCase(MSHFlexGrid2.TextMatrix(1, 4)) & ") " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 6))) <> LCase(MSHFlexGrid2.TextMatrix(1, 6)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of F(" & LCase(MSHFlexGrid2.TextMatrix(1, 5)) & ") " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 7))) <> LCase(MSHFlexGrid2.TextMatrix(1, 7)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of G(" & LCase(MSHFlexGrid2.TextMatrix(1, 6)) & ") " & vbCrLf
   Status = 1
End If
If LCase(Trim(oWS.Cells(1, 8))) <> LCase(MSHFlexGrid2.TextMatrix(1, 8)) Then
   ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of H(" & LCase(MSHFlexGrid2.TextMatrix(1, 7)) & ") " & vbCrLf
   Status = 1
End If
For mCol = 8 To MSHFlexGrid2.Cols - 1
    If LCase(Trim(oWS.Cells(1, mCol))) <> LCase(MSHFlexGrid2.TextMatrix(1, mCol)) Then
       ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of (" & LCase(MSHFlexGrid2.TextMatrix(1, mCol)) & ") " & vbCrLf
       Status = 1
       Exit For
    End If
Next
End Sub

Private Sub UploadFile()
Dim mImportFilename
Status = 0
If ExcelFileName <> "" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(Path)
   mImportFilename = ExcelFileName
   Status = 0
Else
   MsgBox "Please Select File", vbInformation, "Information"
   Status = 1
   mImportFilename = ""
End If
End Sub
Private Sub CmdSearch_Click()
Dim Wc1 As Variant

Call Grid_Head1

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To Date'!!! "
   SFrom.SetFocus
   Exit Sub
End If

tmp = " Select Distinct TSQP.ID,TSQP.EntryDate,ML.LocationName,TSQP.CxTF_DetailsID,TCD.CxTFDate,MCM.Commodity,MC.CMPName,MET.ExchangeType,MA.AgentName,TSQP.Status "
tmp = tmp & " From Txn_ScmQualityParameter TSQP "
tmp = tmp & " Inner Join Txn_CxTF_Details TCD on TCD.CxTF_DetailsID=TSQP.CxTF_DetailsID And TSQP.SM_Prefix=TCD.SM_Prefix And TxnType='D' And SCMFlag='1'"
tmp = tmp & " Inner Join Mst_Agent MA on TCD.AgentID=MA.AgentID"
tmp = tmp & " Inner Join Mst_Location ML on TSQP.LocationId=ML.LocationId And TCD.SM_Prefix=ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_Cmp MC on TCD.CmpID=MC.CmpID"
tmp = tmp & " Inner Join Mst_Commodity MCM on TCD.CommodityID=MCM.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType MET on TCD.ExchangeTypeID=MET.ExchangeTypeID"

If Gen_DBType = "MDB" Then
   Wc1 = " Where TSQP.EntryDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "#"
Else
   Wc1 = " Where TSQP.EntryDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "'"
End If

If CmbSLocation.Text <> "" Then
   Wc1 = Wc1 & " And ML.LocationName='" & CmbSLocation.Text & "'"
End If

If CmbSCxTFNo.Text <> "" Then
   Wc1 = Wc1 & " And TSQP.CxTF_DetailsID=" & mCxTFNo & ""
End If

If CmbSCommodity.Text <> "" Then
   Wc1 = Wc1 & " And MCM.Commodity='" & CmbSCommodity.Text & "'"
End If

If CmbSExchangeType.Text <> "" Then
   Wc1 = Wc1 & " And MET.ExchangeType='" & CmbSExchangeType.Text & "'"
End If

If CmbSCmp.Text <> "" Then
   Wc1 = Wc1 & " And MC.CMPName='" & CmbSCmp.Text & "'"
End If

If CmbAgent.Text <> "" Then
   Wc1 = Wc1 & " And MA.AgentName='" & CmbAgent.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   Wc1 = Wc1 & " And TSQP.Status='" & UCase(Left(CmbSStatus.Text, 1)) & "'"
End If

tmp = tmp & Wc1

tmp = tmp & " order By TSQP.ID"

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       If MSHFlexGrid1.TextMatrix(I, 9) = "A" Then
          MSHFlexGrid1.TextMatrix(I, 9) = "Approved"
       Else
          MSHFlexGrid1.TextMatrix(I, 9) = "Pending"
       End If
       MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 4) = Format(MSHFlexGrid1.TextMatrix(I, 4), "dd-MMM-yyyy")
       TmpRs1.MoveNext
   Next
End If
End Sub

Private Sub CmbCxTFNo_GotFocus()
tmp1 = CmbCxTFNo.Text

If CmbLocation.Text = "" Then Exit Sub

CmbCxTFNo.Clear

tmp = " Select Distinct CxTF_DetailsID "
tmp = tmp & " from Txn_CxTF_Details where TxnType='D' And SCMFlag='1' And SM_Prefix='" & mSMPrefix & "' And Flag='P'"
tmp = tmp & " And CxTF_DetailsID Not In (Select CxTF_DetailsID From Txn_ScmQualityParameter Where SM_Prefix='" & mSMPrefix & "') "
Call Tmp1Table

If TmpRs1.RecordCount = 0 Then
   MsgBox "No CDTF No found"
   CmbCxTF_DetailsID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbCxTFNo.AddItem TmpRs1!CxTF_DetailsID
    TmpRs1.MoveNext
Next
CmbCxTFNo.Text = tmp1
End Sub


Private Sub CmbCxTFNo_LostFocus()
If CmbCxTFNo.Text = "" Then
   CmbCxTFNo.Text = ""
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "CxTF_DetailsID = '" & CmbCxTFNo.Text & "'"

If TmpRs1.EOF Then
   MsgBox "Invalid CDTFNo selection "
   CmbCxTFNo.SetFocus
   Exit Sub
End If
mCxTFNo = TmpRs1.Fields("CxTF_DetailsID")
If mCxTFNo <> "" Then
   Call GetCxTFDetail(mCxTFNo)
   Call GetTruckDetail(mCxTFNo)
   Call GetQualityParatmeter(mCommodityID)
End If
End Sub

Private Sub CmbLocation_GotFocus()
tmp1 = CmbLocation.Text

'If Gen_WcLocation = "" Then
'   Call TableMst_Location("")
'Else
'   Call TableMst_Location("where " & Gen_WcLocation)
'End If

tmp = " Select Distinct ML.LocationName,ML.LocationID,ML.SM_Prefix"
tmp = tmp & " from Txn_CxTF_Details a"
tmp = tmp & " inner join mst_location ml on a.sm_prefix = ml.sm_prefix"
tmp = tmp & " where TxnType='D' And SCMFlag='1'"
If Gen_WcLocation <> "" Then
   tmp = tmp & " And ML." & Gen_WcLocation
End If

Call Tmp1Table

CmbLocation.Clear

'If RsMst_Location.RecordCount = 0 Then
'   MsgBox "No Location found"
'   CmbLocation.SetFocus
'   Exit Sub
'End If
'For i = 1 To RsMst_Location.RecordCount
'    CmbLocation.AddItem RsMst_Location!LocationName
'    RsMst_Location.MoveNext
'Next


If TmpRs1.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbLocation.AddItem TmpRs1!Locationname
    TmpRs1.MoveNext
Next

CmbLocation.Text = tmp1
End Sub


Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   CmbCxTFNo.Text = ""
   Exit Sub
End If

'RsMst_Location.MoveFirst
'RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
'
'If RsMst_Location.EOF Then
'   MsgBox "Invalid Location selection "
'   CmbLocation.SetFocus
'   Exit Sub
'End If
'CmbCxTFNo.Clear
'mLocationID = RsMst_Location.Fields("LocationID")
'mSMPrefix = RsMst_Location.Fields("SM_Prefix")


TmpRs1.MoveFirst
TmpRs1.Find "LocationName = '" & CmbLocation.Text & "'"

If TmpRs1.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
CmbCxTFNo.Clear
mLocationID = TmpRs1!LocationID
mSMPrefix = TmpRs1!SM_Prefix



End Sub


Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If mStatus = "A" Then
   MsgBox "Can't Edit, It's Approved"
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   tmp = " Delete From Txn_ScmQualityParameterDetail where ID=" & TxtID.Text & ""
   Call Tmp4Table
   
   RsTxn_ScmQualityParameter.Delete
   RsTxn_ScmQualityParameter.Update
   If RsTxn_ScmQualityParameter.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_ScmQualityParameter.MoveNext
   If RsTxn_ScmQualityParameter.EOF() Then
      RsTxn_ScmQualityParameter.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_ScmQualityParameter.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"

If mStatus = "A" Then
   MsgBox "This entry is approved by SCM Admin . Please contact SCM Departmenr !!!"
   Exit Sub
End If

Call GButtonLock(Me, False)

Frame1.Enabled = True
Frame2.Enabled = True

'CmbLocation.Visible = True
'TxtLocation.Visible = True
CmbLocation.Text = TxtLocation.Text
'
'CmbCxTFNo.Visible = True
'TxtCxTFNo.Visible = False
CmbCxTFNo.Text = TxtCxTFNo.Text

If LCase(Gen_UserRole) = "scmadmin" Or LCase(Gen_UserRole) = "power" Then
   CmbStatus.Visible = True
   TxtStatus.Visible = False
End If
CmbStatus.Text = TxtStatus.Text


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_ScmQualityParameter.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
    Call GButtonLockAD(Me)
    Frame1.Enabled = False
    ButtonFrm.Enabled = False
    MsgBox "Access denied !!!!!!!!!"
    Exit Sub
End If

Frame1.Enabled = True
ButtonFrm.Enabled = True
SFrom.Value = Date - 30
STo.Value = Date
Call Grid_Head1
Call Grid_Head2

Call TableTxn_ScmQualityParameter(" Order by ID")
If RsTxn_ScmQualityParameter.RecordCount = 0 Then
    Call AddCmd_Click
    Exit Sub
End If

Call GButtonLock(Me, True)
RsTxn_ScmQualityParameter.MoveLast
Call Indata

End Sub


Public Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 10

MSHFlexGrid1.TextMatrix(0, 0) = "ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 4) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "CMP"
MSHFlexGrid1.TextMatrix(0, 7) = "Mode"
MSHFlexGrid1.TextMatrix(0, 8) = "Agent"
MSHFlexGrid1.TextMatrix(0, 9) = "Status"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 3000
MSHFlexGrid1.ColWidth(9) = 1500

End Sub
Public Sub GetCmpDetails(mDirRptID As Variant)
If mDirectReportID <> "" Then
   tmp = " Select ME.EmployeeName+'~'+ME.EmployeeNo as FullName,ML.LocationName,MD.Designation "
   tmp = tmp & " From Mst_Employee ME"
   tmp = tmp & " Inner join Mst_Location ML on ME.BaseLocationID=ML.LocationID"
   tmp = tmp & " Inner join Mst_Designation MD on ME.DesignationID=MD.DesignationID"
   tmp = tmp & " where EmployeeID=" & mDirRptID & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      TxtCmpm.Text = TmpRs1!FullName
      TxtBaseLocation.Text = TmpRs1!Locationname
      TxtDesignation.Text = TmpRs1!Designation
   End If
End If
End Sub

Public Sub Indata()

Call Grid_Head2

Frame1.Enabled = False
Frame2.Enabled = False

CmbLocation.Visible = False
TxtLocation.Visible = True
CmbLocation.Text = ""

CmbCxTFNo.Visible = False
TxtCxTFNo.Visible = True
CmbCxTFNo.Text = ""

CmbStatus.Visible = False
TxtStatus.Visible = True
CmbStatus.Text = ""

TxtCommodity.Text = ""
TxtCmp = ""
TxtMode = ""
TxtAgent = ""
TxtCreated = ""
TxtUpdated = ""

tmp = " Select TSQP.ID,TSQP.EntryDate,TSQP.LocationID,TSQP.CxTF_DetailsID,TSQP.Remarks,ML.LocationName,TSQP.SM_Prefix,TSQP.Status "
tmp = tmp & " From Txn_ScmQualityParameter TSQP "
tmp = tmp & " Inner Join Mst_Location ML on TSQP.LocationID=ML.LocationID"
tmp = tmp & " Where TSQP.ID=" & RsTxn_ScmQualityParameter!ID & ""

Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   TxtID = IIf(IsNull(TmpRs5!ID), "", TmpRs5!ID)
   TxtDate = IIf(IsNull(TmpRs5!EntryDate), "", TmpRs5!EntryDate)
   TxtLocation = IIf(IsNull(TmpRs5!Locationname), "", TmpRs5!Locationname)
   TxtCxTFNo = IIf(IsNull(TmpRs5!CxTF_DetailsID), "", TmpRs5!CxTF_DetailsID)
   TxtRemarks = IIf(IsNull(TmpRs5!Remarks), 0, TmpRs5!Remarks)
   mLocationID = IIf(IsNull(TmpRs5!LocationID), 0, TmpRs5!LocationID)
   mCxTFNo = IIf(IsNull(TmpRs5!CxTF_DetailsID), "", TmpRs5!CxTF_DetailsID)
   mSMPrefix = IIf(IsNull(TmpRs5!SM_Prefix), "", TmpRs5!SM_Prefix)
   mStatus = IIf(IsNull(TmpRs5!Status), "", TmpRs5!Status)
   If mStatus = "P" Then
      TxtStatus.Text = "Pending"
   Else
      TxtStatus.Text = "Approved"
   End If
   Call GetCxTFDetail(mCxTFNo)
   TxtCreated = IIf(IsNull(RsTxn_ScmQualityParameter!CreatedBy), "", RsTxn_ScmQualityParameter!CreatedBy) & " :- " & IIf(IsNull(RsTxn_ScmQualityParameter!CreatedDate), "", RsTxn_ScmQualityParameter!CreatedDate)
   TxtUpdated = IIf(IsNull(RsTxn_ScmQualityParameter!UpdatedBy), "", RsTxn_ScmQualityParameter!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_ScmQualityParameter!UpdatedDate), "", RsTxn_ScmQualityParameter!UpdatedDate)
End If

Call InDataDetail

If RsTxn_ScmQualityParameter.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub LastCmd_Click()
RsTxn_ScmQualityParameter.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub


Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_ScmQualityParameter.MoveFirst
RsTxn_ScmQualityParameter.Find "ID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsTxn_ScmQualityParameter.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub

Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 2 Then Exit Sub
mSelRow = MSHFlexGrid2.row
mSelCol = MSHFlexGrid2.Col

With MSHFlexGrid2
    If mSelCol = 0 Then
      If .TextMatrix(mSelRow, mSelCol) = "" Then Exit Sub
         mSelCol = 0
         If .TextMatrix(mSelRow, mSelCol) = strUnChecked Then
            .TextMatrix(mSelRow, mSelCol) = strChecked
         Else
         .TextMatrix(mSelRow, mSelCol) = strUnChecked
      End If
   End If
   If mSelCol > 7 Then
      If MSHFlexGrid2.TextMatrix(mSelRow, 1) = "" Then Exit Sub
      TxtValue.Visible = True
      TxtValue.Text = .TextMatrix(.row, mSelCol)
      With TxtValue
           .Left = (MSHFlexGrid2.Left + MSHFlexGrid2.CellLeft) - 30
           .Top = (MSHFlexGrid2.Top + MSHFlexGrid2.CellTop) - 30
           .Width = MSHFlexGrid2.CellWidth + 35
           .Height = MSHFlexGrid2.CellHeight
           .ZOrder
           .Visible = True
           .SetFocus
       End With
   End If
End With
End Sub

Private Sub MSHFlexGrid2_Scroll()
If TxtValue.Visible = True Then TxtValue.Visible = False
End Sub


Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_ScmQualityParameter.MoveNext
If RsTxn_ScmQualityParameter.EOF Then
   RsTxn_ScmQualityParameter.MoveLast
   LF_Flag = "Y"
End If
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_ScmQualityParameter.MovePrevious
If RsTxn_ScmQualityParameter.BOF Then
   RsTxn_ScmQualityParameter.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SaveCmd_Click()

If CmbLocation.Text = "" Then
   MsgBox "Blank Commodity Not Allowed !!! "
   CmbLocation.SetFocus
   Exit Sub
End If

If CmbCxTF_DetailsID.Text = "" Then
   MsgBox "Blank CDTF Not Allowed !!! "
   CmbCxTF_DetailsID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows <= 2 Then
   MsgBox "Godown Details Missing!!!"
   MSHFlexGrid2.SetFocus
   Exit Sub
End If

'For i = 2 To MSHFlexGrid2.Rows - 1
' If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
'    For mCol = 8 To MSHFlexGrid2.Cols - 1
'        strSplitString = Split(MSHFlexGrid2.TextMatrix(1, mCol), "~")
'        mParameterID = strSplitString(1)
'        If CheckRange(mCommodityID, mParameterID, Val(MSHFlexGrid2.TextMatrix(i, mCol))) = True Then
'           MsgBox "Value for " & strSplitString(0) & " parameter should be in between " & mMin & " and " & mMax & ""
'           MSHFlexGrid2.Row = i
'           MSHFlexGrid2.Col = mCol
'           MSHFlexGrid2.SetFocus
'           Exit Sub
'        End If
'    Next
'    mCol = 8
' End If
'Next

For I = 2 To MSHFlexGrid2.Rows - 1
 If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
    For mCol = 8 To MSHFlexGrid2.Cols - 1
        If Trim(MSHFlexGrid2.TextMatrix(I, mCol)) = "" Then
           strSplitString = Split(MSHFlexGrid2.TextMatrix(1, mCol), "~")
           mParameterID = strSplitString(1)
           MsgBox "Please Enter Value of " & strSplitString(0) & " parameter for " & Trim(MSHFlexGrid2.TextMatrix(I, 2)) & " Weighbridge Slip No"
           MSHFlexGrid2.row = I
           MSHFlexGrid2.Col = mCol
           MSHFlexGrid2.SetFocus
           Exit Sub
        End If
    Next
    mCol = 8
 End If
Next

If Edit_Flg = "E" Then
   If CmbStatus.Text = "" Then
      MsgBox "Blank CDTF No Not Allowed !!! "
      CmbStatus.SetFocus
      Exit Sub
   End If
End If

If LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
Else
   mStatus = "A"
End If



If Edit_Flg = "A" Then
    
   mStatus = "P"
   If RsTxn_ScmQualityParameter.RecordCount > 0 Then
      RsTxn_ScmQualityParameter.MoveFirst
      tmp = "SM_Prefix='" & mSMPrefix & "' And CxTF_DetailsID=" & mCxTFNo & " "
      RsTxn_ScmQualityParameter.Filter = tmp
      If Not RsTxn_ScmQualityParameter.EOF Then
         MsgBox "Duplicate CDTF No Not Allowed !!! "
         CmbLocation.SetFocus
         Exit Sub
      End If
   End If
   RsTxn_ScmQualityParameter.AddNew
   RsTxn_ScmQualityParameter!ID = GetMaxID("ID", "Txn_ScmQualityParameter")
   TxtID.Text = RsTxn_ScmQualityParameter!ID
   RsTxn_ScmQualityParameter!G_UID = GetGUID
Else

    RsTxn_ScmQualityParameter.MoveFirst
    tmp = "SM_Prefix='" & mSMPrefix & "' And CxTF_DetailsID=" & mCxTFNo & " "
    RsTxn_ScmQualityParameter.Filter = tmp
    If Not RsTxn_ScmQualityParameter.EOF Then
       If RsTxn_ScmQualityParameter!ID <> TxtID.Text Then
          MsgBox "Duplicate CDTF No Not Allowed !!! "
          CmbLocation.SetFocus
          Exit Sub
       End If
    End If
    RsTxn_ScmQualityParameter.MoveFirst
    RsTxn_ScmQualityParameter.Find " ID= " & TxtID
End If

RsTxn_ScmQualityParameter!ID = TxtID.Text
RsTxn_ScmQualityParameter!EntryDate = TxtDate
RsTxn_ScmQualityParameter!LocationID = mLocationID
RsTxn_ScmQualityParameter!CxTF_DetailsID = mCxTFNo
RsTxn_ScmQualityParameter!SM_Prefix = mSMPrefix
RsTxn_ScmQualityParameter!Status = mStatus

If IsNull(RsTxn_ScmQualityParameter!CreatedBy) = True Or RsTxn_ScmQualityParameter!CreatedBy = "" Then
   RsTxn_ScmQualityParameter!CreatedBy = Gen_UserLoginID
   RsTxn_ScmQualityParameter!CreatedDate = Now
Else
   RsTxn_ScmQualityParameter!UpdatedBy = Gen_UserLoginID
   RsTxn_ScmQualityParameter!UpdatedDate = Now
End If

RsTxn_ScmQualityParameter.Update
Call SaveQualityParameterMapping

Call Indata

End Sub

Public Sub SaveQualityParameterMapping()
Dim mParameterID As Variant, mCol As Double
mCol = Val(FCol)

tmp = " Insert into Log_TxnScmQualityParameterDetail" & _
      " (ID," & _
      " TruckNo," & _
      " WeighBridgeSlipno," & _
      " QualityParameterID," & _
      " oldValue," & _
      " UpdatedBy)" & _
      " ( Select ID,TruckNo,WeighBridgeSlipno,QualityParameterID,Value,'" & Gen_UserLoginID & "'" & _
      " From Txn_ScmQualityParameterDetail Where ID= " & Val(TxtID.Text) & ")"

Call Tmp1Table

tmp = " Delete from Txn_ScmQualityParameterDetail where ID = " & TxtID.Text & " "
Call Tmp4Table

Call TableTxn_ScmQualityParameterDetail("Where ID = " & TxtID)

If MSHFlexGrid2.Rows > 2 Then
   For mCol = mCol To MSHFlexGrid2.Cols - 1
       For I = 2 To MSHFlexGrid2.Rows - 1
           If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
       
          If MSHFlexGrid2.TextMatrix(I, 0) = "" Then
             RsTxn_ScmQualityParameterDetail.AddNew
             RsTxn_ScmQualityParameterDetail!DID = GetMaxID("DID", "Txn_ScmQualityParameterDetail")
             RsTxn_ScmQualityParameterDetail!G_UID = GetGUID
          Else
             MAPID = IIf(MAPID = "", MSHFlexGrid2.TextMatrix(I, 0), MAPID & "," & MSHFlexGrid2.TextMatrix(I, 0))
             RsTxn_ScmQualityParameterDetail.MoveFirst
             RsTxn_ScmQualityParameterDetail.Find ("DID = " & MSHFlexGrid2.TextMatrix(I, 0))
          End If
        
          RsTxn_ScmQualityParameterDetail!ID = TxtID
          RsTxn_ScmQualityParameterDetail!TruckNo = MSHFlexGrid2.TextMatrix(I, 1)
          RsTxn_ScmQualityParameterDetail!WeighBridgeSlipno = MSHFlexGrid2.TextMatrix(I, 2)
          RsTxn_ScmQualityParameterDetail!TotalWeight = MSHFlexGrid2.TextMatrix(I, 3)
          RsTxn_ScmQualityParameterDetail!TareWeight = MSHFlexGrid2.TextMatrix(I, 4)
          RsTxn_ScmQualityParameterDetail!GrossWeight = MSHFlexGrid2.TextMatrix(I, 5)
          RsTxn_ScmQualityParameterDetail!NoofBags = MSHFlexGrid2.TextMatrix(I, 6)
          RsTxn_ScmQualityParameterDetail!Remarks = MSHFlexGrid2.TextMatrix(I, 7)
          
          If IsNull(RsTxn_ScmQualityParameterDetail!CreatedBy) = True Or RsTxn_ScmQualityParameterDetail!CreatedBy = "" Then
             RsTxn_ScmQualityParameterDetail!CreatedBy = Gen_UserLoginID
             RsTxn_ScmQualityParameterDetail!CreatedDate = Now
          Else
             RsTxn_ScmQualityParameterDetail!UpdatedBy = Gen_UserLoginID
             RsTxn_ScmQualityParameterDetail!UpdatedDate = Now
          End If
          
          strSplitString = Split(MSHFlexGrid2.TextMatrix(1, mCol), "~")
          mParameterID = strSplitString(1)
          RsTxn_ScmQualityParameterDetail!QualityParameterID = mParameterID
          
          If LCase(MSHFlexGrid2.TextMatrix(I, mCol)) = "na" Then
             RsTxn_ScmQualityParameterDetail!Value = -1
          Else
             RsTxn_ScmQualityParameterDetail!Value = Val(MSHFlexGrid2.TextMatrix(I, mCol))
          End If
          RsTxn_ScmQualityParameterDetail.Update
          RsTxn_ScmQualityParameterDetail.MoveNext
      Next
   Next
End If
End Sub

Public Function GetDirectReportID(EmpID As Variant) As Variant
tmp = " Select DirectReportingTo from Mst_Employee Where EmployeeID=" & EmpID & ""
Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   GetDirectReportID = TmpRs3!DirectReportingTo
End If
End Function

Public Sub InDataDetail()
Call Grid_Head2

Call GetQualityParatmeter(mCommodityID)

tmp = " Select TSQPD.TruckNo,TSQPD.WeighBridgeSlipno,TSQPD.TotalWeight,TSQPD.TareWeight,TSQPD.GrossWeight,TSQPD.NoofBags,TSQPD.Remarks"
tmp = tmp & " from Txn_ScmQualityParameter TSQP"
tmp = tmp & " Inner Join Txn_ScmQualityParameterDetail TSQPD on TSQP.ID=TSQPD.ID"
tmp = tmp & " Inner Join Mst_QualityParameter MQP on TSQPD.QualityParameterID=MQP.QualityParameterID"
tmp = tmp & " Where TSQPD.ID=" & TxtID.Text & ""
tmp = tmp & " Group By TSQPD.TruckNo,TSQPD.WeighBridgeSlipno,TSQPD.TotalWeight,TSQPD.TareWeight,TSQPD.GrossWeight,TSQPD.NoofBags,TSQPD.Remarks"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
MSHFlexGrid2.Rows = TmpRs1.RecordCount + 2
For I = 2 To TmpRs1.RecordCount + 1
    'MSHFlexGrid2.TextMatrix(i, 0) = TmpRs1!DID
    MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!TruckNo
    MSHFlexGrid2.TextMatrix(I, 2) = TmpRs1!WeighBridgeSlipno
    MSHFlexGrid2.TextMatrix(I, 3) = TmpRs1!TotalWeight
    MSHFlexGrid2.TextMatrix(I, 4) = TmpRs1!TareWeight
    MSHFlexGrid2.TextMatrix(I, 5) = TmpRs1!GrossWeight
    MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!NoofBags), 0, TmpRs1!NoofBags)
    MSHFlexGrid2.TextMatrix(I, 7) = TmpRs1!Remarks

    For mCol = FCol To MSHFlexGrid2.Cols - 1
        strSplitString = Split(MSHFlexGrid2.TextMatrix(1, mCol), "~")
        mParameterID = strSplitString(1)
        tmp = " Select Value From Txn_ScmQualityParameterDetail where TruckNo='" & MSHFlexGrid2.TextMatrix(I, 1) & "' And WeighBridgeSlipno='" & MSHFlexGrid2.TextMatrix(I, 2) & "' And QualityParameterID=" & mParameterID & ""
        Call Tmp2Table
        If TmpRs2.RecordCount > 0 Then
           MSHFlexGrid2.TextMatrix(I, mCol) = IIf(IsNull(TmpRs2!Value), 0, TmpRs2!Value)
           If MSHFlexGrid2.TextMatrix(I, mCol) = "-1" Then
              MSHFlexGrid2.TextMatrix(I, mCol) = "NA"
           End If
        End If
    Next
    
    TmpRs1.MoveNext
Next
End If
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_ScmQualityParameter.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_ScmQualityParameter.MoveFirst
   RsTxn_ScmQualityParameter.Find "ID = " & Val(TxtID) & ""
   If RsTxn_ScmQualityParameter.EOF Then
      RsTxn_ScmQualityParameter.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 800
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 150
Frame0.Width = Me.Width - 300
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 500
MSHFlexGrid1.Height = Frame0.Height - 2000

'Call Grid_Head1
'
'SFrom = Date - 30
'STo = Date
'RsTxn_ScmQualityParameter.MoveFirst
'RsTxn_ScmQualityParameter.Filter = ""

End Sub

Public Sub GetCxTFDetail(CxTF As Variant)

tmp = " Select TCD.CxTFDate,MCM.Commodity,MC.CMPName,MET.ExchangeType,MA.AgentName,MCM.CommodityID "
tmp = tmp & " From Txn_CxTF_Details TCD"
tmp = tmp & " Inner Join Mst_Cmp MC on TCD.CmpID=MC.CmpID"
tmp = tmp & " Inner Join Mst_Commodity MCM on TCD.CommodityID=MCM.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType MET on TCD.ExchangeTypeID=MET.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Agent MA on TCD.AgentID=MA.AgentID"
tmp = tmp & " Where TCD.CxTF_DetailsID=" & CxTF & " And TCD.SM_Prefix='" & mSMPrefix & "' And TCD.TxnType = 'D' "
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   TxtCxTFDispalyDate.Text = IIf(IsNull(TmpRs2!CxTFDate), "", Format(TmpRs2!CxTFDate, "dd-MMM-yyyy"))
   TxtCommodity.Text = IIf(IsNull(TmpRs2!Commodity), "", TmpRs2!Commodity)
   TxtCmp.Text = IIf(IsNull(TmpRs2!CMPName), "", TmpRs2!CMPName)
   TxtMode.Text = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
   TxtAgent.Text = IIf(IsNull(TmpRs2!AgentName), "", TmpRs2!AgentName)
   mCommodityID = IIf(IsNull(TmpRs2!CommodityID), "", TmpRs2!CommodityID)
End If
End Sub

Public Sub GetTruckDetail(CxTFNO As Variant)

Call Grid_Head2

Call TableTxn_CxTF_TruckDetail(" Where TxnType='D' AND CxTF_DetailsID = " & mCxTFNo & " And SM_Prefix = '" & mSMPrefix & "' Order by TruckNo")

MSHFlexGrid2.Rows = RsTxn_CxTF_TruckDetail.RecordCount + 2
For I = 2 To RsTxn_CxTF_TruckDetail.RecordCount + 1
    MSHFlexGrid2.TextMatrix(I, 1) = RsTxn_CxTF_TruckDetail!TruckNo
    MSHFlexGrid2.TextMatrix(I, 2) = RsTxn_CxTF_TruckDetail!WeighBridgeSlipno
    MSHFlexGrid2.TextMatrix(I, 3) = RsTxn_CxTF_TruckDetail!TotalWeight
    MSHFlexGrid2.TextMatrix(I, 4) = RsTxn_CxTF_TruckDetail!TareWeight
    MSHFlexGrid2.TextMatrix(I, 5) = RsTxn_CxTF_TruckDetail!GrossWeight
    MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(RsTxn_CxTF_TruckDetail!NoofBags), 0, RsTxn_CxTF_TruckDetail!NoofBags)
    MSHFlexGrid2.TextMatrix(I, 7) = RsTxn_CxTF_TruckDetail!Remarks
    RsTxn_CxTF_TruckDetail.MoveNext
Next
End Sub

Public Sub GetQualityParatmeter(mComID As Variant)
Dim I As Variant
mCol = FCol
Call TableMst_QualityParameter("Where QualityParameterID in (Select QualityParameterID From Mst_CommodityQualityParameter Where CommodityID=" & mComID & ")")
If RsMst_QualityParameter.RecordCount > 0 Then
   For I = 1 To RsMst_QualityParameter.RecordCount
       With MSHFlexGrid2
           MSHFlexGrid2.Cols = MSHFlexGrid2.Cols + 1
           MSHFlexGrid2.TextMatrix(1, mCol) = RsMst_QualityParameter!QualityParameter & "~" & RsMst_QualityParameter!QualityParameterID
           MSHFlexGrid2.ColWidth(mCol) = 1500
       End With
   RsMst_QualityParameter.MoveNext
   mCol = mCol + 1
   Next
End If
End Sub

Private Sub TxtValue_GotFocus()
If mSelCol > Val(FCol) Then
   TxtValue.SelStart = 0
   TxtValue.SelLength = Len(TxtValue.Text)
End If
End Sub


Private Sub TxtValue_LostFocus()

If TxtValue.Text = "" Then
   MSHFlexGrid2.TextMatrix(mSelRow, mSelCol) = ""
   TxtValue.Visible = False
   Exit Sub
End If

If mSelCol > Val(FCol) Then
   If TxtValue.Text <> "" Then
      If LCase(TxtValue.Text) = "na" Then
      
      Else
         If Not IsNumeric(TxtValue.Text) Or Val(TxtValue.Text) < 0 Then
            MsgBox "Invalid Numeric Format !!!"
            TxtValue.SetFocus
            Exit Sub
         End If
      End If
   End If
End If

strSplitString = Split(MSHFlexGrid2.TextMatrix(1, mSelCol), "~")
mParameterID = strSplitString(1)

'If CheckRange(mCommodityID, mParameterID, TxtValue) = True Then
'   MsgBox "Value for " & strSplitString(0) & " parameter should be in between " & mMin & " and " & mMax & ""
'   TxtValue.SetFocus
'   Exit Sub
'End If

MSHFlexGrid2.TextMatrix(mSelRow, mSelCol) = TxtValue.Text
TxtValue.Visible = False
MSHFlexGrid2.SetFocus
End Sub



Public Function CheckRange(CommId As Variant, ParaID As Variant, Valu As Double) As Boolean
tmp = " Select MinValue,MaxValue from Mst_CommodityQualityParameter where CommodityID=" & CommId & " And QualityParameterID=" & ParaID & ""
Call Tmp5Table
If TmpRs5.RecordCount > 0 Then
   If Not IsNull(TmpRs5!MinValue) And Not IsNull(TmpRs5!MaxValue) Then
      If Valu < Val(TmpRs5!MinValue) Or Valu > Val(TmpRs5!MaxValue) Then
         mMin = Val(TmpRs5!MinValue): mMax = Val(TmpRs5!MaxValue)
         CheckRange = True
         Exit Function
      End If
   End If
End If
CheckRange = False
End Function


Public Function GetMaxCountParameter() As Variant
tmp = " Select Count(*) as TotParameter From Mst_CommodityQualityParameter Where CommodityID=" & mCommodityID & ""
Call TmpTable
If Not IsNull(TmpRs!TotParameter) Then
   GetMaxCountParameter = TmpRs!TotParameter
   Exit Function
End If
GetMaxCountParameter = 0
End Function
