VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_FumigationAgreement 
   Caption         =   "Fumigation Agreement"
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
      Left            =   15240
      TabIndex        =   83
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   10200
      Width           =   10935
   End
   Begin VB.PictureBox Frame4 
      Height          =   9375
      Left            =   0
      ScaleHeight     =   9315
      ScaleWidth      =   14910
      TabIndex        =   0
      Top             =   0
      Width           =   14970
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
         Height          =   2085
         Left            =   -1320
         TabIndex        =   53
         Top             =   840
         Visible         =   0   'False
         Width           =   14715
         Begin VB.ComboBox CmbSType 
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
            ItemData        =   "FrmMst_FumigationAgreement.frx":0000
            Left            =   3000
            List            =   "FrmMst_FumigationAgreement.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   570
            Width           =   2055
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
            Left            =   12465
            TabIndex        =   41
            Top             =   1025
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
            Left            =   13470
            TabIndex        =   42
            Top             =   1025
            Width           =   1000
         End
         Begin VB.ComboBox CmbSVendor 
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
            Left            =   8070
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   570
            Width           =   4365
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
            Left            =   5175
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   570
            Width           =   2805
         End
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
            ItemData        =   "FrmMst_FumigationAgreement.frx":0024
            Left            =   12480
            List            =   "FrmMst_FumigationAgreement.frx":0031
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   570
            Width           =   1935
         End
         Begin VB.TextBox TxtSAgreeID 
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
            Left            =   1575
            MaxLength       =   12
            TabIndex        =   40
            Top             =   1020
            Width           =   1290
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   375
            Left            =   1575
            TabIndex        =   35
            Top             =   563
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   111411201
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   135
            TabIndex        =   34
            Top             =   563
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   111411201
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   54
            Top             =   1440
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   1320
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Type"
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
            Left            =   3780
            TabIndex        =   84
            Top             =   210
            Width           =   480
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
            Left            =   1875
            TabIndex        =   62
            Top             =   210
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
            Left            =   330
            TabIndex        =   61
            Top             =   210
            Width           =   945
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
            Left            =   8220
            TabIndex        =   60
            Top             =   1073
            Width           =   2415
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
            Left            =   10695
            TabIndex        =   59
            Top             =   1080
            Width           =   90
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Vendor"
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
            Left            =   9922
            TabIndex        =   58
            Top             =   210
            Width           =   660
         End
         Begin VB.Label Label13 
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
            Left            =   6195
            TabIndex        =   57
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   13170
            TabIndex        =   56
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Agreement ID"
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
            TabIndex        =   55
            Top             =   1080
            Width           =   1230
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3555
         Left            =   60
         TabIndex        =   64
         Top             =   15
         Width           =   14775
         Begin VB.ComboBox CmbVendor 
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
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   1215
            Width           =   12495
         End
         Begin VB.ComboBox CmbLocation 
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
            Left            =   8025
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   765
            Width           =   6150
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
            ItemData        =   "FrmMst_FumigationAgreement.frx":0052
            Left            =   4800
            List            =   "FrmMst_FumigationAgreement.frx":005F
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   765
            Width           =   1800
         End
         Begin VB.ComboBox CmbType 
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
            ItemData        =   "FrmMst_FumigationAgreement.frx":0080
            Left            =   1680
            List            =   "FrmMst_FumigationAgreement.frx":008A
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   765
            Width           =   1770
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
            Left            =   4800
            Locked          =   -1  'True
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   765
            Width           =   1800
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
            Left            =   8025
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   765
            Width           =   6150
         End
         Begin VB.TextBox TxtVendor 
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
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   1215
            Width           =   12495
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
            Height          =   1320
            Left            =   1680
            MaxLength       =   2500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   1635
            Width           =   12495
         End
         Begin VB.TextBox TxtType 
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
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   765
            Width           =   1770
         End
         Begin VB.CommandButton CmdVendor 
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
            Left            =   14280
            TabIndex        =   66
            Top             =   1215
            Width           =   375
         End
         Begin VB.CommandButton CmdLocation 
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
            Left            =   14280
            TabIndex        =   65
            Top             =   765
            Width           =   375
         End
         Begin VB.TextBox TxtAgreementID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1770
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
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   3045
            Width           =   6270
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3045
            Width           =   4920
         End
         Begin MSComCtl2.DTPicker TxtExpiryDate 
            Height          =   375
            Left            =   11400
            TabIndex        =   4
            Top             =   240
            Width           =   1680
            _ExtentX        =   2963
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
            Format          =   111411201
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtStartDate 
            Height          =   375
            Left            =   8025
            TabIndex        =   3
            Top             =   240
            Width           =   1680
            _ExtentX        =   2963
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
            Format          =   111411201
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtEntryDate 
            Height          =   360
            Left            =   4800
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   635
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
            Format          =   111411201
            CurrentDate     =   39884
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Vendor"
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
            TabIndex        =   81
            Top             =   1275
            Width           =   765
         End
         Begin VB.Label Label1 
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
            TabIndex        =   80
            Top             =   825
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Type"
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
            TabIndex        =   79
            Top             =   825
            Width           =   555
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   3720
            TabIndex        =   78
            Top             =   825
            Width           =   660
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Start Date"
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
            TabIndex        =   77
            Top             =   300
            Width           =   1050
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Expiry Date"
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
            Left            =   9960
            TabIndex        =   76
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   120
            TabIndex        =   75
            Top             =   2175
            Width           =   825
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   6840
            TabIndex        =   74
            Top             =   3105
            Width           =   885
         End
         Begin VB.Label Label21 
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
            Left            =   120
            TabIndex        =   73
            Top             =   3105
            Width           =   855
         End
         Begin VB.Label LblAgreemID 
            AutoSize        =   -1  'True
            Caption         =   "Agreement ID"
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
            TabIndex        =   72
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label LblEntryDate 
            AutoSize        =   -1  'True
            Caption         =   "Entry Date"
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
            Left            =   3720
            TabIndex        =   71
            Top             =   300
            Width           =   930
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1080
         Left            =   60
         TabIndex        =   63
         Top             =   8175
         Width           =   14775
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   8685
            TabIndex        =   27
            Top             =   195
            Width           =   2820
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   5865
            TabIndex        =   26
            Top             =   195
            Width           =   2820
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "&Save"
            Height          =   400
            Left            =   3045
            TabIndex        =   25
            Top             =   195
            Width           =   2820
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   225
            TabIndex        =   24
            Top             =   195
            Width           =   2820
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   11505
            TabIndex        =   33
            Top             =   600
            Width           =   2820
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   11505
            TabIndex        =   28
            Top             =   195
            Width           =   2820
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   8685
            TabIndex        =   32
            Top             =   600
            Width           =   2820
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5865
            TabIndex        =   31
            Top             =   600
            Width           =   2820
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3045
            TabIndex        =   30
            Top             =   600
            Width           =   2820
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   225
            TabIndex        =   29
            Top             =   600
            Width           =   2820
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4620
         Left            =   60
         TabIndex        =   12
         Top             =   3570
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8149
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Treatment Detail"
         TabPicture(0)   =   "FrmMst_FumigationAgreement.frx":00A4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Other Quotation Detail"
         TabPicture(1)   =   "FrmMst_FumigationAgreement.frx":00C0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(1)=   "MSHFlexGrid3"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Agreement Attachment"
         TabPicture(2)   =   "FrmMst_FumigationAgreement.frx":00DC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "CmdGetFile"
         Tab(2).Control(1)=   "CmdAttachment"
         Tab(2).Control(2)=   "CmdRemoveAttachment"
         Tab(2).Control(3)=   "TxtAttachmentRemarks"
         Tab(2).Control(4)=   "MSHFlexGrid5"
         Tab(2).Control(5)=   "lblAttachment"
         Tab(2).Control(6)=   "Label10"
         Tab(2).ControlCount=   7
         Begin VB.CommandButton CmdGetFile 
            Caption         =   "Get File"
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
            Left            =   -66600
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   1410
            Width           =   2070
         End
         Begin VB.CommandButton CmdAttachment 
            Caption         =   "Add Attachment"
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
            Left            =   -64530
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   1410
            Width           =   2070
         End
         Begin VB.CommandButton CmdRemoveAttachment 
            Caption         =   "Remove Attachment"
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
            Left            =   -62460
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   1410
            Width           =   2070
         End
         Begin VB.TextBox TxtAttachmentRemarks 
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
            Height          =   510
            Left            =   -68880
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   86
            Top             =   765
            Width           =   8475
         End
         Begin VB.Frame Frame2 
            Height          =   705
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   14535
            Begin VB.CommandButton CmdAddToListTre 
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
               Height          =   360
               Left            =   10740
               TabIndex        =   16
               Top             =   240
               Width           =   1815
            End
            Begin VB.CommandButton CmdDeleteFromListTre 
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
               Height          =   360
               Left            =   12555
               TabIndex        =   17
               Top             =   240
               Width           =   1815
            End
            Begin VB.CommandButton CmdTreatmentType 
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
               Left            =   7080
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   240
               Width           =   375
            End
            Begin VB.ComboBox CmbTreatmentType 
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
               ItemData        =   "FrmMst_FumigationAgreement.frx":00F8
               Left            =   2160
               List            =   "FrmMst_FumigationAgreement.frx":0105
               Sorted          =   -1  'True
               TabIndex        =   13
               Top             =   240
               Width           =   4845
            End
            Begin VB.TextBox TxtRate 
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
               Left            =   8970
               MaxLength       =   12
               TabIndex        =   15
               Top             =   240
               Width           =   1680
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Treatment Type"
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
               TabIndex        =   50
               Top             =   300
               Width           =   1440
            End
            Begin VB.Label lblrate 
               AutoSize        =   -1  'True
               Caption         =   "Rate per MT"
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
               Left            =   7560
               TabIndex        =   49
               Top             =   300
               Width           =   1125
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1545
            Left            =   -74880
            TabIndex        =   43
            Top             =   360
            Width           =   14535
            Begin VB.TextBox TxtRateperMTQuo 
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
               Left            =   12810
               MaxLength       =   12
               TabIndex        =   20
               Top             =   240
               Width           =   1560
            End
            Begin VB.CommandButton CmdDeleteFromListQuo 
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
               Height          =   360
               Left            =   12660
               TabIndex        =   23
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CommandButton CmdAddToListQuo 
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
               Height          =   360
               Left            =   12660
               TabIndex        =   22
               Top             =   720
               Width           =   1695
            End
            Begin VB.ComboBox CmbVendorQuo 
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
               ItemData        =   "FrmMst_FumigationAgreement.frx":0126
               Left            =   1200
               List            =   "FrmMst_FumigationAgreement.frx":0133
               Sorted          =   -1  'True
               TabIndex        =   18
               Top             =   240
               Width           =   4485
            End
            Begin VB.ComboBox CmbTreatmentTypeQuo 
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
               ItemData        =   "FrmMst_FumigationAgreement.frx":0154
               Left            =   7320
               List            =   "FrmMst_FumigationAgreement.frx":0161
               Sorted          =   -1  'True
               TabIndex        =   19
               Top             =   240
               Width           =   4125
            End
            Begin VB.TextBox TxtRemarkQuo 
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
               Height          =   720
               Left            =   1200
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Top             =   720
               Width           =   11295
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Rate per MT"
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
               Left            =   11520
               TabIndex        =   47
               Top             =   300
               Width           =   1125
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Treatment Type"
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
               Left            =   5760
               TabIndex        =   46
               Top             =   300
               Width           =   1440
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Vendor"
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
               TabIndex        =   45
               Top             =   300
               Width           =   660
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
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
               Height          =   240
               Left            =   120
               TabIndex        =   44
               Top             =   960
               Width           =   825
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3375
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   5953
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
            Height          =   2535
            Left            =   -74880
            TabIndex        =   52
            Top             =   1920
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   4471
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   2670
            Left            =   -74880
            TabIndex        =   90
            Top             =   1830
            Width           =   14505
            _ExtentX        =   25585
            _ExtentY        =   4710
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
         Begin VB.Label lblAttachment 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Path of  file selected for attachment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1245
            Left            =   -74880
            TabIndex        =   92
            Top             =   510
            Width           =   5880
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Attachment Remarks"
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
            Left            =   -65715
            TabIndex        =   91
            Top             =   480
            Width           =   2145
         End
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMst_FumigationAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mVendorID, mTreatmentTypeID As Variant
Dim mVendorQuoID, mTreatmentTypeQuoID As Variant
Dim Edit_Flg As String
Dim File_Path As String

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

Frame1.Enabled = True

Call ClearControl
Call ClearTreatmentDetail
Call ClearQuotationDetail
Call ComboVisible(True)

TxtStatus.Text = "Pending"
CmbStatus.Text = TxtStatus.Text
TxtStatus.Visible = True
CmbStatus.Visible = False

Frame2.Enabled = True
Frame3.Enabled = True

TxtRemark.Locked = False
TxtStartDate.Enabled = True
TxtExpiryDate.Enabled = True

lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True

Call Grid_Head4

If RsMst_FumigationAgreement.RecordCount > 0 Then
   TxtStartDate.SetFocus
End If
End Sub

Private Sub Grid_Head4()
With MSHFlexGrid5
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "SR No"
    .TextMatrix(0, 1) = "Attachment File Name"
    .TextMatrix(0, 2) = "Remarks"
    .TextMatrix(0, 3) = "File Path"
    .TextMatrix(0, 4) = "ID"
    .ColWidth(0) = 1000
    .ColWidth(1) = 5000
    .ColWidth(2) = 8200
    .ColWidth(3) = 0
    .ColWidth(4) = 0
End With
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocation.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   MsgBox "Invalid  selection!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection!!! "
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection!!! "
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSType_LostFocus()
If CmbSType.Text = "" Then Exit Sub
If LCase(CmbSType.Text) <> "spraying" And LCase(CmbSType.Text) <> "fumigation" Then
   MsgBox "Invalid Selection!!!"
   CmbSType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSVendor_GotFocus()
tmp = CmbSVendor.Text
Call TableMst_Vendor
CmbSVendor.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
   CmbSVendor.AddItem RsMst_Vendor!VendorName
   RsMst_Vendor.MoveNext
Next
CmbSVendor.Text = tmp

End Sub

Private Sub CmbSVendor_LostFocus()
If CmbSVendor.Text = "" Then
   Exit Sub
End If

RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbSVendor.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbSVendor.SetFocus
   Exit Sub
End If

mVendorID = RsMst_Vendor!VendorID
End Sub

Private Sub CmbTreatmentType_GotFocus()
tmp = CmbTreatmentType.Text
Call TableMst_TreatmentType

CmbTreatmentType.Clear
If RsMst_TreatmentType.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_TreatmentType.RecordCount
   CmbTreatmentType.AddItem RsMst_TreatmentType!TreatmentType
   RsMst_TreatmentType.MoveNext
Next
CmbTreatmentType.Text = tmp
End Sub

Private Sub CmbTreatmentType_LostFocus()
If CmbTreatmentType.Text = "" Then
   Exit Sub
End If

RsMst_TreatmentType.MoveFirst
RsMst_TreatmentType.Find "TreatmentType = '" & CmbTreatmentType.Text & "'"
If RsMst_TreatmentType.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbTreatmentType.SetFocus
   Exit Sub
End If

mTreatmentTypeID = RsMst_TreatmentType!TreatmentTypeID
End Sub

Private Sub CmbTreatmentTypeQuo_GotFocus()
tmp = CmbTreatmentTypeQuo.Text
Call TableMst_TreatmentType

CmbTreatmentTypeQuo.Clear
If RsMst_TreatmentType.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_TreatmentType.RecordCount
   CmbTreatmentTypeQuo.AddItem RsMst_TreatmentType!TreatmentType
   RsMst_TreatmentType.MoveNext
Next
CmbTreatmentTypeQuo.Text = tmp
End Sub

Private Sub CmbTreatmentTypeQuo_LostFocus()
If CmbTreatmentTypeQuo.Text = "" Then
   Exit Sub
End If

RsMst_TreatmentType.MoveFirst
RsMst_TreatmentType.Find "TreatmentType = '" & CmbTreatmentTypeQuo.Text & "'"
If RsMst_TreatmentType.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbTreatmentTypeQuo.SetFocus
   Exit Sub
End If

mTreatmentTypeQuoID = RsMst_TreatmentType!TreatmentTypeID
End Sub

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then Exit Sub
If LCase(CmbType.Text) <> "spraying" And LCase(CmbType.Text) <> "fumigation" Then
   MsgBox "Invalid Selection!!!"
   CmbType.SetFocus
   Exit Sub
End If
If LCase(CmbType.Text) = "fumigation" Then
   lblrate.Caption = "Rate per MT"
   MSHFlexGrid2.TextMatrix(0, 2) = "Rate per MT"
ElseIf LCase(CmbType.Text) = "spraying" Then
   lblrate.Caption = "Rate per Sq. Ft"
   MSHFlexGrid2.TextMatrix(0, 2) = "Rate per Sq. Ft"
End If

End Sub

Private Sub CmbVendor_GotFocus()
tmp = CmbVendor.Text
Call TableMst_Vendor
CmbVendor.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
   CmbVendor.AddItem RsMst_Vendor!VendorName
   RsMst_Vendor.MoveNext
Next
CmbVendor.Text = tmp

End Sub

Private Sub CmbVendor_LostFocus()
If CmbVendor.Text = "" Then
   Exit Sub
End If

RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbVendor.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbVendor.SetFocus
   Exit Sub
End If

mVendorID = RsMst_Vendor!VendorID
End Sub

Private Sub CmbVendorQuo_GotFocus()
tmp = CmbVendorQuo.Text
Call TableMst_Vendor
CmbVendorQuo.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
   CmbVendorQuo.AddItem RsMst_Vendor!VendorName
   RsMst_Vendor.MoveNext
Next
CmbVendorQuo.Text = tmp
End Sub

Private Sub CmbVendorQuo_LostFocus()
If CmbVendorQuo.Text = "" Then
   Exit Sub
End If

RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbVendorQuo.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbVendorQuo.SetFocus
   Exit Sub
End If

mVendorQuoID = RsMst_Vendor!VendorID
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdAddToListQuo_Click()
If Trim(CmbVendorQuo.Text) = "" Then
   MsgBox "Blank Vendor not allowed!!!"
   CmbVendorQuo.SetFocus
   Exit Sub
End If
If Trim(CmbTreatmentTypeQuo.Text) = "" Then
   MsgBox "Blank Treatment Type not allowed!!!"
   CmbTreatmentTypeQuo.SetFocus
   Exit Sub
End If
If Trim(TxtRateperMTQuo.Text) = "" Then
   MsgBox "Blank Rate per MT not allowed!!!"
   TxtRateperMTQuo.SetFocus
   Exit Sub
End If

With MSHFlexGrid3
   For I = 1 To .Rows - 1
      If CmdAddToListQuo.Caption = "Update List" Then
         If I <> .RowSel Then
            If .TextMatrix(I, 1) = CmbVendorQuo.Text And .TextMatrix(I, 2) = CmbTreatmentTypeQuo.Text Then
               MsgBox "Duplicate Vendor and Treatment Type not Allowed!!!"
               Exit Sub
            End If
         End If
      Else
         If .TextMatrix(I, 1) = CmbVendorQuo.Text And .TextMatrix(I, 2) = CmbTreatmentTypeQuo.Text Then
            MsgBox "Duplicate Vendor and Treatment Type not Allowed!!!"
            Exit Sub
         End If
      End If
   Next

   If CmdAddToListQuo.Caption = "Update List" Then
      I = .RowSel
   Else
      I = .Rows - 1
      .Rows = .Rows + 1
   End If
   .TextMatrix(I, 1) = CmbVendorQuo.Text
   .TextMatrix(I, 2) = CmbTreatmentTypeQuo.Text
   .TextMatrix(I, 3) = TxtRateperMTQuo.Text
   .TextMatrix(I, 4) = TxtRemarkQuo.Text
   .TextMatrix(I, 5) = mVendorQuoID
   .TextMatrix(I, 6) = mTreatmentTypeQuoID
End With
Call ClearQuotationDetail
End Sub

Private Sub CmdAddToListTre_Click()
If Trim(CmbTreatmentType.Text) = "" Then
   MsgBox "Blank Treatment Type not allowed!!!"
   CmbTreatmentType.SetFocus
   Exit Sub
End If
If Trim(TxtRate.Text) = "" Then
   MsgBox "Blank Rate per MT not allowed!!!"
   TxtRate.SetFocus
   Exit Sub
End If

With MSHFlexGrid2
   For I = 1 To .Rows - 1
      If CmdAddToListTre.Caption = "Update List" Then
         If I <> .RowSel Then
            If .TextMatrix(I, 1) = CmbTreatmentType.Text Then
               MsgBox "Duplicate Treatment Type not Allowed!!!"
               Exit Sub
            End If
         End If
      Else
         If .TextMatrix(I, 1) = CmbTreatmentType.Text Then
            MsgBox "Duplicate Treatment Type not Allowed!!!"
            Exit Sub
         End If
      End If
   Next

   If CmdAddToListTre.Caption = "Update List" Then
      I = .RowSel
   Else
      I = .Rows - 1
      .Rows = .Rows + 1
   End If
   .TextMatrix(I, 1) = CmbTreatmentType.Text
   .TextMatrix(I, 2) = TxtRate.Text
   .TextMatrix(I, 3) = mTreatmentTypeID
End With
Call ClearTreatmentDetail
End Sub

Private Sub CmdDeleteFromListQuo_Click()
If MSHFlexGrid3.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbVendor.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If

With MSHFlexGrid3
   I = .RowSel
   If I > 0 Then
      For RO = I To .Rows - 2
         For C = 0 To .Cols - 1
            .TextMatrix(RO, C) = .TextMatrix(RO + 1, C)
         Next
      Next
      .Rows = .Rows - 1
   End If
End With

Call ClearQuotationDetail
End Sub

Private Sub CmdDeleteFromListTre_Click()
If MSHFlexGrid2.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbTreatmentType.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If

With MSHFlexGrid2
   I = .RowSel
   If I > 0 Then
      For RO = I To .Rows - 2
         For C = 0 To .Cols - 1
            .TextMatrix(RO, C) = .TextMatrix(RO + 1, C)
         Next
      Next
      .Rows = .Rows - 1
   End If
End With

Call ClearTreatmentDetail
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_FumigationAgreement.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_FumigationAgreement.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
Next
oWB.Save
MsgBox ("Excel file created!!!")
oXL.Visible = True
End Sub

Private Sub CmdSearch_Click()
Dim wc As Variant
If TxtFromDate.Value > Date Then
    MsgBox "Future From Date not Allowed!!!"
    TxtFromDate.SetFocus
    Exit Sub
End If
If TxtToDate.Value > Date Then
    MsgBox "Future To Date not Allowed!!!"
    TxtToDate.SetFocus
    Exit Sub
End If
If TxtFromDate.Value > TxtToDate.Value Then
    MsgBox "From Date should not greater than To Date!!!"
    TxtToDate.SetFocus
    Exit Sub
End If

Call Grid_Head

wc = " Where MFA.EntryDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' And '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
If CmbSType.Text <> "" Then
   wc = " And MFA.Type = '" & Left$(CmbSType.Text, 1) & "'"
End If
If CmbSLocation.Text <> "" Then
   wc = wc & " And MFA.LocationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = wc & " And MFA." & Gen_WcLocation
End If
If CmbSVendor.Text <> "" Then
   wc = wc & " And MFA.VendorID = " & mVendorID
End If
If CmbSStatus.Text <> "" Then
   wc = wc & " And MFA.Status = '" & Left$(CmbSStatus.Text, 1) & "'"
End If
If TxtSAgreeID.Text <> "" Then
   wc = wc & " And MFA.AgreementID  = " & Val(TxtSAgreeID.Text)
End If

tmp = " Select MFA.AgreementID, MFA.EntryDAte, MFA.StartDate, MFA.ExpiryDate, MFA.Type, MFA.Status, ML.LocationName, MV.VendorName"
tmp = tmp & " ,MFA.Remarks, MFA.CreatedBy, MFA.CreatedDate, MFA.UpdatedBy, MFA.UpdatedDate, MFA.LocationID, MFA.VendorID"
tmp = tmp & " from Mst_FumigationAgreement MFA"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MFA.LocationID"
tmp = tmp & " Inner Join Mst_Vendor MV on MV.VendorID = MFA.VendorID"
tmp = tmp & wc & " Order By MFA.AgreementID"
Call TmpTable

With MSHFlexGrid1
   .Rows = .Rows + TmpRs.RecordCount
   For iu = 1 To TmpRs.RecordCount
      .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!AgreementID), "", TmpRs!AgreementID)
      .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!EntryDate), "", Format(TmpRs!EntryDate, "dd-MMM-yyyy"))
      .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
      .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!ExpiryDate), "", Format(TmpRs!ExpiryDate, "dd-MMM-yyyy"))
      .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!Type), "", TmpRs!Type)
      .TextMatrix(iu, 5) = IIf(IsNull(TmpRs!Status), "", TmpRs!Status)
      .TextMatrix(iu, 6) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
      .TextMatrix(iu, 7) = IIf(IsNull(TmpRs!VendorName), "", TmpRs!VendorName)
      .TextMatrix(iu, 8) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)

      If LCase(.TextMatrix(iu, 4)) = "f" Then
         .TextMatrix(iu, 4) = "Fumigation"
      ElseIf LCase(.TextMatrix(iu, 4)) = "s" Then
         .TextMatrix(iu, 4) = "Spraying"
      End If
      If LCase(.TextMatrix(iu, 5)) = "p" Then
         .TextMatrix(iu, 5) = "Pending"
      ElseIf LCase(.TextMatrix(iu, 5)) = "a" Then
         .TextMatrix(iu, 5) = "Approved"
      ElseIf LCase(.TextMatrix(iu, 5)) = "r" Then
         .TextMatrix(iu, 5) = "Rejected"
      End If
      TmpRs.MoveNext
   Next
End With
End Sub

Private Sub CmdLocation_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdTreatmentType_Click()
FrmMst_TreatmentType.Show
End Sub

Private Sub CmdVendor_Click()
FrmMst_Vendor.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
If LCase(TxtStatus.Text) <> "pending" Then
   MsgBox "Fumigation Agreement is " & TxtStatus.Text & ", You can not edit!!!"
   Exit Sub
End If
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = "Delete from Mst_FumigationTreatmentDetail where AgreementID = " & Val(TxtAgreementID.Text)
   Call TmpTable
   
   tmp = "Delete from Mst_FumigationQuotationDetail where AgreementID = " & Val(TxtAgreementID.Text)
   Call TmpTable
   
   RsMst_FumigationAgreement.Find "AgreementID = " & Val(TxtAgreementID.Text)
   If RsMst_FumigationAgreement.EOF Then Exit Sub
   RsMst_FumigationAgreement.Delete
   RsMst_FumigationAgreement.Update
   If RsMst_FumigationAgreement.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_FumigationAgreement.MoveNext
   If RsMst_FumigationAgreement.EOF() Then
      RsMst_FumigationAgreement.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_FumigationAgreement.CancelUpdate
End Sub

Private Sub EditCmd_Click()

'If Lcase(Gen_UserRole) <> "power" And Lcase(Gen_UserRole) <> "superadmin"

'If LCase(TxtStatus.Text) <> "pending" Then
'   MsgBox "Fumigation Agreement is " & TxtStatus.Text & ", You can not edit!!!"
'   Exit Sub
'End If

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
   If LCase(Gen_UserRole) <> "admin" Then
      If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
         MsgBox "This agreement is " & TxtStatus & ". You can't edit this agreement any more!!!"
         Exit Sub
      End If
   Else
      If LCase(TxtStatus) = "approved" Or LCase(TxtStatus) = "rejected" Then
         Edit_Flg = "E"
         Call GButtonLock(Me, False)
         lblAttachment.Enabled = True
         TxtAttachmentRemarks.Locked = False
         CmdAttachment.Enabled = True
         CmdGetFile.Enabled = True
         CmdRemoveAttachment.Enabled = True
         Exit Sub
      End If
   End If
End If

Edit_Flg = "E"
Call GButtonLock(Me, False)

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

TxtRemark.Locked = False
TxtStartDate.Enabled = True
TxtExpiryDate.Enabled = True
TxtStatus.Visible = False
CmbStatus.Visible = True

lblAttachment.Enabled = True
TxtAttachmentRemarks.Locked = False
CmdAttachment.Enabled = True
CmdGetFile.Enabled = True
CmdRemoveAttachment.Enabled = True

End Sub

Private Sub ExitCmd_Click()
Unload Me
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

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I < 1 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub

RsMst_FumigationAgreement.MoveFirst
RsMst_FumigationAgreement.Find " AgreementID= " & Val(MSHFlexGrid1.TextMatrix(I, 0))

If Not RsMst_FumigationAgreement.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid2
   I = .RowSel
   If I < 1 Then Exit Sub
   If .TextMatrix(I, 1) = "" Then Exit Sub

   CmbTreatmentType.Text = .TextMatrix(I, 1)
   TxtRate.Text = .TextMatrix(I, 2)
   mTreatmentTypeID = .TextMatrix(I, 3)
   
   CmdAddToListTre.Caption = "Update List"
   
   CmbTreatmentType.Enabled = False
End With
CmdDeleteFromListTre.Enabled = True
End Sub

Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid3
   I = .RowSel
   If I < 1 Then Exit Sub
   If .TextMatrix(I, 1) = "" Then Exit Sub

   CmbVendorQuo.Text = .TextMatrix(I, 1)
   CmbTreatmentTypeQuo.Text = .TextMatrix(I, 2)
   TxtRateperMTQuo.Text = .TextMatrix(I, 3)
   TxtRemarkQuo.Text = .TextMatrix(I, 4)
   mVendorQuoID = .TextMatrix(I, 5)
   mTreatmentTypeQuoID = .TextMatrix(I, 6)
   
   CmdAddToListQuo.Caption = "Update List"
   
   CmbVendorQuo.Enabled = False
   CmbTreatmentTypeQuo.Enabled = False
End With
CmdDeleteFromListQuo.Enabled = True
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_FumigationAgreement.MoveNext
If RsMst_FumigationAgreement.EOF Then
   RsMst_FumigationAgreement.MoveLast
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
RsMst_FumigationAgreement.MovePrevious
If RsMst_FumigationAgreement.BOF Then
   RsMst_FumigationAgreement.MoveFirst
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
RsMst_FumigationAgreement.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsMst_FumigationAgreement.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub


Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsMst_FumigationAgreement.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 600
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width + 100
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2000

ProgressBar1.Width = MSHFlexGrid1.Width
ProgressBar1.Left = MSHFlexGrid1.Left
ProgressBar1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height
End Sub

Private Sub SaveCmd_Click()

If Frame1.Enabled = True Then

   If TxtStartDate.Value > Date Then
      MsgBox "Future Start Date not allowed!!!"
      TxtStartDate.SetFocus
      Exit Sub
   End If

   If TxtStartDate.Value > TxtExpiryDate.Value Then
      MsgBox "Start Date should not greater than expiry date!!!"
      TxtExpiryDate.SetFocus
      Exit Sub
   End If

   If Trim(CmbType.Text) = "" Then
      MsgBox "Blank Type not allowed!!!"
      CmbType.SetFocus
      Exit Sub
   End If

   If Trim(CmbStatus.Text) = "" Then
      MsgBox "Blank Status not allowed!!!"
      CmbStatus.SetFocus
      Exit Sub
   End If

   If Trim(CmbLocation.Text) = "" Then
      MsgBox "Blank Location not allowed!!!"
      CmbLocation.SetFocus
      Exit Sub
   End If

   If Trim(CmbVendor.Text) = "" Then
      MsgBox "Blank Vendor not allowed!!!"
      CmbVendor.SetFocus
      Exit Sub
   End If

   If MSHFlexGrid2.Rows = 2 Then
      MsgBox "Blank Treatment detail not allowed!!!"
      SSTab1.Tab = 0
      CmbTreatmentType.SetFocus
      Exit Sub
   End If
End If

'If MSHFlexGrid3.Rows = 2 Then
'   MsgBox "Blank Other Quotation detail not allowed!!!"
'   CmbVendor.SetFocus
'   SSTab1.Tab = 1
'   CmbVendorQuo.SetFocus
'   Exit Sub
'End If

If Edit_Flg = "A" Then
   If RsMst_FumigationAgreement.RecordCount > 0 Then
      tmp = " Select AgreementID, StartDate, ExpiryDate From Mst_FumigationAgreement Where "
      tmp = tmp & " LocationID = " & mLocationID & " And VendorID = " & mVendorID & " And ('" & Format(TxtStartDate.Value, "yyyy-mm-dd") & "' between StartDate and ExpiryDate or '" & Format(TxtExpiryDate.Value, "yyyy-mm-dd") & "' between StartDate and ExpiryDate)"
      tmp = tmp & " And Type = '" & Left$(CmbType.Text, 1) & "' And Status in ('A', 'P')"
      Call Tmp7Table
      If TmpRs7.RecordCount > 0 Then
         MsgBox "Agreement(" & IIf(IsNull(TmpRs7!AgreementID), "", TmpRs7!AgreementID) & ") already exist for period of " & IIf(IsNull(TmpRs7!StartDate), "", Format(TmpRs7!StartDate, "dd-MMM-yyyy")) & " And " & IIf(IsNull(TmpRs7!ExpiryDate), "", Format(TmpRs7!ExpiryDate, "dd-MMM-yyyy")) & "!!!"
         Exit Sub
      End If
   End If
   RsMst_FumigationAgreement.AddNew
   RsMst_FumigationAgreement!AgreementID = GetAgreementID
   RsMst_FumigationAgreement!G_UID = GetGUID
   TxtAgreementID.Text = RsMst_FumigationAgreement!AgreementID
Else
   tmp = " Select AgreementID, StartDate, ExpiryDate From Mst_FumigationAgreement Where "
   tmp = tmp & " LocationID = " & mLocationID & " And VendorID = " & mVendorID & " And ('" & Format(TxtStartDate.Value, "yyyy-mm-dd") & "' between StartDate and ExpiryDate or '" & Format(TxtExpiryDate.Value, "yyyy-mm-dd") & "' between StartDate and ExpiryDate)"
   tmp = tmp & " And Type = '" & Left$(CmbType.Text, 1) & "' And Status in ('A', 'P')"
   Call Tmp7Table
   If TmpRs7.RecordCount > 0 Then
      If Val(TxtAgreementID) <> TmpRs7!AgreementID Then
         MsgBox "Agreement(" & IIf(IsNull(TmpRs7!AgreementID), "", TmpRs7!AgreementID) & ") already exist for period of " & IIf(IsNull(TmpRs7!StartDate), "", Format(TmpRs7!StartDate, "dd-MMM-yyyy")) & " And " & IIf(IsNull(TmpRs7!ExpiryDate), "", Format(TmpRs7!ExpiryDate, "dd-MMM-yyyy")) & "!!!"
         Exit Sub
      End If
   End If
   
   RsMst_FumigationAgreement.MoveFirst
   RsMst_FumigationAgreement.Find " AgreementID= " & Val(TxtAgreementID.Text)

End If

RsMst_FumigationAgreement!EntryDate = TxtEntryDate.Value
RsMst_FumigationAgreement!StartDate = TxtStartDate.Value
RsMst_FumigationAgreement!ExpiryDate = TxtExpiryDate.Value
RsMst_FumigationAgreement!Type = Left$(CmbType.Text, 1)
RsMst_FumigationAgreement!Status = Left$(CmbStatus.Text, 1)
RsMst_FumigationAgreement!LocationID = mLocationID
RsMst_FumigationAgreement!VendorID = mVendorID
RsMst_FumigationAgreement!Remarks = TxtRemark.Text
If IsNull(RsMst_FumigationAgreement!CreatedBy) = True Or RsMst_FumigationAgreement!CreatedBy = "" Then
   RsMst_FumigationAgreement!CreatedBy = Gen_UserLoginID
   RsMst_FumigationAgreement!CreatedDate = Now
Else
   RsMst_FumigationAgreement!UpdatedBy = Gen_UserLoginID
   RsMst_FumigationAgreement!UpdatedDate = Now
End If
RsMst_FumigationAgreement.Update

Call SaveTreatmentDetail
Call SaveQuotationDetail
Call SaveAttachmentDetails

RsMst_FumigationAgreement.Requery
RsMst_FumigationAgreement.MoveFirst
RsMst_FumigationAgreement.Find " AgreementID= " & Val(TxtAgreementID.Text)

Call Indata

End Sub

Private Sub SaveTreatmentDetail()
Dim mDetailIDs As String
mDetailIDs = "(0 "
For I = 1 To MSHFlexGrid2.Rows - 2
   If Val(MSHFlexGrid2.TextMatrix(I, 0)) > 0 Then
      mDetailIDs = mDetailIDs & " ," & Val(MSHFlexGrid2.TextMatrix(I, 0))
   End If
Next
mDetailIDs = mDetailIDs & ")"
tmp = " Delete from Mst_FumigationTreatmentDetail Where AgreementID = " & Val(TxtAgreementID.Text) & " and DetailID not in " & mDetailIDs
Call Tmp6Table

If MSHFlexGrid2.Rows > 2 Then
   Call TableMst_FumigationTreatmentDetail("Where AgreementID = " & Val(TxtAgreementID.Text))
   For I = 1 To MSHFlexGrid2.Rows - 2
      If MSHFlexGrid2.TextMatrix(I, 0) = "" Then
         RsMst_FumigationTreatmentDetail.AddNew
         RsMst_FumigationTreatmentDetail!DetailID = GetTreatmentDetailID
         RsMst_FumigationTreatmentDetail!CreatedBy = Gen_UserLoginID
         RsMst_FumigationTreatmentDetail!CreatedDate = Now
         RsMst_FumigationTreatmentDetail!G_UID = GetGUID
      Else
         RsMst_FumigationTreatmentDetail.MoveFirst
         RsMst_FumigationTreatmentDetail.Find "DetailID = " & Val(MSHFlexGrid2.TextMatrix(I, 0))
         RsMst_FumigationTreatmentDetail!UpdatedBy = Gen_UserLoginID
         RsMst_FumigationTreatmentDetail!UpdatedDate = Now
      End If
      RsMst_FumigationTreatmentDetail!AgreementID = Val(TxtAgreementID.Text)
      RsMst_FumigationTreatmentDetail!RateperMT = Val(MSHFlexGrid2.TextMatrix(I, 2))
      RsMst_FumigationTreatmentDetail!TreatmentTypeID = Val(MSHFlexGrid2.TextMatrix(I, 3))
      RsMst_FumigationTreatmentDetail.Update
   Next
End If
End Sub

Private Sub SaveQuotationDetail()
Dim mDetailIDs As String
mDetailIDs = "(0 "
For I = 1 To MSHFlexGrid3.Rows - 2
   If Val(MSHFlexGrid3.TextMatrix(I, 0)) > 0 Then
      mDetailIDs = mDetailIDs & " ," & Val(MSHFlexGrid3.TextMatrix(I, 0))
   End If
Next
mDetailIDs = mDetailIDs & ")"
tmp = " Delete from Mst_FumigationQuotationDetail Where AgreementID = " & Val(TxtAgreementID.Text) & " and DetailID not in " & mDetailIDs
Call Tmp6Table

If MSHFlexGrid3.Rows > 2 Then
   Call TableMst_FumigationQuotationDetail("Where AgreementID = " & Val(TxtAgreementID.Text))
   For I = 1 To MSHFlexGrid3.Rows - 2
      If MSHFlexGrid3.TextMatrix(I, 0) = "" Then
         RsMst_FumigationQuotationDetail.AddNew
         RsMst_FumigationQuotationDetail!DetailID = GetQuotationDetailID
         RsMst_FumigationQuotationDetail!CreatedBy = Gen_UserLoginID
         RsMst_FumigationQuotationDetail!CreatedDate = Now
         RsMst_FumigationQuotationDetail!G_UID = GetGUID
      Else
         RsMst_FumigationQuotationDetail.MoveFirst
         RsMst_FumigationQuotationDetail.Find "DetailID = " & Val(MSHFlexGrid3.TextMatrix(I, 0))
         RsMst_FumigationQuotationDetail!UpdatedBy = Gen_UserLoginID
         RsMst_FumigationQuotationDetail!UpdatedDate = Now
      End If
      RsMst_FumigationQuotationDetail!AgreementID = Val(TxtAgreementID.Text)
      RsMst_FumigationQuotationDetail!VendorID = Val(MSHFlexGrid3.TextMatrix(I, 5))
      RsMst_FumigationQuotationDetail!TreatmentTypeID = Val(MSHFlexGrid3.TextMatrix(I, 6))
      RsMst_FumigationQuotationDetail!RateperMT = Val(MSHFlexGrid3.TextMatrix(I, 3))
      RsMst_FumigationQuotationDetail!Remarks = MSHFlexGrid3.TextMatrix(I, 4)
      RsMst_FumigationQuotationDetail.Update
   Next
End If
End Sub

Private Sub Indata()
Edit_Flg = ""

Call GButtonLock(Me, True)
Call ComboVisible(False)

Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False


TxtRemark.Locked = True
TxtStartDate.Enabled = False
TxtExpiryDate.Enabled = False

tmp = " Select MFA.AgreementID, MFA.EntryDAte, MFA.StartDate, MFA.ExpiryDate, MFA.Type, MFA.Status, ML.LocationName, MV.VendorName"
tmp = tmp & " ,MFA.Remarks, MFA.CreatedBy, MFA.CreatedDate, MFA.UpdatedBy, MFA.UpdatedDate, MFA.LocationID, MFA.VendorID"
tmp = tmp & " from Mst_FumigationAgreement MFA"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MFA.LocationID"
tmp = tmp & " Inner Join Mst_Vendor MV on MV.VendorID = MFA.VendorID"
tmp = tmp & " Where MFA.AgreementID = " & RsMst_FumigationAgreement!AgreementID
Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
   TxtAgreementID.Text = IIf(IsNull(TmpRs7!AgreementID), "", TmpRs7!AgreementID)
   TxtEntryDate.Value = IIf(IsNull(TmpRs7!EntryDate), Date, TmpRs7!EntryDate)
   TxtStartDate.Value = IIf(IsNull(TmpRs7!StartDate), Date, TmpRs7!StartDate)
   TxtExpiryDate.Value = IIf(IsNull(TmpRs7!ExpiryDate), TxtStartDate.Value + 365, TmpRs7!ExpiryDate)
   TxtType.Text = IIf(IsNull(TmpRs7!Type), "", TmpRs7!Type)
   TxtLocation.Text = IIf(IsNull(TmpRs7!LocationName), "", TmpRs7!LocationName)
   TxtVendor.Text = IIf(IsNull(TmpRs7!VendorName), "", TmpRs7!VendorName)
   TxtStatus.Text = IIf(IsNull(TmpRs7!Status), "", TmpRs7!Status)
   If LCase(TxtType.Text) = "f" Then
      TxtType.Text = "Fumigation"
   ElseIf LCase(TxtType.Text) = "s" Then
      TxtType.Text = "Spraying"
   End If
   If LCase(TxtStatus.Text) = "p" Then
      TxtStatus.Text = "Pending"
   ElseIf LCase(TxtStatus.Text) = "a" Then
      TxtStatus.Text = "Approved"
   ElseIf LCase(TxtStatus.Text) = "r" Then
      TxtStatus.Text = "Rejected"
   End If
   TxtRemark.Text = IIf(IsNull(TmpRs7!Remarks), "", TmpRs7!Remarks)
   TxtCreated.Text = IIf(IsNull(TmpRs7!CreatedBy), "", TmpRs7!CreatedBy & " :- " & TmpRs7!CreatedDate)
   TxtUpdated.Text = IIf(IsNull(TmpRs7!UpdatedBy), "", TmpRs7!UpdatedBy & " :- " & TmpRs7!UpdatedDate)
   
   CmbType.Text = TxtType.Text
   CmbLocation.Text = TxtLocation.Text
   CmbVendor.Text = TxtVendor.Text
   CmbStatus.Text = TxtStatus.Text
   
   mLocationID = IIf(IsNull(TmpRs7!LocationID), "", TmpRs7!LocationID)
   mVendorID = IIf(IsNull(TmpRs7!VendorID), "", TmpRs7!VendorID)
End If
Call IndataTreatmentDetail
Call IndataQuotationDetail
Call IndataAttachmentDetails

End Sub

Private Sub IndataTreatmentDetail()
Call ClearTreatmentDetail
Call Grid_Head2

tmp = " Select MFT.DetailID, MFT.RateperMT, MTT.TreatmentType, MFT.TreatmentTypeID"
tmp = tmp & " From Mst_FumigationTreatmentDetail MFT"
tmp = tmp & " Inner Join Mst_TreatmentType MTT on MTT.TreatmentTypeID = MFT.TreatmentTypeID"
tmp = tmp & " Where MFT.AgreementID = " & Val(TxtAgreementID.Text)
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   With MSHFlexGrid2
      .Rows = .Rows + TmpRs7.RecordCount
      For I = 1 To TmpRs7.RecordCount
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs7!DetailID), "", TmpRs7!DetailID)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!TreatmentType), "", TmpRs7!TreatmentType)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs7!RateperMT), "", TmpRs7!RateperMT)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs7!TreatmentTypeID), "", TmpRs7!TreatmentTypeID)
         TmpRs7.MoveNext
      Next
   End With
End If
If LCase(CmbType.Text) = "fumigation" Then
   lblrate.Caption = "Rate per MT"
   MSHFlexGrid2.TextMatrix(0, 2) = "Rate per MT"
ElseIf LCase(CmbType.Text) = "spraying" Then
   lblrate.Caption = "Rate per Sq. Ft"
   MSHFlexGrid2.TextMatrix(0, 2) = "Rate per Sq. Ft"
End If

End Sub

Private Sub IndataQuotationDetail()
Call ClearQuotationDetail
Call Grid_Head3

tmp = " Select MFQ.DetailID, MV.VendorName, MTT.TreatmentType, MFQ.RateperMT, MFQ.Remarks, MFQ.TreatmentTypeID, MFQ.VendorID"
tmp = tmp & " From Mst_FumigationQuotationDetail MFQ"
tmp = tmp & " Inner Join Mst_TreatmentType MTT on MTT.TreatmentTypeID = MFQ.TreatmentTypeID"
tmp = tmp & " Inner Join Mst_Vendor MV on MV.VendorID = MFQ.VendorID"
tmp = tmp & " Where MFQ.AgreementID = " & Val(TxtAgreementID.Text)
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
   With MSHFlexGrid3
      .Rows = .Rows + TmpRs7.RecordCount
      For I = 1 To TmpRs7.RecordCount
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs7!DetailID), "", TmpRs7!DetailID)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!VendorName), "", TmpRs7!VendorName)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs7!TreatmentType), "", TmpRs7!TreatmentType)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs7!RateperMT), "", TmpRs7!RateperMT)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs7!Remarks), "", TmpRs7!Remarks)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs7!VendorID), "", TmpRs7!VendorID)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs7!TreatmentTypeID), "", TmpRs7!TreatmentTypeID)
         TmpRs7.MoveNext
      Next
   End With
End If
End Sub

Private Sub TxtRate_Lostfocus()
If Trim(TxtRate.Text) = "" Then Exit Sub
If IsNumeric(TxtRate.Text) = False Or Val(TxtRate.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtRate.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRateperMTQuo_LostFocus()
If Trim(TxtRateperMTQuo.Text) = "" Then Exit Sub
If IsNumeric(TxtRateperMTQuo.Text) = False Or Val(TxtRateperMTQuo.Text) < 0 Then
   MsgBox "Invalid numeric format"
   TxtRateperMTQuo.SetFocus
   Exit Sub
End If
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

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbSType.Text = ""
CmbSLocation.Text = ""
CmbSVendor.Text = ""
CmbSStatus.Text = ""
TxtSAgreeID.Text = ""

Call Grid_Head
Call Grid_Head2
Call Grid_Head3

Call TableMst_FumigationAgreement
If RsMst_FumigationAgreement.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
RsMst_FumigationAgreement.MoveLast
Call Indata
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
   .Clear
   .Cols = 9
   .Rows = 2
  
   .TextMatrix(0, 0) = "Agreement ID"
   .TextMatrix(0, 1) = "Entry Date"
   .TextMatrix(0, 2) = "Start Date "
   .TextMatrix(0, 3) = "Expiry Date"
   .TextMatrix(0, 4) = "Type"
   .TextMatrix(0, 5) = "Status"
   .TextMatrix(0, 6) = "Location"
   .TextMatrix(0, 7) = "Vendor"
   .TextMatrix(0, 8) = "Remarks"
   
   .ColWidth(0) = 1100
   .ColWidth(1) = 1200
   .ColWidth(2) = 1200
   .ColWidth(3) = 1200
   .ColWidth(4) = 1200
   .ColWidth(5) = 1200
   .ColWidth(6) = 2200
   .ColWidth(7) = 2200
   .ColWidth(8) = 3500
   
   .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 4

   .TextMatrix(0, 0) = "Detail ID"
   .TextMatrix(0, 1) = "Treatment Type"
   .TextMatrix(0, 2) = "Rate Per MT"
   .TextMatrix(0, 3) = "TreatmentID"
   
   .ColWidth(0) = 1200
   .ColWidth(1) = 4500
   .ColWidth(2) = 2000
   .ColWidth(3) = 0
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid3
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 7

   .TextMatrix(0, 0) = "Detail ID"
   .TextMatrix(0, 1) = "Vendor"
   .TextMatrix(0, 2) = "Treatment Type"
   .TextMatrix(0, 3) = "Rate Per MT"
   .TextMatrix(0, 4) = "Remarks"
   .TextMatrix(0, 5) = "VendorID"
   .TextMatrix(0, 6) = "TreatmentID"
      
   .ColWidth(0) = 1200
   .ColWidth(1) = 3000
   .ColWidth(2) = 3000
   .ColWidth(3) = 1500
   .ColWidth(4) = 4500
   .ColWidth(5) = 0
   .ColWidth(6) = 0
End With
End Sub


Private Sub ClearControl()
Call Grid_Head2
Call Grid_Head3

TxtAgreementID.Text = ""
TxtEntryDate.Value = Date
TxtStartDate.Value = Date
TxtExpiryDate.Value = Date + 365
TxtType.Text = ""
CmbType.Text = ""
CmbStatus.Text = ""
TxtStatus.Text = ""
CmbLocation.Text = ""
TxtLocation.Text = ""
CmbVendor.Text = ""
TxtVendor.Text = ""
TxtRemark.Text = ""
TxtCreated.Text = ""
TxtUpdated.Text = ""
End Sub

Private Sub ClearTreatmentDetail()
CmbTreatmentType.Text = ""
TxtRate.Text = ""
CmdAddToListTre.Caption = "Add to List"
CmdDeleteFromListTre.Enabled = False
CmbTreatmentType.Enabled = True
End Sub

Private Sub ClearQuotationDetail()
CmbVendorQuo.Text = ""
CmbTreatmentTypeQuo.Text = ""
TxtRateperMTQuo.Text = ""
TxtRemarkQuo.Text = ""
CmdAddToListQuo.Caption = "Add to List"
CmdDeleteFromListQuo.Enabled = False
CmbVendorQuo.Enabled = True
CmbTreatmentTypeQuo.Enabled = True
End Sub

Private Sub ComboVisible(var1 As Boolean)
CmbType.Visible = var1
CmbStatus.Visible = var1
CmbLocation.Visible = var1
CmbVendor.Visible = var1
CmdLocation.Visible = var1
CmdVendor.Visible = var1

If var1 = False Then
   TxtType.Visible = True
   TxtStatus.Visible = True
   TxtLocation.Visible = True
   TxtVendor.Visible = True
Else
   TxtType.Visible = False
   TxtStatus.Visible = False
   TxtLocation.Visible = False
   TxtVendor.Visible = False
End If
End Sub

Private Function GetAgreementID() As Double
Dim Retval As Double
tmp = "Select max(AgreementID) from Mst_FumigationAgreement"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetAgreementID = Retval
End Function

Private Function GetTreatmentDetailID() As Double
Dim Retval As Double
tmp = "Select max(DetailID) from Mst_FumigationTreatmentDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetTreatmentDetailID = Retval
End Function

Private Function GetQuotationDetailID() As Double
Dim Retval As Double
tmp = "Select max(DetailID) from Mst_FumigationQuotationDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetQuotationDetailID = Retval
End Function


Private Sub IndataAttachmentDetails()
Call Grid_Head4
Call ClearFrameAttachment

Call TableMst_FSAgreeAttachment(" Where FSAID = " & TxtAgreementID.Text & " ")

MSHFlexGrid5.Rows = RsMst_FSAgreeAttachment.RecordCount + 2
For I = 1 To RsMst_FSAgreeAttachment.RecordCount
'Fumigation And Spraying
    MSHFlexGrid5.TextMatrix(I, 0) = I
    MSHFlexGrid5.TextMatrix(I, 1) = RsMst_FSAgreeAttachment!FileName
    MSHFlexGrid5.TextMatrix(I, 2) = RsMst_FSAgreeAttachment!Remarks
    MSHFlexGrid5.TextMatrix(I, 3) = App.Path & "\Agreement Attachment\Fumigation And Spraying\" & MSHFlexGrid5.TextMatrix(I, 1)
    File_Path = MSHFlexGrid5.TextMatrix(I, 3) 'App.Path & "\Project Attachment\"
    Call LoadFileFromDB(File_Path, RsMst_FSAgreeAttachment, "File")
    MSHFlexGrid5.TextMatrix(I, 4) = RsMst_FSAgreeAttachment!TxnID
    RsMst_FSAgreeAttachment.MoveNext
Next

lblAttachment.Enabled = False
TxtAttachmentRemarks.Locked = True
CmdAttachment.Enabled = False
CmdGetFile.Enabled = False
CmdRemoveAttachment.Enabled = False
End Sub

Private Sub ClearFrameAttachment()

lblAttachment.Caption = "File Attachment"
CmdAttachment.Caption = "Add Attachment"
TxtAttachmentRemarks = ""

End Sub

Private Sub SaveAttachmentDetails()

If MSHFlexGrid5.Rows <= 1 Then Exit Sub


Call TableMst_FSAgreeAttachment(" Where FSAID = " & Val(TxtAgreementID) & "")

For I = 1 To MSHFlexGrid5.Rows - 1
    If MSHFlexGrid5.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid5.TextMatrix(I, 4) = "" Then
       RsMst_FSAgreeAttachment.AddNew
 
       RsMst_FSAgreeAttachment!TxnID = GetMaxAttachID
       RsMst_FSAgreeAttachment!FSAID = TxtAgreementID
       mFlag = SaveFileToDB(MSHFlexGrid5.TextMatrix(I, 3), RsMst_FSAgreeAttachment, "File")
       mFlag = GetFileFromPath(MSHFlexGrid5.TextMatrix(I, 3))
       RsMst_FSAgreeAttachment!FileName = "FSAgreement" & TxtAgreementID & "-" & mFlag
    Else
       If RsMst_FSAgreeAttachment.RecordCount > 0 Then
          RsMst_FSAgreeAttachment.MoveFirst
       End If
       RsMst_FSAgreeAttachment.Find "TxnID = " & MSHFlexGrid5.TextMatrix(I, 4) & ""
    End If
    RsMst_FSAgreeAttachment!Remarks = MSHFlexGrid5.TextMatrix(I, 2)
    RsMst_FSAgreeAttachment!UpdatedDate = Now
    RsMst_FSAgreeAttachment!UpdatedBy = Gen_UserLoginID
    RsMst_FSAgreeAttachment!G_UID = GetGUID
    If IsNull(RsMst_FSAgreeAttachment!CreatedBy) = True Or RsMst_FSAgreeAttachment!CreatedBy = "" Then
       RsMst_FSAgreeAttachment!CreatedBy = Gen_UserLoginID
       RsMst_FSAgreeAttachment!CreatedDate = Now
    Else
       RsMst_FSAgreeAttachment!UpdatedBy = Gen_UserLoginID
       RsMst_FSAgreeAttachment!UpdatedDate = Now
    End If
    RsMst_FSAgreeAttachment.Update
Next

End Sub


Private Sub CmdGetFile_Click()

With cdialog
     .InitDir = App.Path
     .DialogTitle = "Select File"
     .FilterIndex = 1
     .ShowOpen
End With

If cdialog.CancelError Then
   lblAttachment.Caption = ""
Else
   strFilename = cdialog.FileName
   If strFilename = "" Then Exit Sub
   lblAttachment.Refresh
   lblAttachment.Caption = strFilename 'GetFileFromPath(strFilename)
End If
End Sub

Private Sub CmdRemoveAttachment_Click()
I = MSHFlexGrid5.RowSel
If I > 0 Then
    If MSHFlexGrid5.TextMatrix(I, 4) = "" And MSHFlexGrid5.TextMatrix(I, 1) <> "" Then
        For RO = I To MSHFlexGrid5.Rows - 2
            For C = 0 To MSHFlexGrid5.Cols - 1
                MSHFlexGrid5.TextMatrix(RO, C) = MSHFlexGrid5.TextMatrix(RO + 1, C)
            Next
        Next
        MSHFlexGrid5.Rows = MSHFlexGrid5.Rows - 1
    End If
End If

Call ClearFrameAttachment
End Sub

Private Sub CmdAttachment_Click()
Dim r As Variant

If lblAttachment.Caption = "File Attachment" Then
   MsgBox "Blank Attachment not allowed!!!"
   Exit Sub
End If

If TxtAttachmentRemarks = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtAttachmentRemarks.SetFocus
   Exit Sub
End If

If CmdAttachment.Caption = "Update Attachment" Then
   r = MSHFlexGrid3.RowSel
Else
   r = MSHFlexGrid5.Rows - 1
   MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
   MSHFlexGrid5.TextMatrix(r, 0) = ""
End If

MSHFlexGrid5.TextMatrix(r, 1) = GetFileFromPath(lblAttachment.Caption)
MSHFlexGrid5.TextMatrix(r, 2) = TxtAttachmentRemarks
MSHFlexGrid5.TextMatrix(r, 3) = lblAttachment.Caption
Call ClearFrameAttachment

End Sub


Private Sub MSHFlexGrid5_Click()
Dim ans, r As Variant
Dim ID As Integer
r = MSHFlexGrid5.RowSel
If r < 1 Then Exit Sub
If MSHFlexGrid5.TextMatrix(r, 1) = "" Then Exit Sub

If SaveCmd.Enabled = False Then
'Fumigation And Spraying
   ans = MsgBox("Do You want to view this document ", vbYesNo)
   If ans = vbYes Then
      If Right(App.Path, 1) = "\" Then
         Shell "Explorer.exe " & App.Path & "Agreement Attachment\Fumigation And Spraying", vbMaximizedFocus
      Else
         Shell "Explorer.exe " & App.Path & "Agreement Attachment\Fumigation And Spraying\", vbMaximizedFocus
      End If
   End If
   Exit Sub
End If

lblAttachment.Caption = MSHFlexGrid5.TextMatrix(r, 3)
TxtAttachmentRemarks = MSHFlexGrid5.TextMatrix(r, 2)

CmdRemoveAttachment.Enabled = True
If MSHFlexGrid5.TextMatrix(r, 4) <> "" Then
   CmdRemoveAttachment.Enabled = False
End If

CmdAttachment.Caption = "Update Attachment"
End Sub

Private Function GetMaxAttachID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Mst_FSAgreeAttachment"
Call TmpTableSMAttachment
Retval = IIf(IsNull(TmpRsSM_Attachment.Fields(0)), 0, TmpRsSM_Attachment.Fields(0))
Retval = Retval + 1
GetMaxAttachID = Retval
End Function

