VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_TRF 
   Caption         =   "TRF"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9975
      Left            =   0
      ScaleHeight     =   9945
      ScaleWidth      =   15105
      TabIndex        =   41
      Top             =   0
      Width           =   15135
      Begin VB.PictureBox Frame0 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3705
         ScaleWidth      =   14865
         TabIndex        =   42
         Top             =   5040
         Visible         =   0   'False
         Width           =   14895
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
            ItemData        =   "FrmTxn_TRF.frx":0000
            Left            =   8760
            List            =   "FrmTxn_TRF.frx":0010
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   450
            Width           =   6015
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
            ItemData        =   "FrmTxn_TRF.frx":0039
            Left            =   6840
            List            =   "FrmTxn_TRF.frx":0049
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   960
            Width           =   7935
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
            ItemData        =   "FrmTxn_TRF.frx":0072
            Left            =   1485
            List            =   "FrmTxn_TRF.frx":0082
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   960
            Width           =   4065
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
            ItemData        =   "FrmTxn_TRF.frx":00AB
            Left            =   5655
            List            =   "FrmTxn_TRF.frx":00BB
            Sorted          =   -1  'True
            TabIndex        =   34
            Top             =   450
            Width           =   3030
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
            ItemData        =   "FrmTxn_TRF.frx":00E4
            Left            =   2895
            List            =   "FrmTxn_TRF.frx":00F4
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   450
            Width           =   2670
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
            TabIndex        =   38
            Top             =   1380
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
            Height          =   360
            Left            =   13725
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1380
            Width           =   1000
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   43
            Top             =   1800
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
            TabIndex        =   31
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
            Format          =   109051905
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker TxtSTo 
            Height          =   360
            Left            =   1485
            TabIndex        =   32
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
            Format          =   109051905
            CurrentDate     =   36494
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   3360
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
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
            Left            =   11250
            TabIndex        =   68
            Top             =   120
            Width           =   1035
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
            Left            =   5640
            TabIndex        =   67
            Top             =   1020
            Width           =   1095
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
            TabIndex        =   66
            Top             =   1020
            Width           =   1020
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
            Left            =   6788
            TabIndex        =   65
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   3518
            TabIndex        =   64
            Top             =   120
            Width           =   1425
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
            Left            =   11730
            TabIndex        =   47
            Top             =   1440
            Width           =   165
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
            TabIndex        =   46
            Top             =   120
            Width           =   945
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
            TabIndex        =   45
            Top             =   120
            Width           =   735
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
            Left            =   9120
            TabIndex        =   44
            Top             =   1440
            Width           =   2415
         End
      End
      Begin VB.PictureBox ButtonFrm 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   120
         ScaleHeight     =   1005
         ScaleWidth      =   14775
         TabIndex        =   55
         Top             =   8880
         Width           =   14805
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
            TabIndex        =   30
            Top             =   525
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   120
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
            TabIndex        =   21
            Top             =   120
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   525
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
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   525
            Width           =   2835
         End
      End
      Begin VB.PictureBox Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3675
         Left            =   120
         ScaleHeight     =   3645
         ScaleWidth      =   14850
         TabIndex        =   48
         Top             =   60
         Width           =   14880
         Begin VB.TextBox TxtCMPRepresnatative 
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
            Left            =   1755
            MaxLength       =   50
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   1923
            Width           =   12930
         End
         Begin VB.TextBox TxtSamplerName 
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
            Left            =   1755
            MaxLength       =   50
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1476
            Width           =   12930
         End
         Begin VB.TextBox TxtTRFID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   135
            Width           =   2070
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
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   3240
            Width           =   4710
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
            Left            =   8055
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   3240
            Width           =   6630
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
            Height          =   780
            Left            =   1755
            MaxLength       =   5000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   2370
            Width           =   12945
         End
         Begin VB.TextBox TxtTRFNo 
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
            Left            =   5175
            MaxLength       =   50
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   135
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker TxtSamplingDate 
            Height          =   360
            Left            =   8070
            TabIndex        =   3
            Top             =   135
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
            Format          =   109051905
            CurrentDate     =   39884
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
            ItemData        =   "FrmTxn_TRF.frx":011D
            Left            =   1755
            List            =   "FrmTxn_TRF.frx":012D
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   582
            Width           =   2070
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
            Left            =   1755
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   582
            Width           =   2070
         End
         Begin VB.ComboBox CmbExchangeType 
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
            ItemData        =   "FrmTxn_TRF.frx":0156
            Left            =   11295
            List            =   "FrmTxn_TRF.frx":0166
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   135
            Width           =   3390
         End
         Begin VB.ComboBox CmbClientName 
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
            ItemData        =   "FrmTxn_TRF.frx":018F
            Left            =   1755
            List            =   "FrmTxn_TRF.frx":019F
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   1029
            Width           =   12945
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
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1029
            Width           =   12930
         End
         Begin VB.TextBox TxtExchangeType 
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
            Left            =   11295
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   135
            Width           =   3390
         End
         Begin VB.ComboBox CmbCommodity 
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
            ItemData        =   "FrmTxn_TRF.frx":01C8
            Left            =   11295
            List            =   "FrmTxn_TRF.frx":01D8
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   582
            Width           =   3390
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
            ItemData        =   "FrmTxn_TRF.frx":0201
            Left            =   5175
            List            =   "FrmTxn_TRF.frx":0211
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   582
            Width           =   4230
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
            Left            =   11295
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   582
            Width           =   3390
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
            Left            =   5175
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   582
            Width           =   4230
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Exchange Type"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9480
            TabIndex        =   63
            Top             =   195
            Width           =   1650
         End
         Begin VB.Label Label8 
            Caption         =   "Name of CMP Representative"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   150
            TabIndex        =   62
            Top             =   1863
            Width           =   1515
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Sampler Name"
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
            TabIndex        =   61
            Top             =   1536
            Width           =   1395
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Sampling Date"
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
            Left            =   6600
            TabIndex        =   57
            Top             =   195
            Width           =   1335
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
            TabIndex        =   54
            Top             =   3300
            Width           =   855
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
            Left            =   6600
            TabIndex        =   53
            Top             =   3300
            Width           =   930
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
            TabIndex        =   51
            Top             =   2640
            Width           =   810
         End
         Begin VB.Label LbClaimID 
            AutoSize        =   -1  'True
            Caption         =   "TRF ID"
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
            TabIndex        =   50
            Top             =   195
            Width           =   645
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "TRF No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3960
            TabIndex        =   49
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CMP Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3960
            TabIndex        =   58
            Top             =   645
            Width           =   1155
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Client Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   60
            Top             =   1095
            Width           =   1320
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Commodity"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9480
            TabIndex        =   59
            Top             =   645
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   150
            TabIndex        =   52
            Top             =   642
            Width           =   930
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4995
         Left            =   120
         TabIndex        =   56
         Top             =   3795
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   8811
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   882
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "GSL Details"
         TabPicture(0)   =   "FrmTxn_TRF.frx":023A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Check1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.CheckBox Check1 
            Caption         =   "Select All"
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
            TabIndex        =   19
            Top             =   120
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   4350
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   7673
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   10320
      Visible         =   0   'False
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmTxn_TRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID, mLocationID, mCMPID, mCommodityID, mClientID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mSelRow As Variant
Dim mSelCol As Variant
Dim Edit_Flg As Variant
Dim tmp1 As String, mSM_Prefix As String
Dim mSExchangeTypeID, mSLocationID, mSCMPID, mSCommodityID, mSClientID As Variant

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)

Frame1.Enabled = True

Call ClearControl
Call ComboVisible(True)
Check1.Enabled = True
Check1.Value = vbUnchecked
'TxtStatus.Text = "Pending"
'CmbStatus.Text = TxtStatus.Text
'TxtStatus.Visible = True
'CmbStatus.Visible = False

If RsTxn_TRF.RecordCount > 0 Then
   TxtTRFNo.SetFocus
End If
End Sub

Private Sub Check1_Click()
If MSHFlexGrid2.Rows = 3 Then
   Check1.Value = vbUnchecked
   Exit Sub
End If
For I = 2 To MSHFlexGrid2.Rows - 2
   If Check1.Value = vbChecked Then
      MSHFlexGrid2.row = I
      MSHFlexGrid2.Col = 0
      MSHFlexGrid2.CellFontName = "Wingdings"
      MSHFlexGrid2.CellFontSize = 12
      MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid2.Text = strChecked
   Else
      MSHFlexGrid2.row = I
      MSHFlexGrid2.Col = 0
      MSHFlexGrid2.CellFontName = "Wingdings"
      MSHFlexGrid2.CellFontSize = 12
      MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid2.Text = strUnChecked
   End If
Next
End Sub

Private Sub CmbClientName_GotFocus()
If CmbExchangeType.Text = "" Then
   MsgBox "Select Exchange Type!!!"
   CmbExchangeType.SetFocus
   Exit Sub
End If
If CmbLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP Name!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
If CmbCommodity.Text = "" Then
   MsgBox "Select Commodity!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If

tmp1 = CmbClientName.Text
'tmp = " Select Distinct MCL.CLientIDRow, MCL.CLientName from Mst_Client MCL"
tmp = " Select Distinct MCL.ClientID, MCL.CLientName, MCL.ClientIDROw, MD.DepositoryName from Mst_Client MCL"
tmp = tmp & " inner join Mst_Depository MD on MD.DepositoryID=MCL.DepositoryID"
tmp = tmp & " Inner Join Mst_GSL MGSL on MGSL.ClientID = MCL.ClientID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
tmp = tmp & " Where mgsl.balanceweight > 0 And mgsl.balancebags > 0 "
tmp = tmp & " and mg.cmpid = " & mCMPID
tmp = tmp & " and mgsl.exchangetypeid = " & mExchangeTypeID
tmp = tmp & " and mgsl.commodityid = " & mCommodityID
tmp = tmp & " and mgsl.cdtfdate > '2013-03-31'"

Call Tmp5Table

CmbClientName.Clear
If TmpRs5.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To TmpRs5.RecordCount
   CmbClientName.AddItem TmpRs5!ClientName & "~" & TmpRs5!ClientIDRow & "~" & TmpRs5!DepositoryName & "~" & TmpRs5!ClientID
   TmpRs5.MoveNext
Next
CmbClientName.Text = tmp1
End Sub

Private Sub CmbClientName_LostFocus()
If CmbClientName.Text = "" Then
   Call Grid_Head2
   Exit Sub
End If
If TmpRs5.RecordCount <= 0 Then Exit Sub

strSplitString = Split(CmbClientName, "~")

If UBound(strSplitString) <> 3 Then
   MsgBox "Invalid Selection!!!"
   CmbClientName.SetFocus
   Exit Sub
End If

ClientName_ = Trim(strSplitString(0))
ClientNo_ = Trim(strSplitString(1))
DPName_ = Trim(strSplitString(2))
ClientIDRow_ = Trim(strSplitString(3))

TmpRs5.MoveFirst
TmpRs5.Find "ClientID = " & ClientIDRow_
If TmpRs5.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbClientName.SetFocus
   Exit Sub
End If
mClientID = TmpRs5!ClientID
If tmp1 <> CmbClientName.Text Then
   Call FillGSLDetails
End If
End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text
If Gen_WcCMP <> "" Then
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID & " And " & Gen_WcCMP)
Else
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID)
End If
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   Call Grid_Head2
   CmbClientName.Text = ""
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
If tmp <> CmbCMPName.Text Then
   CmbClientName.Text = ""
   Call FillGSLDetails
End If
End Sub

Private Sub CmbCommodity_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCommodity.Text
Call TableMst_Commodity(" Where CommodityID in (Select CommodityID from Mst_LocationCommodityMapping Where LocationID = " & mLocationID & ")")
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
   CmbCommodity.AddItem RsMst_Commodity!Commodity
   RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   Call Grid_Head2
   CmbClientName.Text = ""
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
If tmp <> CmbCommodity.Text Then
   CmbClientName.Text = ""
   Call FillGSLDetails
End If
End Sub


Private Sub CmbExchangeType_GotFocus()
tmp = CmbExchangeType.Text
Call TableMst_ExchangeType
CmbExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
   CmbExchangeType.AddItem RsMst_ExchangeType!ExchangeType
   RsMst_ExchangeType.MoveNext
Next
CmbExchangeType.Text = tmp
End Sub

Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   Call Grid_Head2
   CmbClientName.Text = ""
   Exit Sub
End If

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
If tmp <> "" Then
   Call Grid_Head2
   CmbClientName.Text = ""
End If
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where  " & Gen_WcLocation & "")
Else
   Call TableMst_Location
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
'----
Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Call Grid_Head2
   CmbCMPName.Text = ""
   CmbCommodity.Text = ""
   CmbClientName.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = IIf(IsNull(RsMst_Location!SM_Prefix), "", RsMst_Location!SM_Prefix)
If tmp <> CmbLocation.Text Then
   Call Grid_Head2
   CmbCMPName.Text = ""
   CmbCommodity.Text = ""
   CmbClientName.Text = ""
End If
End Sub

Private Sub CmbSClientName_GotFocus()
If CmbSExchangeType.Text = "" Then
   MsgBox "Select Exchange Type!!!"
   CmbSExchangeType.SetFocus
   Exit Sub
End If
If CmbSLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
If CmbSCMPName.Text = "" Then
   MsgBox "Select CMP Name!!!"
   CmbSCMPName.SetFocus
   Exit Sub
End If
If CmbSCommodity.Text = "" Then
   MsgBox "Select Commodity!!!"
   CmbSCommodity.SetFocus
   Exit Sub
End If

tmp1 = CmbSClientName.Text
'Call TableMst_Client("Where ClientIDRow in (Select ClientIDRow from Mst_GSL Where CommodityID = " & mCommodityID & " )ExchangeTypeID = " & mExchangeTypeID)
'tmp = " Select Distinct MCL.CLientIDRow, MCL.CLientName from Mst_Client MCL"
tmp = " Select Distinct MCL.ClientID, MCL.CLientName, MCL.ClientID, MD.DepositoryName from Mst_Client MCL"
tmp = tmp & " Inner Join Mst_GSL MGSL on MGSL.ClientID = MCL.ClientID"
tmp = tmp & " inner join Mst_Depository MD on MD.DepositoryID=MCL.DepositoryID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
tmp = tmp & " Where mgsl.balanceweight > 0 And mgsl.balancebags > 0 "
tmp = tmp & " and mg.cmpid = " & mSCMPID
tmp = tmp & " and mgsl.exchangetypeid = " & mSExchangeTypeID
tmp = tmp & " and mgsl.commodityid = " & mSCommodityID
tmp = tmp & " and mgsl.cdtfdate > '2013-03-31'"
Call Tmp2Table

CmbSClientName.Clear
If TmpRs2.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To TmpRs2.RecordCount
   'CmbSClientName.AddItem TmpRs2!Clientname
   CmbSClientName.AddItem TmpRs2!ClientName & "~" & TmpRs2!ClientIDRow & "~" & TmpRs2!DepositoryName & "~" & TmpRs2!ClientID
   TmpRs2.MoveNext
Next
CmbSClientName.Text = tmp1
End Sub

Private Sub CmbSClientName_LostFocus()
If CmbSClientName.Text = "" Then
   Call Grid_Head2
   Exit Sub
End If

strSplitString = Split(CmbSClientName, "~")

If UBound(strSplitString) <> 3 Then
   MsgBox "Invalid Selection!!!"
   CmbSClientName.SetFocus
   Exit Sub
End If

ClientName_ = Trim(strSplitString(0))
ClientNo_ = Trim(strSplitString(1))
DPName_ = Trim(strSplitString(2))
ClientIDRow_ = Trim(strSplitString(3))

TmpRs2.MoveFirst
TmpRs2.Find "ClientID = " & ClientIDRow_
If TmpRs2.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSClientName.SetFocus
   Exit Sub
End If
mSClientID = TmpRs2!ClientID
End Sub


Private Sub CmbSCMPName_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCMPName.Text
tmp1 = " "
If Gen_WcCMP <> "" Then
   tmp1 = " And " & Gen_WcCMP
End If
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mSLocationID & tmp1)

CmbSCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
   CmbSCMPName.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbSCMPName.Text = tmp
End Sub

Private Sub CmbSCMPName_LostFocus()
If CmbSCMPName.Text = "" Then
   CmbSClientName.Text = ""
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSCMPName.SetFocus
   Exit Sub
End If
mSCMPID = RsMst_CMP!CMPID
If CmbSCMPName.Text <> tmp Then
   CmbSClientName.Text = ""
End If
End Sub

Private Sub CmbSCommodity_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCommodity.Text
Call TableMst_Commodity(" Where CommodityID in (Select CommodityID from Mst_LocationCommodityMapping Where LocationID = " & mSLocationID & ")")
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbSClientName.Text = ""
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSCommodity.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_Commodity!CommodityID
If CmbSCommodity.Text <> tmp Then
   CmbSClientName.Text = ""
End If
End Sub


Private Sub CmbSExchangeType_GotFocus()
tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbSClientName.Text = ""
   Exit Sub
End If

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mSExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
If CmbSExchangeType.Text <> tmp Then
   CmbSClientName.Text = ""
End If
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where  " & Gen_WcLocation & "")
Else
   Call TableMst_Location
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
'----
Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   CmbSClientName.Text = ""
   CmbSCMPName.Text = ""
   CmbSCommodity.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
mSM_Prefix = IIf(IsNull(RsMst_Location!SM_Prefix), "", RsMst_Location!SM_Prefix)
If CmbSLocation.Text <> tmp Then
   CmbSClientName.Text = ""
   CmbSCMPName.Text = ""
   CmbSCommodity.Text = ""
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_TRF.xls")
Call Xls_Detail_Rpt
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
If TxtSFrom.Value > Date Then
    MsgBox "Future From Date not Allowed!!!"
    TxtSFrom.SetFocus
    Exit Sub
End If
If TxtSTo.Value > Date Then
    MsgBox "Future To Date not Allowed!!!"
    TxtSTo.SetFocus
    Exit Sub
End If
If TxtSFrom.Value > TxtSTo.Value Then
    MsgBox "From Date should not greater than To Date!!!"
    TxtSTo.SetFocus
    Exit Sub
End If

Call Grid_Head

wc = " Where TRF.SamplingDate between '" & Format(TxtSFrom.Value, "yyyy-mm-dd") & "' And '" & Format(TxtSTo.Value, "yyyy-mm-dd") & "'"
If CmbSExchangeType.Text <> "" Then
   wc = wc & " And TRF.ExchangeTypeID = " & mSExchangeTypeID
End If
If CmbSLocation.Text <> "" Then
   wc = wc & " And ML.LocationID = " & mSLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = wc & " And ML." & Gen_WcLocation
End If
If CmbSCMPName.Text <> "" Then
   wc = wc & " And TRF.CMPID = " & mSCMPID
ElseIf Gen_WcCMP <> "" Then
   wc = wc & " And TRF." & Gen_WcCMP
End If
If CmbSCommodity.Text <> "" Then
   wc = wc & " And TRF.CommodityID = " & mSCommodityID
End If
If CmbSClientName.Text <> "" Then
   wc = wc & " And TRF.ClientID = " & mSClientID
End If

tmp = " Select TRF.TRFID, TRF.TRFNo, TRF.SamplingDate, TRF.ExchangeTypeID, MET.ExchangeType, ML.LocationID, ML.LocationName, ML.SM_prefix"
tmp = tmp & " ,TRF.CMPID, MCMP.CMPName, TRF.CommodityID, MCom.Commodity, TRF.ClientID, MCL.CLientName, MCL.ClientIDRow, MD.DepositoryName,"
tmp = tmp & " TRF.SamplerName, TRF.CMPRepresentive, TRF.Remarks, TRF.G_UID ,TRF.Concurrency "
tmp = tmp & " ,TRF.CreatedBy ,TRF.CreatedDate ,TRF.UpdatedBy ,TRF.UpdatedDate"
tmp = tmp & " from Txn_TRF TRF"
tmp = tmp & " Left Join Mst_CMP MCMP on MCMP.CMPID = TRF.CMPID"
tmp = tmp & " Left Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Left Join Mst_ExchangeType MET on MET.ExchangeTypeID = TRF.ExchangeTypeID "
tmp = tmp & " Left Join Mst_Commodity MCom on MCom.CommodityID = TRF.CommodityID"
tmp = tmp & " Left Join Mst_Client MCL on MCL.ClientID = TRF.ClientID"
tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=MCL.DepositoryID"
tmp = tmp & wc & " Order By TRF.TRFID"

Call TmpTable

With MSHFlexGrid1
   .Rows = .Rows + TmpRs.RecordCount
   Label22.Caption = TmpRs.RecordCount
   ProgressBar1.Max = .Rows
   For iu = 1 To TmpRs.RecordCount
      .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!TRFID), "", TmpRs!TRFID)
      .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!TRFNo), "", TmpRs!TRFNo)
      .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!SamplingDate), "", Format(TmpRs!SamplingDate, "dd-MMM-yyyy"))
      .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
      .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
      .TextMatrix(iu, 5) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
      .TextMatrix(iu, 6) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
      .TextMatrix(iu, 7) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) & "~" & IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID) & "~" & IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName) & "~" & IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID) 'IIf(IsNull(TmpRs!Clientname), "", TmpRs!Clientname)
      .TextMatrix(iu, 8) = IIf(IsNull(TmpRs!SamplerName), "", TmpRs!SamplerName)
      .TextMatrix(iu, 9) = IIf(IsNull(TmpRs!CMPRepresentive), "", TmpRs!CMPRepresentive)
      .TextMatrix(iu, 10) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
      TmpRs.MoveNext
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End With
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
'If LCase(TxtStatus.Text) <> "pending" Then
'   MsgBox "SCM Payment Tracker is " & TxtStatus.Text & ", You can not edit!!!"
'   Exit Sub
'End If
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   tmp = "Delete from Txn_TRFDetail where TRFID = " & Val(TxtTRFID.Text)
   Call TmpTable
   
   RsTxn_TRF.Find "TRFID = " & Val(TxtTRFID.Text)
   If RsTxn_TRF.EOF Then Exit Sub
   RsTxn_TRF.Delete
   RsTxn_TRF.Update
   If RsTxn_TRF.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_TRF.MoveNext
   If RsTxn_TRF.EOF() Then
      RsTxn_TRF.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_TRF.CancelUpdate
End Sub


Private Sub EditCmd_Click()
'If LCase(TxtStatus.Text) <> "pending" Then
'   MsgBox "SCM Payment Tracker is " & TxtStatus.Text & ", You can not edit!!!"
'   Exit Sub
'End If
'Edit_Flg = "E"
'Call GButtonLock(Me, False)
'
'Frame1.Enabled = True
'Check1.Enabled = True
'Call ComboVisible(True)
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
TxtSFrom.Value = Date - 30
TxtSTo.Value = Date
CmbSClientName.Text = ""
CmbSCMPName.Text = ""
CmbSExchangeType.Text = ""
CmbSCommodity.Text = ""
CmbSLocation.Text = ""

Call Grid_Head
Call Grid_Head2

Call TableTxn_TRF
If RsTxn_TRF.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
RsTxn_TRF.MoveLast
Call Indata

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = strChecked Then
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = strUnChecked
ElseIf MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = strUnChecked Then
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = strChecked
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


Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I < 1 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub

RsTxn_TRF.MoveFirst
RsTxn_TRF.Find " TRFID= " & Val(MSHFlexGrid1.TextMatrix(I, 0))

If Not RsTxn_TRF.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_TRF.MoveNext
If RsTxn_TRF.EOF Then
   RsTxn_TRF.MoveLast
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
RsTxn_TRF.MovePrevious
If RsTxn_TRF.BOF Then
   RsTxn_TRF.MoveFirst
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
RsTxn_TRF.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsTxn_TRF.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SaveCmd_Click()
If Trim(TxtTRFNo.Text) = "" Then
   MsgBox "Blank TRF No not allowed!!!"
   TxtTRFNo.SetFocus
   Exit Sub
End If
If TxtSamplingDate.Value > Date Then
   MsgBox "Future Sampling Date not allowed!!!"
   TxtSamplingDate.SetFocus
   Exit Sub
End If
If Trim(CmbExchangeType.Text) = "" Then
   MsgBox "Blank Exchange Type not allowed!!!"
   CmbExchangeType.SetFocus
   Exit Sub
End If
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
If Trim(CmbCMPName.Text) = "" Then
   MsgBox "Blank CMP Name not allowed!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
If Trim(CmbCommodity.Text) = "" Then
   MsgBox "Blank Commodity not allowed!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
If Trim(CmbClientName.Text) = "" Then
   MsgBox "Blank Client Name not allowed!!!"
   CmbClientName.SetFocus
   Exit Sub
End If

If CheckBlankGSLDetail = False Then
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsTxn_TRF.RecordCount > 0 Then
      tmp = " Select TRFNo From Txn_TRF Where TRFNo = '" & Trim(TxtTRFNo.Text) & "' "
      Call Tmp5Table
      If TmpRs5.RecordCount > 0 Then
         MsgBox "Duplicate TRFNo not allowed!!!"
         TxtTRFNo.SetFocus
         Exit Sub
      End If
   End If
   RsTxn_TRF.AddNew
'   RsTxn_TRF!TRFID = GetTRFID
   RsTxn_TRF!G_UID = GetGUID
   mG_UID = RsTxn_TRF!G_UID
'   TxtTRFID.Text = RsTxn_TRF!TRFID
Else
   RsTxn_TRF.MoveFirst
   RsTxn_TRF.Find " TRFID = " & Val(TxtTRFID.Text)
End If

RsTxn_TRF!TRFNo = Trim(TxtTRFNo.Text)
RsTxn_TRF!SamplingDate = TxtSamplingDate.Value
RsTxn_TRF!ExchangeTypeID = mExchangeTypeID
'RsTxn_TRF!LocationID = mLocationID
RsTxn_TRF!SM_Prefix = mSM_Prefix
RsTxn_TRF!CMPID = mCMPID
RsTxn_TRF!CommodityID = mCommodityID
RsTxn_TRF!ClientID = mClientID
RsTxn_TRF!SamplerName = TxtSamplerName.Text
RsTxn_TRF!CMPRepresentive = TxtCMPRepresnatative.Text
RsTxn_TRF!Remarks = TxtRemarks.Text
If IsNull(RsTxn_TRF!CreatedBy) = True Or RsTxn_TRF!CreatedBy = "" Then
   RsTxn_TRF!CreatedBy = Gen_UserLoginID
   RsTxn_TRF!CreatedDate = Now
Else
   RsTxn_TRF!UpdatedBy = Gen_UserLoginID
   RsTxn_TRF!UpdatedDate = Now
End If
RsTxn_TRF.Update

tmp = "Select TRFID From Txn_TRF Where G_UID= '" & mG_UID & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   TxtTRFID.Text = TmpRs!TRFID
End If

Call SaveTRFDetail

RsTxn_TRF.Requery
RsTxn_TRF.MoveFirst
RsTxn_TRF.Find " TRFID = " & Val(TxtTRFID.Text)
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_TRF.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 200
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 2500

ProgressBar1.Width = MSHFlexGrid1.Width
ProgressBar1.Left = MSHFlexGrid1.Left
ProgressBar1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height
End Sub

Private Sub Grid_Head()
ProgressBar1.Value = 0
Label22.Caption = ""
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .FixedRows = 1
   .Cols = 11
   
   .TextMatrix(0, 0) = "TRF ID"
   .TextMatrix(0, 1) = "TRF No"
   .TextMatrix(0, 2) = "Sampling Date"
   .TextMatrix(0, 3) = "Exchange Type"
   .TextMatrix(0, 4) = "Location"
   .TextMatrix(0, 5) = "CMP Name"
   .TextMatrix(0, 6) = "Commodity"
   .TextMatrix(0, 7) = "Client Name"
   .TextMatrix(0, 8) = "Sampler Name"
   .TextMatrix(0, 9) = "Name of CMP Representative"
   .TextMatrix(0, 10) = "Remarks"
   
   .ColWidth(0) = 1100
   .ColWidth(1) = 1500
   .ColWidth(2) = 1500
   .ColWidth(3) = 1500
   .ColWidth(4) = 1800
   .ColWidth(5) = 2500
   .ColWidth(6) = 1900
   .ColWidth(7) = 3500
   .ColWidth(8) = 1700
   .ColWidth(9) = 1100
   .ColWidth(10) = 1100
   
   .RowHeight(1) = 300
End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 10
   
   .TextMatrix(1, 0) = "Select"
   .TextMatrix(1, 1) = "Godown no"
   .TextMatrix(1, 2) = "Stack No"
   .TextMatrix(1, 3) = "Lot No"
   .TextMatrix(1, 4) = "Balance Bags"
   .TextMatrix(1, 5) = "Balance Qty"
   .TextMatrix(1, 6) = "CDTF No"
   .TextMatrix(1, 7) = "CDTF Date"
   .TextMatrix(1, 8) = "GSLID"
   .TextMatrix(1, 9) = "DetailID"
      
   .ColWidth(0) = 1100
   .ColWidth(1) = 1500
   .ColWidth(2) = 1500
   .ColWidth(3) = 1200
   .ColWidth(4) = 1500
   .ColWidth(5) = 1500
   .ColWidth(6) = 1500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1500
   .ColWidth(9) = 0
      
   .RowHeight(1) = 600
End With
End Sub

Private Sub ClearControl()
Call Grid_Head2

TxtTRFID.Text = ""
TxtTRFNo.Text = ""
TxtSamplingDate.Value = Date
CmbExchangeType.Text = ""
TxtExchangeType.Text = ""
CmbLocation.Text = ""
TxtLocation.Text = ""
CmbClientName.Text = ""
TxtClientName.Text = ""
CmbCMPName.Text = ""
TxtCMPName.Text = ""
CmbCommodity.Text = ""
TxtCommodity.Text = ""
TxtSamplerName.Text = ""
TxtCMPRepresnatative.Text = ""
TxtRemarks.Text = ""
TxtCreated.Text = ""
TxtUpdated.Text = ""
End Sub

Private Sub ComboVisible(var1 As Boolean)
CmbExchangeType.Visible = var1
CmbLocation.Visible = var1
CmbCMPName.Visible = var1
CmbCommodity.Visible = var1
CmbClientName.Visible = var1
If var1 = False Then
   TxtExchangeType.Visible = True
   TxtLocation.Visible = True
   TxtCMPName.Visible = True
   TxtCommodity.Visible = True
   TxtClientName.Visible = True
Else
   TxtExchangeType.Visible = False
   TxtLocation.Visible = False
   TxtCMPName.Visible = False
   TxtCommodity.Visible = False
   TxtClientName.Visible = False
End If
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_TRF.xls")
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

Private Sub SetScrollBars()

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

Private Sub Indata()
Edit_Flg = ""
Check1.Enabled = False
Call GButtonLock(Me, True)
Call ComboVisible(False)
Frame1.Enabled = False
'Frame2.Enabled = False

tmp = " Select TRF.TRFID, TRF.TRFNo, TRF.SamplingDate, TRF.ExchangeTypeID, MET.ExchangeType, ML.LocationID, ML.LocationName, ML.SM_Prefix"
tmp = tmp & " ,TRF.CMPID, MCMP.CMPName, TRF.CommodityID, MCom.Commodity, TRF.ClientID, MCL.CLientName, MCL.ClientID, MCL.ClientIDRow, MD.DepositoryName,"
tmp = tmp & " TRF.SamplerName, TRF.CMPRepresentive, TRF.Remarks, TRF.G_UID ,TRF.Concurrency "
tmp = tmp & " ,TRF.CreatedBy ,TRF.CreatedDate ,TRF.UpdatedBy ,TRF.UpdatedDate"
tmp = tmp & " from Txn_TRF TRF"
tmp = tmp & " Left Join Mst_CMP MCMP on MCMP.CMPID = TRF.CMPID"
tmp = tmp & " Left Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Left Join Mst_ExchangeType MET on MET.ExchangeTypeID = TRF.ExchangeTypeID "
tmp = tmp & " Left Join Mst_Commodity MCom on MCom.CommodityID = TRF.CommodityID"
tmp = tmp & " Left Join Mst_Client MCL on MCL.ClientID = TRF.ClientID"
tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=MCL.DepositoryID"
tmp = tmp & " Where TRF.TRFID = " & RsTxn_TRF!TRFID
Call Tmp6Table

If TmpRs6.RecordCount > 0 Then
   TxtTRFID.Text = IIf(IsNull(TmpRs6!TRFID), "", TmpRs6!TRFID)
   TxtTRFNo.Text = IIf(IsNull(TmpRs6!TRFNo), "", TmpRs6!TRFNo)
   TxtSamplingDate.Value = IIf(IsNull(TmpRs6!SamplingDate), Date, Format(TmpRs6!SamplingDate, "dd-MMM-yyyy"))
   TxtExchangeType.Text = IIf(IsNull(TmpRs6!ExchangeType), "", TmpRs6!ExchangeType)
   TxtLocation.Text = IIf(IsNull(TmpRs6!LocationName), "", TmpRs6!LocationName)
   TxtCMPName.Text = IIf(IsNull(TmpRs6!CMPName), "", TmpRs6!CMPName)
   TxtCommodity.Text = IIf(IsNull(TmpRs6!Commodity), "", TmpRs6!Commodity)
   'TxtClientName.Text = IIf(IsNull(TmpRs6!Clientname), "", TmpRs6!Clientname)
   TxtClientName.Text = IIf(IsNull(TmpRs6!ClientName), "", TmpRs6!ClientName) & "~" & IIf(IsNull(TmpRs6!ClientIDRow), "", TmpRs6!ClientIDRow) & "~" & IIf(IsNull(TmpRs6!DepositoryName), "", TmpRs6!DepositoryName) & "~" & IIf(IsNull(TmpRs6!ClientID), "", TmpRs6!ClientID)
   TxtSamplerName.Text = IIf(IsNull(TmpRs6!SamplerName), "", TmpRs6!SamplerName)
   TxtCMPRepresnatative.Text = IIf(IsNull(TmpRs6!CMPRepresentive), "", TmpRs6!CMPRepresentive)
   TxtRemarks.Text = IIf(IsNull(TmpRs6!Remarks), "", TmpRs6!Remarks)
   TxtCreated.Text = IIf(IsNull(TmpRs6!CreatedBy), "", TmpRs6!CreatedBy & " :- " & TmpRs6!CreatedDate)
   TxtUpdated.Text = IIf(IsNull(TmpRs6!UpdatedBy), "", TmpRs6!UpdatedBy & " :- " & TmpRs6!UpdatedDate)
   
   CmbExchangeType.Text = TxtExchangeType.Text
   CmbLocation.Text = TxtLocation.Text
   CmbCMPName.Text = TxtCMPName.Text
   CmbCommodity.Text = TxtCommodity.Text
   CmbClientName.Text = TxtClientName.Text
   mExchangeTypeID = IIf(IsNull(TmpRs6!ExchangeTypeID), "", TmpRs6!ExchangeTypeID)
   mLocationID = IIf(IsNull(TmpRs6!LocationID), "", TmpRs6!LocationID)
   mSM_Prefix = IIf(IsNull(TmpRs6!SM_Prefix), "", TmpRs6!SM_Prefix)
   mCMPID = IIf(IsNull(TmpRs6!CMPID), "", TmpRs6!CMPID)
   mCommodityID = IIf(IsNull(TmpRs6!CommodityID), "", TmpRs6!CommodityID)
   mClientID = IIf(IsNull(TmpRs6!ClientID), "", TmpRs6!ClientID)
End If
Call FillGSLDetails
End Sub

Private Function GetTRFID() As Double
Dim Retval As Double
tmp = "Select max(TRFID) from Txn_TRF"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetTRFID = Retval
End Function

Private Function GetTRFDetailID() As Double
Dim Retval As Double
tmp = "Select max(DetailID) from Txn_TRFDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetTRFDetailID = Retval
End Function

Private Sub FillGSLDetails()
Call Grid_Head2
If CmbExchangeType.Text = "" Or CmbLocation.Text = "" Or CmbCMPName.Text = "" Or CmbCommodity.Text = "" Or CmbClientName.Text = "" Then Exit Sub
tmp = " select mgsl.Mst_GSL_ID, mg.godownno, mgsl.stackno, mgsl.lotno, mgsl.balancebags , mgsl.balanceweight ,mgsl.cdtfno, mgsl.cdtfdate "
If Edit_Flg <> "A" Then
   tmp = tmp & " , trd.detailid "
End If
tmp = tmp & " from mst_gsl mgsl"
tmp = tmp & " inner join  mst_godown mg on mg.godownid = mgsl.godownid"
If Edit_Flg <> "A" Then
   tmp = tmp & " inner join Txn_TRFdetail trd on trd.Mst_GSL_ID = mgsl.Mst_GSL_ID and trd.trfid = " & Val(TxtTRFID.Text)
End If
tmp = tmp & " Where mg.cmpid = " & mCMPID
tmp = tmp & " and mgsl.exchangetypeid = " & mExchangeTypeID
tmp = tmp & " and mgsl.commodityid = " & mCommodityID
tmp = tmp & " and mgsl.ClientID = " & mClientID
If Edit_Flg = "A" Then
   'tmp = tmp & " And mgsl.Flag = 'O' And GSLID not in (Select isnull(TRFD.GSLID,0) from Txn_TRFDetail TRFD Inner Join Txn_TRF TRF on TRF.TRFID = TRFD.TRFID Where TRF.TRFID <> " & Val(TxtTRFID.Text) & " And TRF.LocationID = " & mLocationID & ")"
   tmp = tmp & " And mgsl.Status = 'O' And Mst_GSL_ID not in (Select isnull(TRFD.Mst_GSL_ID,0) from Txn_TRFDetail TRFD Where TRFD.TRFID <> " & Val(TxtTRFID.Text) & " And TRFD.SM_Prefix = '" & mSM_Prefix & "')"
End If
tmp = tmp & " and mgsl.cdtfdate > '2013-03-31'"
tmp = tmp & " Order by mgsl.Mst_GSL_ID"
Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
   With MSHFlexGrid2
      .Rows = .Rows + TmpRs7.RecordCount
      For I = 2 To TmpRs7.RecordCount + 1
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs7!GodownNo), "", TmpRs7!GodownNo)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs7!StackNo), "", TmpRs7!StackNo)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs7!LotNo), "", TmpRs7!LotNo)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs7!BalanceBags), "", TmpRs7!BalanceBags)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs7!BalanceWeight), "", TmpRs7!BalanceWeight)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs7!CdtfNo), "", TmpRs7!CdtfNo)
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs7!CDTFDate), "", Format(TmpRs7!CDTFDate, "dd-MMM-yyyy"))
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs7!Mst_GSL_ID), "", TmpRs7!Mst_GSL_ID)
         If Edit_Flg <> "A" Then
            .TextMatrix(I, 9) = IIf(IsNull(TmpRs7!DetailID), "", TmpRs7!DetailID)
         End If
         .row = I
         .Col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 12
         .CellAlignment = flexAlignCenterCenter
         If Val(.TextMatrix(I, 9)) > 0 Then
            .Text = strChecked
         Else
            .Text = strUnChecked
         End If
         .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(I, 4))
         .TextMatrix(0, 5) = Val(.TextMatrix(0, 5)) + Val(.TextMatrix(I, 5))
         TmpRs7.MoveNext
      Next
   End With
End If
End Sub

Private Function CheckBlankGSLDetail() As Boolean
CheckBlankGSLDetail = False
If MSHFlexGrid2.Rows = 3 Then
   MsgBox "GSL Detail not found!!!"
   Exit Function
End If
For I = 2 To MSHFlexGrid2.Rows - 2
   If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
      CheckBlankGSLDetail = True
      Exit Function
   End If
Next
MsgBox "Select atleast one GSL from grid!!!"
End Function

Private Sub SaveTRFDetail()
Call TableTxn_TRFDetail(" Where TRFID = " & Val(TxtTRFID.Text))
With MSHFlexGrid2
   For I = 2 To .Rows - 2
      If .TextMatrix(I, 0) = strChecked Then
         If Val(.TextMatrix(I, 9)) = 0 Then
            RsTxn_TRFDetail.AddNew
'            RsTxn_TRFDetail!DetailID = GetTRFDetailID
            RsTxn_TRFDetail!G_UID = GetGUID
         Else
            RsTxn_TRFDetail.MoveFirst
            RsTxn_TRFDetail.Find " DetailID = " & Val(.TextMatrix(I, 9))
         End If
         RsTxn_TRFDetail!TRFID = Val(TxtTRFID.Text)
         RsTxn_TRFDetail!Mst_GSL_ID = Val(.TextMatrix(I, 8))
         RsTxn_TRFDetail!SM_Prefix = mSM_Prefix
         RsTxn_TRFDetail!BalanceBags = Val(.TextMatrix(I, 4))
         RsTxn_TRFDetail!BalanceWeight = Val(.TextMatrix(I, 5))
         RsTxn_TRFDetail.Update
      ElseIf .TextMatrix(I, 0) = strUnChecked And Val(.TextMatrix(I, 9)) > 0 Then
         RsTxn_TRFDetail.MoveFirst
         RsTxn_TRFDetail.Find " DetailID = " & Val(.TextMatrix(I, 9))
         RsTxn_TRFDetail.Delete
      End If
   Next
End With
End Sub
