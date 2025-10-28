VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_QQSGodownMapping 
   Caption         =   "QQS Godown Mapping"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   9570
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   15060
      Begin VB.CheckBox Check2 
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
         TabIndex        =   54
         Top             =   3120
         Width           =   3375
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
         TabIndex        =   30
         Top             =   4200
         Visible         =   0   'False
         Width           =   14115
         Begin VB.CheckBox Check4 
            Caption         =   "Free QQS"
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
            Left            =   5760
            TabIndex        =   60
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox TxtSID 
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
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   43
            Top             =   1425
            Width           =   2490
         End
         Begin VB.ComboBox CmbSDesignation 
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
            ItemData        =   "FrmTxn_QQSGodownMapping.frx":0000
            Left            =   10545
            List            =   "FrmTxn_QQSGodownMapping.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   555
            Width           =   3525
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
            TabIndex        =   32
            Top             =   1425
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
            TabIndex        =   31
            Top             =   1425
            Width           =   1000
         End
         Begin VB.ComboBox CmbSCMPM 
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
            ItemData        =   "FrmTxn_QQSGodownMapping.frx":0004
            Left            =   6780
            List            =   "FrmTxn_QQSGodownMapping.frx":0006
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   555
            Width           =   3720
         End
         Begin VB.ComboBox CmbSQQS 
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
            ItemData        =   "FrmTxn_QQSGodownMapping.frx":0008
            Left            =   3045
            List            =   "FrmTxn_QQSGodownMapping.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   555
            Width           =   3735
         End
         Begin VB.ComboBox CmbSBaseLoc 
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
            ItemData        =   "FrmTxn_QQSGodownMapping.frx":000C
            Left            =   120
            List            =   "FrmTxn_QQSGodownMapping.frx":000E
            Sorted          =   -1  'True
            TabIndex        =   42
            Top             =   1440
            Width           =   2910
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   375
            Left            =   1590
            TabIndex        =   35
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
            Format          =   107282433
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   375
            Left            =   135
            TabIndex        =   33
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
            Format          =   107282433
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   34
            Top             =   1920
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "QQS Godown Mapping ID"
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
            Left            =   3210
            TabIndex        =   52
            Top             =   1080
            Width           =   2310
         End
         Begin VB.Label Label29 
            Caption         =   "StoreMan Role"
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
            Left            =   11587
            TabIndex        =   51
            Top             =   255
            Width           =   1440
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
            TabIndex        =   47
            Top             =   255
            Width           =   855
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
            TabIndex        =   46
            Top             =   255
            Width           =   975
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
            TabIndex        =   45
            Top             =   1440
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
            Left            =   14295
            TabIndex        =   44
            Top             =   1455
            Width           =   90
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "CMPM"
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
            Left            =   8340
            TabIndex        =   41
            Top             =   255
            Width           =   600
         End
         Begin VB.Label Label13 
            Caption         =   "QQS"
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
            TabIndex        =   38
            Top             =   255
            Width           =   585
         End
         Begin VB.Label Label14 
            Caption         =   "Base Location"
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
            Left            =   975
            TabIndex        =   36
            Top             =   1110
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2850
         Left            =   75
         TabIndex        =   14
         Top             =   210
         Width           =   14895
         Begin VB.CheckBox Check3 
            Caption         =   "Free QQS"
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
            Left            =   6480
            TabIndex        =   59
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox TxtStatus 
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
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1320
         End
         Begin VB.TextBox TxtHrDesignation 
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
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2280
         End
         Begin VB.TextBox TxtCmpm 
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
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   840
            Width           =   5520
         End
         Begin VB.TextBox TxtDesignation 
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
            Left            =   6165
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1312
            Width           =   2280
         End
         Begin VB.TextBox TxtBaseLocation 
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
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1312
            Width           =   2640
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2385
            Width           =   5880
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
            Left            =   8760
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2385
            Width           =   6000
         End
         Begin VB.TextBox TxtQQSGodownID 
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
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   280
            Width           =   2115
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
            Height          =   570
            Left            =   1800
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1755
            Width           =   12960
         End
         Begin MSComCtl2.DTPicker TxtDate 
            Height          =   375
            Left            =   4680
            TabIndex        =   18
            Top             =   285
            Width           =   1680
            _ExtentX        =   2963
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
            Format          =   107282433
            CurrentDate     =   39884
         End
         Begin VB.ComboBox CmbQQS 
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
            TabIndex        =   1
            Top             =   800
            Width           =   5880
         End
         Begin VB.TextBox TxtQQS 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   800
            Visible         =   0   'False
            Width           =   5880
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Show Already Mapped Godown"
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
            Left            =   10560
            TabIndex        =   53
            Top             =   360
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label Label6 
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
            Height          =   300
            Left            =   12720
            TabIndex        =   58
            Top             =   1335
            Width           =   840
         End
         Begin VB.Label Label5 
            Caption         =   "HR Designation"
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
            Left            =   8520
            TabIndex        =   55
            Top             =   1335
            Width           =   1800
         End
         Begin VB.Label Label3 
            Caption         =   "StoreMan Role"
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
            Left            =   4560
            TabIndex        =   28
            Top             =   1350
            Width           =   1680
         End
         Begin VB.Label Label2 
            Caption         =   "Base Location"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1305
            Width           =   1575
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
            TabIndex        =   26
            Top             =   2445
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
            Left            =   7800
            TabIndex        =   25
            Top             =   2445
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "QQS Godown Mapping ID"
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
            Left            =   120
            TabIndex        =   24
            Top             =   270
            Width           =   1440
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
            Left            =   4080
            TabIndex        =   23
            Top             =   317
            Width           =   720
         End
         Begin VB.Label Label26 
            Caption         =   "CMPM Name"
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
            Left            =   7800
            TabIndex        =   22
            Top             =   870
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "QQS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   21
            Top             =   885
            Width           =   1185
         End
         Begin VB.Label Label27 
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
            TabIndex        =   20
            Top             =   1920
            Width           =   1335
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1125
         Left            =   75
         TabIndex        =   12
         Top             =   8265
         Width           =   14925
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
            Left            =   120
            TabIndex        =   13
            Top             =   195
            Width           =   2940
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
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   2940
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
            Top             =   195
            Width           =   3060
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
            Top             =   600
            Width           =   3060
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
            Left            =   6120
            TabIndex        =   4
            Top             =   195
            Width           =   2940
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
            Left            =   6120
            TabIndex        =   9
            Top             =   600
            Width           =   2940
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
            Left            =   9060
            TabIndex        =   5
            Top             =   195
            Width           =   2940
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
            Left            =   9060
            TabIndex        =   10
            Top             =   600
            Width           =   2940
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
            Left            =   12000
            TabIndex        =   6
            Top             =   195
            Width           =   2820
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
            Left            =   12000
            TabIndex        =   11
            Top             =   600
            Width           =   2820
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4785
         Left            =   75
         TabIndex        =   29
         Top             =   3465
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   8440
         _Version        =   393216
         FixedCols       =   0
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
Attribute VB_Name = "FrmTxn_QQSGodownMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mEmployeeID, mDirectReportID, Edit_Flg, mFreeQQS As Variant
Dim mSEmployeeID, mSDirectReportID, mSLocationID, mSDesignationID As Variant



Private Sub AddCmd_Click()
Edit_Flg = "A"

Call GButtonLock(Me, False)
Call Grid_Head2

Frame1.Enabled = True

TxtQQSGodownID = ""
TxtDate.Value = Date

CmbQQS.Visible = True
TxtQQS.Visible = False
CmbQQS.Text = ""
mEmployeeID = ""
mDirectReportID = ""

TxtCmpm.Text = ""
TxtBaseLocation.Text = ""
TxtDesignation.Text = ""
TxtRemarks.Text = ""
TxtCreated = ""
TxtUpdated = ""
Check3.Value = 0
mDirectReportID = ""
mEmployeeID = ""

If RsTxn_QQSGodownMapping.RecordCount > 0 Then
   If CmbQQS.Enabled = True Then CmbQQS.SetFocus
End If

End Sub

Private Sub Check1_Click()
If Edit_Flg = "E" Then
   If CmbQQS.Text <> "" Then
      If TxtCmpm.Text <> "" Then
         Call GetGodownDetails(mDirectReportID)
      End If
   End If
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   For I = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
          MSHFlexGrid2.row = I
          MSHFlexGrid2.Col = 0
          MSHFlexGrid2.CellFontName = "Wingdings"
          MSHFlexGrid2.CellFontSize = 12
          MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid2.Text = strChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
          MSHFlexGrid2.row = I
          MSHFlexGrid2.Col = 0
          MSHFlexGrid2.CellFontName = "Wingdings"
          MSHFlexGrid2.CellFontSize = 12
          MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid2.Text = strUnChecked
       End If
   Next
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 0 Then
   Call IndataGodownMappingEdit
Else
   Call Grid_Head2
End If
End Sub

Private Sub CmbQQS_GotFocus()
tmp = CmbQQS.Text
Call TableMst_Employee("Where DesignationID=8 And Flag='Y' And EmployeeID not in (Select EmployeeID From Txn_QQSGodownMapping)")
CmbQQS.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbQQS.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbQQS.Text = tmp
End Sub

Private Sub CmbQQS_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbQQS.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbQQS, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbQQS.SetFocus
   Exit Sub
End If

mEmployeeID = RsMst_Employee!EmployeeID
mDirectReportID = RsMst_Employee!DirectReportingTo

Call GetCmpDetails(mDirectReportID, mEmployeeID)
If TxtCmpm.Text <> "" Then
   Call GetGodownDetails(mDirectReportID)
End If
End Sub



Private Sub CmbSBaseLoc_GotFocus()
tmp = CmbSBaseLoc.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbSBaseLoc.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSBaseLoc.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSBaseLoc.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSBaseLoc.Text = tmp
End Sub


Private Sub CmbSBaseLoc_LostFocus()
If CmbSBaseLoc.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSBaseLoc.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbSBaseLoc.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location.Fields("LocationID")
End Sub


Private Sub CmbSCMPM_GotFocus()
tmp = CmbSCMPM.Text
Call TableMst_Employee("Where EmployeeID in (Select DirectReportingTo From Mst_Employee Where DesignationID=8 And Flag='Y')")
CmbSCMPM.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSCMPM.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSCMPM.Text = tmp
End Sub


Private Sub CmbSCMPM_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSCMPM.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSCMPM, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSCMPM.SetFocus
   Exit Sub
End If

mSDirectReportID = RsMst_Employee!EmployeeID
End Sub
Private Sub CmbSDesignation_GotFocus()
tmp = CmbSDesignation.Text
Call TableMst_Designation("")
CmbSDesignation.Clear
If RsMst_Designation.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbSDesignation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Designation.RecordCount
    CmbSDesignation.AddItem RsMst_Designation!Designation
    RsMst_Designation.MoveNext
Next
CmbSDesignation.Text = tmp
End Sub

Private Sub CmbSDesignation_LostFocus()
If CmbSDesignation.Text = "" Then
   mSDesignationID = Null
   Exit Sub
End If

RsMst_Designation.MoveFirst
RsMst_Designation.Find "Designation = '" & CmbSDesignation.Text & "'"

If RsMst_Designation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbSDesignation.SetFocus
   Exit Sub
End If
mSDesignationID = RsMst_Designation!DesignationID
End Sub

Private Sub CmbSQQS_GotFocus()
tmp = CmbSQQS.Text
Call TableMst_Employee("Where DesignationID=8 And Flag='Y'")
CmbSQQS.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSQQS.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSQQS.Text = tmp
End Sub


Private Sub CmbSQQS_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSQQS.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSQQS, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSQQS.SetFocus
   Exit Sub
End If

mSEmployeeID = RsMst_Employee!EmployeeID
mSDirectReportID = RsMst_Employee!DirectReportingTo

End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "QQS Godown Mapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "QQS Godown Mapping.xls")
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
Private Sub CmdSearch_Click()
Dim Wc1 As Variant

Call Grid_Head1

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To Date'!!! "
   SFrom.SetFocus
   Exit Sub
End If

tmp = " Select TQQS.QQSMapID,TQQS.EntryDate,ME.EmployeeName+'~'+ME.EmployeeNo as FullName,ME.Flag,MEE.EmployeeName+'~'+MEE.EmployeeNo as FullName1,ML.LocationName,MD.Designation,TQQS.Remarks"
tmp = tmp & " From Txn_QQSGodownMapping TQQS "
tmp = tmp & " Inner Join Mst_Employee ME on TQQS.EmployeeID=ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee MEE on ME.DirectReportingTo=MEE.EmployeeID"
tmp = tmp & " Inner join Mst_Location ML on MEE.BaseLocationID=ML.LocationID"
tmp = tmp & " Inner join Mst_Designation MD on MEE.DesignationID=MD.DesignationID"

If Gen_DBType = "MDB" Then
   Wc1 = " Where TQQS.EntryDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "#"
Else
   Wc1 = " Where TQQS.EntryDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "'"
End If

If CmbSQQS.Text <> "" Then
   Wc1 = Wc1 & " And ME.EmployeeID=" & mSEmployeeID & ""
End If

If CmbSCMPM.Text <> "" Then
   Wc1 = Wc1 & " And MEE.EmployeeID=" & mSDirectReportID & ""
End If

If CmbSDesignation.Text <> "" Then
   Wc1 = Wc1 & " And MD.Designation='" & CmbSDesignation.Text & "'"
End If

If CmbSBaseLoc.Text <> "" Then
   Wc1 = Wc1 & " And ML.LocationName='" & CmbSBaseLoc.Text & "'"
End If

If TxtSID.Text <> "" Then
   Wc1 = Wc1 & " And TQQS.QQSMapID='" & TxtSID.Text & "'"
End If

If Check4.Value = 1 Then
   Wc1 = Wc1 & " And TQQS.FreeFlag='Y'"
End If

tmp = tmp & Wc1

tmp = tmp & " order By TQQS.QQSMapID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(MSHFlexGrid1.TextMatrix(I, 3) = "Y", "Active", "DeActive")
       TmpRs1.MoveNext
   Next
End If
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

   tmp = " Insert into Log_Txn_QQSGodownMappingDetail" & _
         " (ID," & _
         "  QQSMAPID," & _
         "  GodownID," & _
         "  UpdatedBy)" & _
         " ( Select Isnull((Select Max(ID) from Log_Txn_QQSGodownMappingDetail)+1,1) , QQSMAPID,GodownID,'" & Gen_UserLoginID & "'" & _
         " From Txn_QQSGodownMappingDetail Where QQSMapID= " & Val(TxtQQSGodownID.Text) & ")"

   Call Tmp1Table

   tmp = " Delete From Txn_QQSGodownMappingDetail where QQSMapID=" & TxtQQSGodownID.Text & ""
   Call Tmp4Table
   
   RsTxn_QQSGodownMapping.Delete
   RsTxn_QQSGodownMapping.Update
   If RsTxn_QQSGodownMapping.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_QQSGodownMapping.MoveNext
   If RsTxn_QQSGodownMapping.EOF() Then
      RsTxn_QQSGodownMapping.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_QQSGodownMapping.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmbQQS.Visible = True
TxtQQS.Visible = True
CmbQQS.Text = TxtQQS.Text

If Check3.Value = 0 Then
   Call IndataGodownMappingEdit
End If
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_QQSGodownMapping.MoveFirst
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

Call Grid_Head2

Call TableTxn_QQSGodownMapping(" Order by QQSMapID")
If RsTxn_QQSGodownMapping.RecordCount = 0 Then
    Call AddCmd_Click
    Exit Sub
End If

Call GButtonLock(Me, True)
RsTxn_QQSGodownMapping.MoveLast
Call Indata
End Sub

Public Sub Grid_Head2()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 12
MSHFlexGrid2.TextMatrix(0, 0) = ""
MSHFlexGrid2.TextMatrix(0, 1) = "CMP"
MSHFlexGrid2.TextMatrix(0, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(0, 3) = "Location"
MSHFlexGrid2.TextMatrix(0, 4) = "State"
MSHFlexGrid2.TextMatrix(0, 5) = "Capacity"
MSHFlexGrid2.TextMatrix(0, 6) = "Start Date"
MSHFlexGrid2.TextMatrix(0, 7) = "Agreement End Date"
MSHFlexGrid2.TextMatrix(0, 8) = "CMPID"
MSHFlexGrid2.TextMatrix(0, 9) = "GodownID"
MSHFlexGrid2.TextMatrix(0, 10) = "RowID"
MSHFlexGrid2.TextMatrix(0, 11) = "Unit Type"

MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 2500
MSHFlexGrid2.ColWidth(2) = 1500
MSHFlexGrid2.ColWidth(3) = 2000
MSHFlexGrid2.ColWidth(4) = 2000
MSHFlexGrid2.ColWidth(5) = 2000
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 1500
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(10) = 0
MSHFlexGrid2.ColWidth(11) = 1500

End Sub

Public Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8
MSHFlexGrid1.TextMatrix(0, 0) = "QQSMapID"
MSHFlexGrid1.TextMatrix(0, 1) = "Date"
MSHFlexGrid1.TextMatrix(0, 2) = "QQS"
MSHFlexGrid1.TextMatrix(0, 3) = "Status"
MSHFlexGrid1.TextMatrix(0, 4) = "CMPM"
MSHFlexGrid1.TextMatrix(0, 5) = "BaseLocation"
MSHFlexGrid1.TextMatrix(0, 6) = "Designation"
MSHFlexGrid1.TextMatrix(0, 7) = "Remarks"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 3000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 4000
End Sub
Public Sub GetCmpDetails(mDirRptID As Variant, mEmpID)
If mDirRptID <> "" Then
   tmp = " Select ME.EmployeeName+'~'+ME.EmployeeNo as FullName,ML.LocationName,MD.Designation,ME.EmployeeNo,ME.Flag "
   tmp = tmp & " From Mst_Employee ME"
   tmp = tmp & " Inner join Mst_Location ML on ME.BaseLocationID=ML.LocationID"
   tmp = tmp & " Inner join Mst_Designation MD on ME.DesignationID=MD.DesignationID"
   tmp = tmp & " where EmployeeID=" & mDirRptID & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      TxtCmpm.Text = TmpRs1!FullName
      TxtBaseLocation.Text = TmpRs1!LocationName
      TxtDesignation.Text = TmpRs1!Designation
      'TxtHrDesignation.Text = GetHRDesignationFromHRPro(TmpRs1!EmployeeNo)
      TxtStatus.Text = IIf(TmpRs1!Flag = "Y", "Active", "DeActive")
   End If
   tmp = " Select ME.EmployeeName+'~'+ME.EmployeeNo as FullName,ML.LocationName,MD.Designation,ME.EmployeeNo,ME.Flag "
   tmp = tmp & " From Mst_Employee ME"
   tmp = tmp & " Inner join Mst_Location ML on ME.BaseLocationID=ML.LocationID"
   tmp = tmp & " Inner join Mst_Designation MD on ME.DesignationID=MD.DesignationID"
   tmp = tmp & " where EmployeeID=" & mEmpID & ""
   Call Tmp1Table
   If TmpRs1.RecordCount > 0 Then
      TxtBaseLocation.Text = TmpRs1!LocationName
      TxtDesignation.Text = TmpRs1!Designation
 '     TxtHrDesignation.Text = GetHRDesignationFromHRPro(TmpRs1!EmployeeNo)
      TxtStatus.Text = IIf(TmpRs1!Flag = "Y", "Active", "DeActive")
   End If
End If
End Sub

Public Sub Indata()

Frame1.Enabled = False
Check2.Visible = True
Call Grid_Head2

CmbQQS.Visible = False
TxtQQS.Visible = True
CmbQQS.Text = ""

TxtCmpm.Text = ""
TxtBaseLocation = ""
TxtDesignation = ""
TxtRemarks = ""
TxtCreated = ""
TxtUpdated = ""
TxtMinimum = ""

tmp = " Select TQQS.QQSMapID,TQQS.EntryDate,TQQS.Remarks,ME.EmployeeName+'~'+ME.EmployeeNo as FullName,ME.DirectReportingTo,TQQS.EmployeeID,TQQS.FreeFlag"
tmp = tmp & " From Txn_QQSGodownMapping TQQS "
tmp = tmp & " Inner Join Mst_Employee ME on TQQS.EmployeeID=ME.EmployeeID"
tmp = tmp & " Where TQQS.QQSMapID=" & RsTxn_QQSGodownMapping!QQSMapID & ""
Call Tmp5Table
If TmpRs5.RecordCount > 0 Then
   TxtQQSGodownID = IIf(IsNull(TmpRs5!QQSMapID), "", TmpRs5!QQSMapID)
   TxtDate = IIf(IsNull(TmpRs5!EntryDate), "", TmpRs5!EntryDate)
   TxtQQS = IIf(IsNull(TmpRs5!FullName), "", TmpRs5!FullName)
   TxtRemarks = IIf(IsNull(TmpRs5!Remarks), 0, TmpRs5!Remarks)
   mEmployeeID = IIf(IsNull(TmpRs5!EmployeeID), 0, TmpRs5!EmployeeID)
   mDirectReportID = IIf(IsNull(TmpRs5!DirectReportingTo), 0, TmpRs5!DirectReportingTo)
   mFreeQQS = IIf(IsNull(TmpRs5!FreeFlag), 0, TmpRs5!FreeFlag)
   Call GetCmpDetails(mDirectReportID, mEmployeeID)
   TxtCreated = IIf(IsNull(RsTxn_QQSGodownMapping!CreatedBy), "", RsTxn_QQSGodownMapping!CreatedBy) & " :- " & IIf(IsNull(RsTxn_QQSGodownMapping!CreatedDate), "", RsTxn_QQSGodownMapping!CreatedDate)
   TxtUpdated = IIf(IsNull(RsTxn_QQSGodownMapping!UpdatedBy), "", RsTxn_QQSGodownMapping!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_QQSGodownMapping!UpdatedDate), "", RsTxn_QQSGodownMapping!UpdatedDate)
   If mFreeQQS = "Y" Then
      Check3.Value = 1
   Else
      Check3.Value = 0
   End If
End If
Call IndataGodownMapping

If RsTxn_QQSGodownMapping.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Public Sub GetGodownDetails(mDirRptID As Variant)
Call Grid_Head2

Dim I
'tmp = " SELECT DISTINCT STUFF( (SELECT ',' + CAST([CmpID] AS varchar) " & _
'      " From Mst_EmpCmpMapping where EmployeeID=" & mDirRptID & " " & _
'      " FOR XML PATH('')),1,1,'') as Numbers FROM Mst_EmpCmpMapping"
      
'Call Tmp2Table
'If TmpRs2.RecordCount > 0 Then
'   If Not IsNull(TmpRs2(0)) Then
      tmp = " Select MC.CMPName,MG.GodownNo,ML.LocationName,MS.StateName,MG.GodownCapacity,MG.StartDate,MG.AgreementEndDate,MC.CmpID,MG.GodownId,'' RowID,'' as tt,MG.UnitTypeID,MU.UnitType"
      tmp = tmp & " from Mst_Godown MG"
      tmp = tmp & " inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
      tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
      tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
      tmp = tmp & " inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
      tmp = tmp & " where MG.cmpid in (SELECT CmpID From Mst_EmpCmpMapping where EmployeeID=" & mDirRptID & ")  And MG.CloseDate is Null And MG.GodownID Not In "
      tmp = tmp & " ( Select isnull(TQGMD.GodownID,0) From Txn_QQSGodownMapping TQGM "
      tmp = tmp & " Inner Join Txn_QQSGodownMappingDetail TQGMD on TQGM.QQSMapID=TQGMD.QQSMapID"
      tmp = tmp & " Left Join Mst_Employee ME on TQGM.EmployeeID=ME.EmployeeID"
      tmp = tmp & " Where TQGM.EmployeeID<>" & mEmployeeID & " And (ME.Flag is null or ME.Flag ='Y'))"
      
'   End If
   Call Tmp3Table
   
   tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA "
   tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
   tmp = tmp & " Where TA.Flag not in ('R') "
   tmp = tmp & " Group BY TGA.GodownID "
      
   Call Tmp4Table
   For I = 1 To TmpRs3.RecordCount
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       MSHFlexGrid2.row = I
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellFontName = "Wingdings"
       MSHFlexGrid2.CellFontSize = 12
       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid2.Text = strUnChecked
         
       For C = 1 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs3.Fields(C - 1)), "", TmpRs3.Fields(C - 1))
       Next
         
       MSHFlexGrid2.TextMatrix(I, 6) = Format(MSHFlexGrid2.TextMatrix(I, 6), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(I, 7) = Format(MSHFlexGrid2.TextMatrix(I, 7), "dd-MMM-yyyy")
       
       If TmpRs3!UnitTypeID = 1 Then
          If TmpRs4.RecordCount > 0 Then
             TmpRs4.MoveFirst
             TmpRs4.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(I, 9)) & ""
             If TmpRs4.EOF = False Then
                MSHFlexGrid2.TextMatrix(I, 7) = Format(TmpRs4!ExpDate, "dd-mmm-yyyy")
             Else
                MSHFlexGrid2.TextMatrix(I, 7) = ""
             End If
          End If
       End If
       MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs3!UnitType), "", TmpRs3!UnitType)
       TmpRs3.MoveNext
   Next
'End If
End Sub


Private Sub LastCmd_Click()
RsTxn_QQSGodownMapping.MoveLast
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
RsTxn_QQSGodownMapping.MoveFirst
RsTxn_QQSGodownMapping.Find "QQSMapID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsTxn_QQSGodownMapping.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid2
    If .Col = 0 Then
      If .TextMatrix(.row, .Col) = "" Then Exit Sub
         .Col = 0
         If .TextMatrix(.row, .Col) = strUnChecked Then
            .TextMatrix(.row, .Col) = strChecked
         Else
         .TextMatrix(.row, .Col) = strUnChecked
      End If
   End If
End With
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_QQSGodownMapping.MoveNext
If RsTxn_QQSGodownMapping.EOF Then
   RsTxn_QQSGodownMapping.MoveLast
   LF_Flag = "Y"
End If
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_QQSGodownMapping.MovePrevious
If RsTxn_QQSGodownMapping.BOF Then
   RsTxn_QQSGodownMapping.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SaveCmd_Click()
Dim Str As Variant

If CmbQQS.Text = "" Then
   MsgBox "Blank QQS Not Allowed !!! "
   CmbQQS.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows <= 2 Then
   MsgBox "Godown Details Missing!!!"
   MSHFlexGrid2.SetFocus
   Exit Sub
End If

If Check3.Value = 0 Then
   For I = 1 To MSHFlexGrid2.Rows - 1
       MSHFlexGrid2.row = I
       MSHFlexGrid2.Col = 0
       If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
          Str = Str & "Y"
       End If
   Next
   If InStr(1, Str, "Y", vbTextCompare) > 0 Then
   Else
      MsgBox "Select atleast one Godown"
      Exit Sub
   End If
End If

If Edit_Flg = "A" Then
   If RsTxn_QQSGodownMapping.RecordCount > 0 Then
      RsTxn_QQSGodownMapping.MoveFirst
      RsTxn_QQSGodownMapping.Find "EmployeeID= " & mEmployeeID & ""
      If Not RsTxn_QQSGodownMapping.EOF Then
         MsgBox "Duplicate QQS Not Allowed !!! "
         CmbQQS.SetFocus
         Exit Sub
      End If
   End If
      RsTxn_QQSGodownMapping.AddNew
      RsTxn_QQSGodownMapping!QQSMapID = GetMaxID("QQSMapID", "Txn_QQSGodownMapping")
      TxtQQSGodownID.Text = RsTxn_QQSGodownMapping!QQSMapID
      RsTxn_QQSGodownMapping!G_UID = GetGUID
Else
      RsTxn_QQSGodownMapping.MoveFirst
      RsTxn_QQSGodownMapping.Find "EmployeeID= " & mEmployeeID & ""
      If Not RsTxn_QQSGodownMapping.EOF Then
         If RsTxn_QQSGodownMapping!QQSMapID <> TxtQQSGodownID.Text Then
            MsgBox "Duplicate QQS Not Allowed !!! "
            CmbQQS.SetFocus
            Exit Sub
         End If
      End If
      RsTxn_QQSGodownMapping.MoveFirst
      RsTxn_QQSGodownMapping.Find " QQSMapID= " & TxtQQSGodownID
End If

RsTxn_QQSGodownMapping!QQSMapID = TxtQQSGodownID.Text
RsTxn_QQSGodownMapping!EntryDate = TxtDate
RsTxn_QQSGodownMapping!EmployeeID = mEmployeeID
RsTxn_QQSGodownMapping!Remarks = TxtRemarks

If Check3.Value = 1 Then
   RsTxn_QQSGodownMapping!FreeFlag = "Y"
Else
   RsTxn_QQSGodownMapping!FreeFlag = "N"
End If

If IsNull(RsTxn_QQSGodownMapping!CreatedBy) = True Or RsTxn_QQSGodownMapping!CreatedBy = "" Then
   RsTxn_QQSGodownMapping!CreatedBy = Gen_UserLoginID
   RsTxn_QQSGodownMapping!CreatedDate = Now
Else
   RsTxn_QQSGodownMapping!UpdatedBy = Gen_UserLoginID
   RsTxn_QQSGodownMapping!UpdatedDate = Now
End If

RsTxn_QQSGodownMapping.Update
Call SaveGodownMapping
Call Indata

End Sub

Public Sub SaveGodownMapping()
Dim mGodownID As Variant
mGodownID = ""

If Edit_Flg = "E" Then
   For I = 1 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
         If mGodownID = "" Then
            mGodownID = MSHFlexGrid2.TextMatrix(I, 9)
         Else
            mGodownID = mGodownID & "," & MSHFlexGrid2.TextMatrix(I, 9)
         End If
      End If
   Next
   If mGodownID <> "" Then
      tmp = " Insert into Log_Txn_QQSGodownMappingDetail" & _
            " (ID," & _
            "  QQSMAPID," & _
            "  GodownID," & _
            "  UpdatedBy)" & _
            " ( Select Isnull((Select Max(ID) from Log_Txn_QQSGodownMappingDetail)+1,1) , QQSMAPID,GodownID,'" & Gen_UserLoginID & "'" & _
            " From Txn_QQSGodownMappingDetail Where GodownId In (" & mGodownID & ") And QQSMapID= " & TxtQQSGodownID.Text & ")"
      Call Tmp1Table
    End If
End If

tmp = " Delete from Txn_QQSGodownMappingDetail where QQSMapID = " & TxtQQSGodownID.Text & " "
Call Tmp4Table

If Check3.Value = 0 Then
   Call TableTxn_QQSGodownMappingDetail("Where QQSMapID = " & TxtQQSGodownID)
   If MSHFlexGrid2.Rows > 2 Then
      For I = 1 To MSHFlexGrid2.Rows - 1
          If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
             If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
                RsTxn_QQSGodownMappingDetail.AddNew
                RsTxn_QQSGodownMappingDetail!QQSMapDID = GetMaxID("QQSMapDID", "Txn_QQSGodownMappingDetail")
                RsTxn_QQSGodownMappingDetail!G_UID = GetGUID
            
                RsTxn_QQSGodownMappingDetail!QQSMapID = TxtQQSGodownID
                RsTxn_QQSGodownMappingDetail!GodownID = MSHFlexGrid2.TextMatrix(I, 9)
             
                If IsNull(RsTxn_QQSGodownMappingDetail!CreatedBy) = True Or RsTxn_QQSGodownMappingDetail!CreatedBy = "" Then
                   RsTxn_QQSGodownMappingDetail!CreatedBy = Gen_UserLoginID
                   RsTxn_QQSGodownMappingDetail!CreatedDate = Now
                Else
                   RsTxn_QQSGodownMappingDetail!UpdatedBy = Gen_UserLoginID
                   RsTxn_QQSGodownMappingDetail!UpdatedDate = Now
                End If
                RsTxn_QQSGodownMappingDetail.Update
                RsTxn_QQSGodownMappingDetail.MoveNext
       End If
  Next I
End If
End If

End Sub

Public Function GetDirectReportID(EmpID As Variant) As Variant
tmp = " Select DirectReportingTo from Mst_Employee Where EmployeeID=" & EmpID & ""
Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   GetDirectReportID = TmpRs3!DirectReportingTo
End If
End Function

Public Sub IndataGodownMapping()
Call Grid_Head2

tmp = " Select MC.CMPName,MG.GodownNo,ML.LocationName,MS.StateName,MG.GodownCapacity,MG.StartDate,MG.AgreementEndDate,MC.CmpID,MG.GodownId,TQMD.QQSMapDID RowID,MU.UnitType,MG.UnitTypeID"
tmp = tmp & " from Txn_QQSGodownMapping TQM"
tmp = tmp & " Inner Join Txn_QQSGodownMappingDetail TQMD on TQM.QQSMapID=TQMD.QQSMapID"
tmp = tmp & " Inner Join Mst_Godown MG on TQMD.GodownID=MG.GodownID"
tmp = tmp & " inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
tmp = tmp & " where TQM.QQSMapID =" & TxtQQSGodownID.Text & ""

Call TmpTable

tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA "
tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
tmp = tmp & " Where TA.Flag not in ('R') "
tmp = tmp & " Group BY TGA.GodownID "
      
Call Tmp4Table
   
If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 1 To TmpRs.RecordCount
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = 0
        MSHFlexGrid2.CellFontName = "Wingdings"
        MSHFlexGrid2.CellFontSize = 12
        MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
        MSHFlexGrid2.Text = strChecked
       
        For C = 1 To MSHFlexGrid2.Cols - 1
            MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C - 1)), "", TmpRs.Fields(C - 1))
        Next
       
        MSHFlexGrid2.TextMatrix(I, 6) = Format(MSHFlexGrid2.TextMatrix(I, 6), "dd-MMM-yyyy")
        MSHFlexGrid2.TextMatrix(I, 7) = Format(MSHFlexGrid2.TextMatrix(I, 7), "dd-MMM-yyyy")
       
       
        If TmpRs!UnitTypeID = 1 Then
           If TmpRs4.RecordCount > 0 Then
              TmpRs4.MoveFirst
              TmpRs4.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(I, 9)) & ""
              If TmpRs4.EOF = False Then
                 MSHFlexGrid2.TextMatrix(I, 7) = Format(TmpRs4!ExpDate, "dd-mmm-yyyy")
              Else
                 MSHFlexGrid2.TextMatrix(I, 7) = ""
              End If
           End If
        End If
       
       TmpRs.MoveNext
    Next
End If
End Sub

Public Sub IndataGodownMappingEdit()
If mDirectReportID <> "" Then

'   tmp = " SELECT DISTINCT STUFF( (SELECT ',' + CAST([CmpID] AS varchar) " & _
'         " From Mst_EmpCmpMapping where EmployeeID=" & mDirectReportID & " " & _
'         " FOR XML PATH('')),1,1,'') as Numbers FROM Mst_EmpCmpMapping"
          
'   Call Tmp2Table
'   If TmpRs2.RecordCount > 0 Then
'      If Not IsNull(TmpRs2(0)) Then
         Call Grid_Head2
         tmp = "Select Distinct CMPName,GodownNo,LocationName,StateName,GodownCapacity,StartDate,AgreementEndDate,CmpID,GodownId,RowID,tt,UnitTypeID,UnitType From "
         tmp = tmp & " (Select MC.CMPName,MG.GodownNo,ML.LocationName,MS.StateName,MG.GodownCapacity,MG.StartDate,MG.AgreementEndDate,MC.CmpID,MG.GodownId,TQMD.QQSMapDID RowID,'old' as tt,MG.UnitTypeID,MU.UnitType "
         tmp = tmp & " from Txn_QQSGodownMapping TQM"
         tmp = tmp & " Inner Join Txn_QQSGodownMappingDetail TQMD on TQM.QQSMapID=TQMD.QQSMapID"
         tmp = tmp & " Inner Join Mst_Godown MG on TQMD.GodownID=MG.GodownID"
         tmp = tmp & " inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
         tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
         tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
         tmp = tmp & " inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
         tmp = tmp & " where TQM.QQSMapID =" & TxtQQSGodownID.Text & " And MG.CloseDate is Null"
         tmp = tmp & " Union All"
         tmp = tmp & " Select MC.CMPName,MG.GodownNo,ML.LocationName,MS.StateName,MG.GodownCapacity,MG.StartDate,MG.AgreementEndDate,MC.CmpID,MG.GodownId,0 as RowID,'New' as tt,MG.UnitTypeID,MU.UnitType"
         tmp = tmp & " from Mst_Godown MG"
         tmp = tmp & " inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
         tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
         tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
         tmp = tmp & " inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
         tmp = tmp & " where MG.cmpid in (SELECT CmpID From Mst_EmpCmpMapping where EmployeeID=" & mDirectReportID & ") And MG.CloseDate is Null And MG.GodownID Not In "
         tmp = tmp & " ( Select isnull(TQGMD.GodownID,0) From Txn_QQSGodownMapping TQGM "
         tmp = tmp & " Inner Join Txn_QQSGodownMappingDetail TQGMD on TQGM.QQSMapID=TQGMD.QQSMapID"
         tmp = tmp & " Left Join Mst_Employee ME on TQGM.EmployeeID=ME.EmployeeID"
         tmp = tmp & " Where TQGM.EmployeeID<>" & mEmployeeID & " And (ME.Flag is null or ME.Flag ='Y')) And MG.GodownID Not In (Select Distinct GodownID From Txn_QQSGodownMappingDetail Where QQSMapID =" & TxtQQSGodownID.Text & " )) A"
    
         Call TmpTable
            
         tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA "
         tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
         tmp = tmp & " Where TA.Flag not in ('R') "
         tmp = tmp & " Group BY TGA.GodownID "
         Call Tmp4Table
         
         If TmpRs.RecordCount > 0 Then
            MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
            For I = 1 To TmpRs.RecordCount
                MSHFlexGrid2.row = I
                MSHFlexGrid2.Col = 0
                MSHFlexGrid2.CellFontName = "Wingdings"
                MSHFlexGrid2.CellFontSize = 12
                MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
                  
                If LCase(TmpRs!TT) = "old" Then
                   MSHFlexGrid2.Text = strChecked
                Else
                   MSHFlexGrid2.Text = strUnChecked
                End If
                    
                For C = 1 To MSHFlexGrid2.Cols - 1
                    MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C - 1)), "", TmpRs.Fields(C - 1))
                    If LCase(TmpRs!TT) = "old" Then
                       MSHFlexGrid2.row = I
                       MSHFlexGrid2.Col = C
                       MSHFlexGrid2.CellBackColor = vbYellow
                    End If
                Next
                   
                MSHFlexGrid2.TextMatrix(I, 6) = Format(MSHFlexGrid2.TextMatrix(I, 6), "dd-MMM-yyyy")
                MSHFlexGrid2.TextMatrix(I, 7) = Format(MSHFlexGrid2.TextMatrix(I, 7), "dd-MMM-yyyy")
                
                If TmpRs!UnitTypeID = 1 Then
                   If TmpRs4.RecordCount > 0 Then
                      TmpRs4.MoveFirst
                      TmpRs4.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(I, 9)) & ""
                      If TmpRs4.EOF = False Then
                         MSHFlexGrid2.TextMatrix(I, 7) = Format(TmpRs4!ExpDate, "dd-mmm-yyyy")
                      Else
                         MSHFlexGrid2.TextMatrix(I, 7) = ""
                      End If
                   End If
                End If
                MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
                TmpRs.MoveNext
              Next
          End If
'        End If
'    End If
End If
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_QQSGodownMapping.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_QQSGodownMapping.MoveFirst
   RsTxn_QQSGodownMapping.Find "QQSMapID = " & Val(TxtQQSGodownID) & ""
   If RsTxn_QQSGodownMapping.EOF Then
      RsTxn_QQSGodownMapping.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Check2.Visible = False

Frame0.Height = Frame4.Height '- 1000
Frame0.Left = Frame4.Left - 100
Frame0.Top = Frame4.Top
Frame0.Width = Me.Width - 400
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 2100
SFrom = Date - 30
STo = Date
RsTxn_QQSGodownMapping.MoveFirst
RsTxn_QQSGodownMapping.Filter = ""

End Sub


