VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_AuditDetails 
   Caption         =   "Audit Details"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   94
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   93
      Top             =   10560
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   10095
      Left            =   -15
      TabIndex        =   0
      Top             =   -105
      Width           =   17655
      Begin VB.Frame Frame0 
         Height          =   1830
         Left            =   165
         TabIndex        =   65
         Top             =   5220
         Visible         =   0   'False
         Width           =   14490
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
            ItemData        =   "Frm_Txn_AuditDetails.frx":0000
            Left            =   1095
            List            =   "Frm_Txn_AuditDetails.frx":000A
            TabIndex        =   72
            Top             =   780
            Width           =   2010
         End
         Begin VB.ComboBox CmbSAuditType 
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
            Left            =   11340
            Sorted          =   -1  'True
            TabIndex        =   71
            Top             =   360
            Width           =   3030
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
            Height          =   345
            Left            =   12165
            TabIndex        =   74
            Top             =   788
            Width           =   1125
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
            Left            =   11040
            TabIndex        =   73
            Top             =   788
            Width           =   1125
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
            Left            =   3165
            Sorted          =   -1  'True
            TabIndex        =   69
            Top             =   360
            Width           =   3480
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
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
            Left            =   13290
            TabIndex        =   75
            Top             =   788
            Width           =   1125
         End
         Begin VB.ComboBox CmbSAuditor 
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
            Left            =   6690
            Sorted          =   -1  'True
            TabIndex        =   70
            Top             =   360
            Width           =   4605
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   66
            Top             =   1260
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   330
            Left            =   120
            TabIndex        =   67
            Top             =   375
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            Format          =   78184449
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   330
            Left            =   1650
            TabIndex        =   68
            Top             =   375
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            Format          =   78184449
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin VB.Label Label30 
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
            Left            =   120
            TabIndex        =   100
            Top             =   833
            Width           =   855
         End
         Begin VB.Label Label29 
            Caption         =   "Auditor Type"
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
            Left            =   12270
            TabIndex        =   99
            Top             =   120
            Width           =   1170
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
            Left            =   5040
            TabIndex        =   81
            Top             =   840
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
            Left            =   7560
            TabIndex        =   80
            Top             =   840
            Width           =   90
         End
         Begin VB.Label Label19 
            Caption         =   "To Audit Date"
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
            Left            =   1800
            TabIndex        =   79
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "From Audit  Date"
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
            TabIndex        =   78
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label24 
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
            Height          =   255
            Left            =   4478
            TabIndex        =   77
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "Auditor's Name"
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
            Left            =   8265
            TabIndex        =   76
            Top             =   127
            Width           =   1455
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1035
         Left            =   120
         TabIndex        =   82
         Top             =   8625
         Width           =   14400
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   375
            Left            =   11340
            TabIndex        =   92
            Top             =   555
            Width           =   2775
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   375
            Left            =   8565
            TabIndex        =   91
            Top             =   555
            Width           =   2775
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   375
            Left            =   5670
            TabIndex        =   90
            Top             =   555
            Width           =   2895
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   375
            Left            =   2895
            TabIndex        =   89
            Top             =   555
            Width           =   2775
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   375
            Left            =   120
            TabIndex        =   88
            Top             =   555
            Width           =   2775
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "&Search"
            Height          =   375
            Left            =   11340
            TabIndex        =   87
            Top             =   180
            Width           =   2775
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   8565
            TabIndex        =   86
            Top             =   180
            Width           =   2775
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   5670
            TabIndex        =   85
            Top             =   180
            Width           =   2895
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   375
            Left            =   2895
            TabIndex        =   84
            Top             =   180
            Width           =   2775
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New "
            Height          =   375
            Left            =   120
            TabIndex        =   83
            Top             =   180
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Audit Point Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4965
         Left            =   120
         TabIndex        =   95
         Top             =   3675
         Width           =   14415
         Begin VB.Frame Frame3 
            Height          =   2805
            Left            =   120
            TabIndex        =   96
            Top             =   180
            Width           =   13335
            Begin VB.TextBox TxtValue 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   9300
               TabIndex        =   15
               Top             =   2025
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.TextBox TxtActual 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   4140
               TabIndex        =   13
               Top             =   2025
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.TextBox TxtRequired 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1620
               TabIndex        =   12
               Top             =   2025
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.CheckBox Check1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Check1"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4515
               TabIndex        =   16
               Top             =   2445
               Width           =   210
            End
            Begin VB.TextBox TxtComment 
               Appearance      =   0  'Flat
               Height          =   630
               Left            =   1620
               MaxLength       =   7950
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   1350
               Width           =   11115
            End
            Begin VB.CommandButton CmdAdd2List 
               Caption         =   "Add to List"
               Height          =   345
               Left            =   9795
               TabIndex        =   18
               Top             =   2400
               Width           =   1575
            End
            Begin VB.CommandButton CmdRemoveFrmList 
               Caption         =   "Remove from List"
               Height          =   345
               Left            =   11355
               TabIndex        =   52
               Top             =   2400
               Width           =   1455
            End
            Begin VB.TextBox TxtObservation 
               Appearance      =   0  'Flat
               Height          =   750
               Left            =   1620
               MaxLength       =   7950
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   555
               Width           =   11115
            End
            Begin VB.CommandButton CmdAuditPoint 
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
               Left            =   12780
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   165
               Width           =   375
            End
            Begin VB.ComboBox CmbAuditPoint 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1620
               Sorted          =   -1  'True
               TabIndex        =   9
               Top             =   195
               Width           =   11130
            End
            Begin VB.TextBox TxtAuditPointID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1200
               TabIndex        =   49
               Top             =   195
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox TxtOldAuditID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   7920
               TabIndex        =   48
               Top             =   2410
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox TxtAuditDetailsID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   9315
               TabIndex        =   47
               Top             =   2415
               Visible         =   0   'False
               Width           =   375
            End
            Begin MSComCtl2.DTPicker TxtStartDate 
               Height          =   330
               Left            =   1620
               TabIndex        =   50
               Top             =   2400
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   78184449
               CurrentDate     =   39850
            End
            Begin MSComCtl2.DTPicker TxtCloseDate 
               Height          =   330
               Left            =   4860
               TabIndex        =   17
               Top             =   2400
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
               _Version        =   393216
               Format          =   78184449
               CurrentDate     =   39850
            End
            Begin MSComCtl2.DTPicker TxtRequiredDate 
               Height          =   330
               Left            =   7140
               TabIndex        =   14
               Top             =   2010
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   582
               _Version        =   393216
               Format          =   78184449
               CurrentDate     =   39850
            End
            Begin VB.Label Label14 
               Caption         =   "Value"
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
               Left            =   8700
               TabIndex        =   63
               Top             =   2055
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label13 
               Caption         =   "Required Date"
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
               Left            =   5820
               TabIndex        =   62
               Top             =   2055
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.Label Label12 
               Caption         =   "Actual"
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
               Left            =   3420
               TabIndex        =   61
               Top             =   2055
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label11 
               Caption         =   "Required/As per book"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   60
               TabIndex        =   60
               Top             =   1995
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label8 
               Caption         =   "Closed Date"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3420
               TabIndex        =   59
               Top             =   2460
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Auditee's Reply"
               Height          =   255
               Left            =   60
               TabIndex        =   58
               Top             =   1605
               Width           =   1455
            End
            Begin VB.Label Label4 
               Caption         =   "Auditor's Obseravtion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   60
               TabIndex        =   57
               Top             =   705
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Audit Point"
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
               Left            =   60
               TabIndex        =   56
               Top             =   225
               Width           =   1455
            End
            Begin VB.Label Label15 
               Caption         =   "Started Date"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   60
               TabIndex        =   55
               Top             =   2460
               Width           =   1095
            End
            Begin VB.Label Label16 
               Caption         =   "OldAuditID"
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
               Left            =   6600
               TabIndex        =   54
               Top             =   2445
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label17 
               Caption         =   "AuditDetailsID"
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
               Left            =   8520
               TabIndex        =   53
               Top             =   2450
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   1785
            Left            =   120
            TabIndex        =   64
            Top             =   3030
            Width           =   13335
            _ExtentX        =   23521
            _ExtentY        =   3149
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2910
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   14400
         Begin VB.TextBox TxtGodAddress 
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
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   2055
            Width           =   11535
         End
         Begin VB.TextBox TxtAuditID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   180
            Width           =   3855
         End
         Begin VB.ComboBox CmbGodownName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   915
            Width           =   5715
         End
         Begin VB.ComboBox CmbLocation 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   540
            Width           =   5700
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
            Height          =   315
            Left            =   12840
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   540
            Width           =   375
         End
         Begin VB.CommandButton CmdGodown 
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
            Left            =   12840
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   915
            Width           =   375
         End
         Begin VB.ComboBox CmbAuditor 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7080
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1305
            Width           =   5715
         End
         Begin VB.CommandButton CmdAuditor 
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
            Left            =   12840
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   1305
            Width           =   375
         End
         Begin VB.CommandButton CmdAuditType 
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
            Left            =   12840
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1680
            Width           =   375
         End
         Begin VB.ComboBox CmbAuditType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   1680
            Width           =   11115
         End
         Begin VB.CommandButton CmdAuditSegment 
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
            Left            =   5640
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   540
            Width           =   375
         End
         Begin VB.ComboBox CmbAuditSegment 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   540
            Width           =   3870
         End
         Begin VB.ComboBox CmbDesignationID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   1305
            Width           =   3870
         End
         Begin VB.CommandButton CmdWarehouse 
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
            Left            =   5640
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   915
            Width           =   375
         End
         Begin VB.ComboBox CmbCMPID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   915
            Width           =   3870
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
            Left            =   8385
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2460
            Width           =   4815
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
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2460
            Width           =   5415
         End
         Begin MSComCtl2.DTPicker TxtAuditDate 
            Height          =   330
            Left            =   7080
            TabIndex        =   1
            Top             =   180
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   78184449
            CurrentDate     =   39850
         End
         Begin VB.TextBox TxtAuditSegment 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   540
            Width           =   3855
         End
         Begin VB.TextBox TxtCMPID 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   915
            Width           =   3855
         End
         Begin VB.TextBox TxtDesignation 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1305
            Width           =   3855
         End
         Begin VB.TextBox TxtLocationName 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   540
            Width           =   5700
         End
         Begin VB.TextBox TxtGodownName 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   915
            Width           =   5700
         End
         Begin VB.TextBox TxtAuditor 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1305
            Width           =   5700
         End
         Begin VB.TextBox TxtAuditType 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1680
            Width           =   11100
         End
         Begin VB.Label Label26 
            Caption         =   "Godown Address :"
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
            TabIndex        =   98
            Top             =   2115
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Audit ID"
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
            Left            =   135
            TabIndex        =   46
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Audit Date"
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
            Left            =   6120
            TabIndex        =   45
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label28 
            Caption         =   "Godown Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6120
            TabIndex        =   44
            Top             =   885
            Width           =   855
         End
         Begin VB.Label Label53 
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
            Height          =   255
            Left            =   6120
            TabIndex        =   43
            Top             =   570
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "Designation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   1365
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Auditor's Name"
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
            Left            =   6120
            TabIndex        =   41
            Top             =   1275
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Audit Type"
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
            TabIndex        =   40
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Audit Segment"
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
            Left            =   135
            TabIndex        =   39
            Top             =   570
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "CMP Name"
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
            TabIndex        =   38
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label7 
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
            TabIndex        =   37
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label20 
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
            TabIndex        =   36
            Top             =   2520
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   101
         Top             =   3015
         Width           =   14430
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
            Height          =   390
            Left            =   1695
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   102
            Top             =   195
            Width           =   11790
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Remarks  :"
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
            Left            =   165
            TabIndex        =   103
            Top             =   255
            Width           =   960
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_AuditDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mDesignationID, mSegmentTypeID, mAuditTypeID, mAuditPointID, mAuditID As Variant
Dim mRequired, mEmployeeID, mSearchSegmentType, mSAuditTypeID As Variant
Dim mLocationID As Double
Dim mGodownID As Double
Dim mCMPID As Double
Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
'mSegmentTypeID = Null
'mGodownID = 0
'mAuditTypeID = Null

 
If RsTxn_Audit.RecordCount = 0 Then
   mSegmentTypeID = Null
   mGodownID = 0
   mAuditTypeID = Null
End If

TxtAuditID = ""
TxtAuditDate.Value = Date
TxtAuditDate.Enabled = True
CmbAuditSegment.Visible = True
CmdAuditSegment.Visible = True
TxtAuditSegment.Visible = False
'CmbAuditSegment.Text = ""

CmbLocation.Visible = True
CmdLocation.Visible = True
TxtLocationName.Visible = False
'CmbLocation.Text = ""

CmbCMPID.Visible = True
CmdWarehouse.Visible = True
TxtCMPID.Visible = False
'CmbCMPID.Text = ""

CmbGodownName.Visible = True
CmdGodown.Visible = True
TxtGodownName.Visible = False
'CmbGodownName.Text = ""

CmbDesignationID.Visible = True
TxtDesignation.Visible = False
CmbDesignationID.Text = GetDesignaion(Gen_DesignationID)

CmbAuditor.Visible = True
TxtAuditor.Visible = False
CmdAuditor.Visible = True
CmbAuditor.Text = GetEmployeeNonName(Gen_EmployeeID)

If CmbAuditor.Text <> "" Then
   mEmployeeID = Gen_EmployeeID
End If

CmbAuditType.Visible = True
TxtAuditType.Visible = False
CmdAuditType.Visible = True

'CmbAuditType.Text = ""


TxtCreated = ""
TxtUpdated = ""
TxtRemarks = ""

Call Grid_Head1
Frame1.Enabled = True
Frame3.Enabled = True
Frame5.Enabled = True

Call ClearFrame

If RsTxn_Audit.RecordCount > 0 Then
   TxtAuditDate.SetFocus
End If

Call GetPreviousPoints(mSegmentTypeID, mGodownID, mAuditTypeID)

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 10

MSHFlexGrid1.TextMatrix(0, 0) = "Audit ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Audit Date"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "CMP"
MSHFlexGrid1.TextMatrix(0, 4) = "Godown"
MSHFlexGrid1.TextMatrix(0, 5) = "Designation"
MSHFlexGrid1.TextMatrix(0, 6) = "Auditor"
MSHFlexGrid1.TextMatrix(0, 7) = "Audit Type"
MSHFlexGrid1.TextMatrix(0, 8) = "Audit Point"
MSHFlexGrid1.TextMatrix(0, 9) = "Close Date"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 2800
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 3600
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 1100

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 500

End Sub
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 17

MSHFlexGrid2.TextMatrix(0, 0) = ""
MSHFlexGrid2.TextMatrix(0, 1) = "Audit Point"
MSHFlexGrid2.TextMatrix(0, 2) = "Previous Observation"
MSHFlexGrid2.TextMatrix(0, 3) = "Present Observation"
MSHFlexGrid2.TextMatrix(0, 4) = "Previous Reply"
MSHFlexGrid2.TextMatrix(0, 5) = "Present Reply"
MSHFlexGrid2.TextMatrix(0, 6) = "Required Type"
MSHFlexGrid2.TextMatrix(0, 7) = "Required"
MSHFlexGrid2.TextMatrix(0, 8) = "Actual"
MSHFlexGrid2.TextMatrix(0, 9) = "Required Date"
MSHFlexGrid2.TextMatrix(0, 10) = "Value"
MSHFlexGrid2.TextMatrix(0, 11) = "Start Date"
MSHFlexGrid2.TextMatrix(0, 12) = "Close Date"
MSHFlexGrid2.TextMatrix(0, 13) = "AuditPointID"
MSHFlexGrid2.TextMatrix(0, 14) = "AuditDetailsID"
MSHFlexGrid2.TextMatrix(0, 15) = "AuditID"
MSHFlexGrid2.TextMatrix(0, 16) = "UpdateOrNot"


MSHFlexGrid2.ColWidth(0) = 400
MSHFlexGrid2.ColWidth(1) = 3000
MSHFlexGrid2.ColWidth(2) = 4000
MSHFlexGrid2.ColWidth(3) = 4000
MSHFlexGrid2.ColWidth(4) = 4000
MSHFlexGrid2.ColWidth(5) = 4000
MSHFlexGrid2.ColWidth(6) = 0
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(10) = 0
MSHFlexGrid2.ColWidth(11) = 1600
MSHFlexGrid2.ColWidth(12) = 1600
MSHFlexGrid2.ColWidth(13) = 0
MSHFlexGrid2.ColWidth(14) = 0
MSHFlexGrid2.ColWidth(15) = 0
MSHFlexGrid2.ColWidth(16) = 0

End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
   TxtCloseDate.Visible = True
   TxtCloseDate.Value = Date
Else
   TxtCloseDate.Visible = False
End If
End Sub

Private Sub CmbAuditor_GotFocus()
If CmbDesignationID.Text = "" Then Exit Sub
tmp = CmbAuditor.Text
Call TableMst_Employee("Where DesignationID= " & mDesignationID & " And Flag = 'Y' ")
CmbAuditor.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee found"
   CmbOldEmployeeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbAuditor.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbAuditor.Text = tmp
End Sub

Private Sub CmbAuditor_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAuditor")
End If
End Sub

Private Sub CmbAuditor_LostFocus()
If CmbAuditor.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

Dim strSplitString() As String
Dim strEmpNo As String
Dim strEmpName As String
strSplitString = Split(CmbAuditor.Text, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid selection!!!"
   CmbAuditor.SetFocus
   Exit Sub
End If
strEmpNo = strSplitString(1)
strEmpName = strSplitString(0)

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & strEmpNo & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid  Selection "
   CmbAuditor.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
End Sub


Private Sub CmbAuditPoint_GotFocus()
tmp = CmbAuditPoint.Text
If CmbAuditSegment.Text = "" Then Exit Sub
If CmbAuditType.Text = "" Then Exit Sub
Call TableMst_AuditPoint("Where SegmentTypeID = " & mSegmentTypeID & " And AuditTypeID = " & mAuditTypeID & " ")
CmbAuditPoint.Clear
If RsMst_AuditPoint.RecordCount = 0 Then
   MsgBox "No Audit Points found"
   CmbAuditPoint.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditPoint.RecordCount
    CmbAuditPoint.AddItem RsMst_AuditPoint!AuditPoint
    RsMst_AuditPoint.MoveNext
Next
CmbAuditPoint.Text = tmp
End Sub

Private Sub CmbAuditPoint_LostFocus()
If CmbAuditPoint.Text = "" Then
   mAuditPointID = Null
   mRequired = Null
   Exit Sub
End If

RsMst_AuditPoint.MoveFirst
RsMst_AuditPoint.Find "AuditPoint = '" & CmbAuditPoint.Text & "'"

If RsMst_AuditPoint.EOF Then
   MsgBox "Invalid point Selection"
   CmbAuditPoint.SetFocus
   Exit Sub
End If
mAuditPointID = RsMst_AuditPoint!AuditPointID
TxtAuditPointID = RsMst_AuditPoint!AuditPointID

If RsMst_AuditPoint!RequiredType = "N" Then
   mRequired = "N"
   TxtValue = ""
   Label11.Visible = True 'Required
   Label12.Visible = True 'Actual
   Label13.Visible = False 'Date
   Label14.Visible = False 'Value
   TxtRequired.Visible = True
   TxtActual.Visible = True
   TxtRequiredDate.Visible = False
   TxtValue.Visible = False
ElseIf RsMst_AuditPoint!RequiredType = "D" Then
   mRequired = "D"
   TxtRequired = ""
   TxtActual = ""
   TxtValue = ""
   Label11.Visible = False
   Label12.Visible = False
   Label13.Visible = True
   Label14.Visible = False
 
   TxtRequired.Visible = False
   TxtActual.Visible = False
   TxtRequiredDate.Visible = True
   TxtValue.Visible = False
   
   
ElseIf RsMst_AuditPoint!RequiredType = "V" Then
   
   mRequired = "V"
   TxtRequired = ""
   TxtActual = ""
   Label11.Visible = False
   Label12.Visible = False
   Label13.Visible = False
   Label14.Visible = True
   TxtRequired.Visible = False
   TxtActual.Visible = False
   TxtRequiredDate.Visible = False
   TxtValue.Visible = True
   
End If

Call PositionForDisplay(mRequired)

End Sub
Private Sub PositionForDisplay(mRequired_ As Variant)

If mRequired_ = "N" Then
   
   Label11.Left = 60
   Label12.Left = 3420
  
   TxtRequired.Left = 1620
   TxtActual.Left = 4140
   
ElseIf RsMst_AuditPoint!RequiredType = "D" Then
   
   Label13.Left = 60
   TxtRequiredDate.Left = 1620
      
ElseIf RsMst_AuditPoint!RequiredType = "V" Then
   Label14.Left = 60
   TxtValue.Left = 1620
End If



End Sub

Private Sub CmbAuditSegment_GotFocus()
tmp = CmbAuditSegment.Text
Call TableMst_AuditSegment("")
CmbAuditSegment.Clear
If RsMst_AuditSegment.RecordCount = 0 Then
   MsgBox "No Audit Segment found"
   CmbAuditSegment.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditSegment.RecordCount
    CmbAuditSegment.AddItem RsMst_AuditSegment!SegmentType
    RsMst_AuditSegment.MoveNext
Next
CmbAuditSegment.Text = tmp
End Sub

Private Sub CmbAuditSegment_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAuditSegment")
End If
End Sub

Private Sub CmbAuditSegment_LostFocus()
If CmbAuditSegment.Text = "" Then
   mSegmentTypeID = Null
   Exit Sub
End If

RsMst_AuditSegment.MoveFirst
RsMst_AuditSegment.Find "SegmentType = '" & CmbAuditSegment.Text & "'"

If RsMst_AuditSegment.EOF Then
   MsgBox "Invalid Segment Selection "
   CmbAuditSegment.SetFocus
   Exit Sub
End If
mSegmentTypeID = RsMst_AuditSegment!SegmentTypeID

Call GetPreviousPoints(mSegmentTypeID, mGodownID, mAuditTypeID)

End Sub

Private Sub CmbAuditType_GotFocus()
tmp = CmbAuditType.Text
Call TableMst_AuditType("")
CmbAuditType.Clear
If RsMst_AuditType.RecordCount = 0 Then
   MsgBox "No Audit Type found"
   CmbAuditType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditType.RecordCount
    CmbAuditType.AddItem RsMst_AuditType!AuditType
    RsMst_AuditType.MoveNext
Next
CmbAuditType.Text = tmp
End Sub

Private Sub CmbAuditType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAuditType")
End If
End Sub

Private Sub CmbAuditType_LostFocus()
If CmbAuditType.Text = "" Then
   mAuditTypeID = Null
   Exit Sub
End If

RsMst_AuditType.MoveFirst
RsMst_AuditType.Find "AuditType = '" & CmbAuditType.Text & "'"

If RsMst_AuditType.EOF Then
   MsgBox "Invalid Audit type Selection "
   CmbAuditType.SetFocus
   Exit Sub
End If
mAuditTypeID = RsMst_AuditType!AuditTypeID
Call GetPreviousPoints(mSegmentTypeID, mGodownID, mAuditTypeID)
End Sub

Private Sub CmbDesignationID_GotFocus()
tmp = CmbDesignationID.Text
Call TableMst_Designation("")
CmbDesignationID.Clear
If RsMst_Designation.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbDesignationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Designation.RecordCount
    CmbDesignationID.AddItem RsMst_Designation!Designation
    RsMst_Designation.MoveNext
Next
CmbDesignationID.Text = tmp
End Sub

Private Sub CmbDesignationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbDesignationID")
End If
End Sub

Private Sub CmbDesignationID_LostFocus()
If CmbDesignationID.Text = "" Then
   mDesignationID = Null
   Exit Sub
End If

If CmbDesignationID.Text <> tmp Then
   CmbAuditor.Text = ""
End If

RsMst_Designation.MoveFirst
RsMst_Designation.Find "Designation = '" & CmbDesignationID.Text & "'"

If RsMst_Designation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbDesignationID.SetFocus
   Exit Sub
End If
mDesignationID = RsMst_Designation!DesignationID

End Sub

Private Sub CmbSearchCMPName_GotFocus()

tmp = CmbSearchCMPName.Text
Call TableMst_CMP '("Where CloseDate is Null ")

CmbSearchCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No CMP found"
   CmbSearchCMPName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSearchCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSearchCMPName.Text = tmp

End Sub

Private Sub CmbSearchCMPName_KeyPress(KeyAscii As Integer)

If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbSearchCMPName")
End If

End Sub

Private Sub CmbSearchCMPName_LostFocus()
If CmbSearchCMPName.Text = "" Then Exit Sub

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSearchCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid CMP selection "
   CmbSearchCMPName.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSearchSegment_GotFocus()
tmp = CmbSearchSegment.Text
Call TableMst_AuditSegment("")
CmbSearchSegment.Clear
If RsMst_AuditSegment.RecordCount = 0 Then
   MsgBox "No Audit Segment found"
   CmbSearchSegment.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditSegment.RecordCount
    CmbSearchSegment.AddItem RsMst_AuditSegment!SegmentType
    RsMst_AuditSegment.MoveNext
Next
CmbSearchSegment.Text = tmp

End Sub

Private Sub CmbSearchSegment_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbSearchSegment")
End If

End Sub

Private Sub CmbSearchSegment_LostFocus()
If CmbSearchSegment.Text = "" Then
   mSearchSegmentType = Null
   Exit Sub
End If

RsMst_AuditSegment.MoveFirst
RsMst_AuditSegment.Find "SegmentType = '" & CmbSearchSegment.Text & "'"

If RsMst_AuditSegment.EOF Then
   MsgBox "Invalid Segment Selection "
   CmbSearchSegment.SetFocus
   Exit Sub
End If
mSearchSegmentType = RsMst_AuditSegment!SegmentType

End Sub

Private Sub CmbSAuditor_GotFocus()
tmp = CmbSAuditor.Text
Call TableMst_Employee
CmbSAuditor.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee found"
   CmbSAuditor.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSAuditor.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSAuditor.Text = tmp
End Sub

Private Sub CmbSAuditor_LostFocus()
If CmbSAuditor.Text = "" Then
   Exit Sub
End If

Dim strSplitString() As String
Dim strEmpNo As String
Dim strEmpName As String
strSplitString = Split(CmbSAuditor.Text, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid selection!!!"
   CmbSAuditor.SetFocus
   Exit Sub
End If
strEmpNo = strSplitString(1)
strEmpName = strSplitString(0)

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & strEmpNo & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid  Selection "
   CmbSAuditor.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSAuditType_GotFocus()
tmp = CmbSAuditType.Text
Call TableMst_AuditType("")
CmbSAuditType.Clear
If RsMst_AuditType.RecordCount = 0 Then
   MsgBox "No Audit Type found"
   CmbSAuditType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditType.RecordCount
    CmbSAuditType.AddItem RsMst_AuditType!AuditType
    RsMst_AuditType.MoveNext
Next
CmbSAuditType.Text = tmp
End Sub

Private Sub CmbSAuditType_LostFocus()
If CmbSAuditType.Text = "" Then
   mAuditTypeID = Null
   Exit Sub
End If

RsMst_AuditType.MoveFirst
RsMst_AuditType.Find "AuditType = '" & CmbSAuditType.Text & "'"

If RsMst_AuditType.EOF Then
   MsgBox "Invalid Audit type Selection "
   CmbSAuditType.SetFocus
   Exit Sub
End If
mSAuditTypeID = RsMst_AuditType!AuditTypeID
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location("")
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSLocation.SetFocus
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
   MsgBox "Invalid Location selection "
   CmbSLocation.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdAdd2List_Click()
Dim I As Variant

If CmbAuditPoint.Text = "" Then
   MsgBox "Blank Audit Point not allowed!!!"
   CmbAuditPoint.SetFocus
   Exit Sub
End If

If TxtObservation = "" Then
   MsgBox "Blank Observation not allowed!!!"
   TxtObservation.SetFocus
   Exit Sub
End If

If Edit_Flg = "E" Then
   If TxtComment = "" Then
      MsgBox "Blank Comment not allowed!!!"
      TxtComment.SetFocus
      Exit Sub
   End If
End If

If mRequired = "N" Then
   If TxtRequired = "" Then
      MsgBox "Blank Required not allowed!!!"
      TxtRequired.SetFocus
      Exit Sub
   End If
   If TxtActual = "" Then
      MsgBox "Blank Actual not allowed!!!"
      TxtActual.SetFocus
      Exit Sub
   End If
ElseIf mRequired = "V" Then
   If TxtValue = "" Then
      MsgBox "Blank Value not allowed!!!"
      TxtValue.SetFocus
      Exit Sub
   End If
ElseIf mRequired = "D" Then
'   If TxtRequiredDate > Date Then
'      MsgBox "Required date can't be future date!!!"
'      TxtRequiredDate.SetFocus
'      Exit Sub
'   End If
Else
    CmbAuditPoint.SetFocus
    Exit Sub
End If

'If TxtStartDate.Value > Date Then
'   MsgBox "Start Date can't be future date!!!"
'   TxtStartDate.SetFocus
'   Exit Sub
'End If
If Check1.Value = 1 Then
   If TxtCloseDate.Value > Date Then
      MsgBox "Close date can't be future date!!!"
      TxtCloseDate.SetFocus
      Exit Sub
   End If
   If TxtCloseDate.Value < TxtStartDate.Value Then
      MsgBox "Close date must be greater than start date!!!"
      TxtCloseDate.SetFocus
      Exit Sub
   End If
End If

If CheckDuplicate = True Then
   MsgBox "Duplicate Entry for Audit Point!!!"
   CmbAuditPoint.SetFocus
   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   I = MSHFlexGrid2.RowSel
Else
    I = MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    
End If

MSHFlexGrid2.TextMatrix(I, 0) = ""
MSHFlexGrid2.TextMatrix(I, 1) = CmbAuditPoint.Text
MSHFlexGrid2.TextMatrix(I, 3) = TxtObservation
MSHFlexGrid2.TextMatrix(I, 5) = TxtComment
MSHFlexGrid2.TextMatrix(I, 6) = mRequired
MSHFlexGrid2.TextMatrix(I, 7) = Val(TxtRequired)
MSHFlexGrid2.TextMatrix(I, 8) = Val(TxtActual)
If mRequired = "D" Then
   MSHFlexGrid2.TextMatrix(I, 9) = TxtRequiredDate.Value
End If
MSHFlexGrid2.TextMatrix(I, 10) = Val(TxtValue)

MSHFlexGrid2.TextMatrix(I, 11) = TxtStartDate.Value
If Check1.Value = 1 Then
   MSHFlexGrid2.TextMatrix(I, 12) = TxtCloseDate.Value
Else
   MSHFlexGrid2.TextMatrix(I, 12) = "" 'TxtCloseDate.Value
End If
MSHFlexGrid2.TextMatrix(I, 13) = TxtAuditPointID
MSHFlexGrid2.TextMatrix(I, 14) = TxtAuditDetailsID
MSHFlexGrid2.TextMatrix(I, 15) = TxtOldAuditID
MSHFlexGrid2.TextMatrix(I, 16) = 1

CmdAdd2List.Caption = "Add to List"
Call ClearFrame
CmbAuditPoint.SetFocus
TxtStartDate.Value = TxtAuditDate.Value
End Sub
Private Sub ClearFrame()

TxtAuditPointID = ""
CmbAuditPoint.Text = ""
TxtObservation = ""
TxtComment = ""
mRequired = Null
TxtRequired = ""
TxtActual = ""
TxtRequiredDate.Value = Date
TxtValue = ""
'TxtStartDate.Value = Date
Check1.Value = 0
TxtCloseDate.Value = Date
TxtOldAuditID = ""
TxtAuditDetailsID = ""
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
TxtRequired.Visible = False
TxtActual.Visible = False
TxtRequiredDate.Visible = False
TxtValue.Visible = False
CmdAdd2List.Caption = "Add to List"
CmbAuditPoint.Enabled = True
CmdAuditPoint.Enabled = True

End Sub

Private Sub CmdAuditor_Click()
FrmMst_Employee.Show
End Sub

Private Sub CmdAuditPoint_Click()
FrmMst_AuditPoint.Show
End Sub

Private Sub CmdAuditSegment_Click()
FrmMst_AuditSegment.Show
End Sub

Private Sub CmdAuditType_Click()
FrmMst_AuditType.Show
End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date - 365
TxtToDate.Value = Date
CmbSLocation.Text = ""
CmbSAuditor.Text = ""
CmbSAuditType.Text = ""
CmbStatus.Text = ""
Label22.Caption = ""
Call Grid_Head
End Sub

Private Sub CmdGo_Click()
Dim tmp1 As Variant
Dim Wc As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim dblTmpAuditNo   As Double

Call Grid_Head

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "'From Date' must be less than 'To Date' always !!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

tmp = ""

tmp = "Select A.AuditID,A.AuditDate, AP.AuditPoint, AD.CloseDate, L.LocationName,C.CMPName,G.GodownNo,D.Designation,"
tmp = tmp & " Emp.EmployeeName + '~' + Emp.EmployeeNo as Auditor,AType.AuditType" ',AD.StartDate,AD.CloseDate"
tmp = tmp & " From Txn_Audit A"
tmp = tmp & " Inner Join Mst_AuditSegment S On A.SegmentTypeID = S.SegmentTypeID"
tmp = tmp & " Inner Join Txn_AuditDetails AD on AD.AuditID = A.AuditID"
tmp = tmp & " Inner Join Mst_Godown G On A.GodownID = G.GodownID"
tmp = tmp & " Inner Join Mst_Designation D On A.DesignationID = D.DesignationID"
tmp = tmp & " Inner Join Mst_Employee EMP On A.EmployeeID = EMP.EmployeeID"
tmp = tmp & " Inner Join Mst_AuditType AType on A.AuditTypeID = AType.AuditTypeID"
tmp = tmp & " Inner Join Mst_AuditPoint AP on AD.AuditPointID = AP.AuditPointID"
tmp = tmp & " Inner Join Mst_CMP C on G.CMPID = C.CMPID"
tmp = tmp & " Inner Join Mst_Location L on C.LocationID = L.LocationID"
tmp = tmp & " Where A.Flag = 'Y'" ' And AD.CloseDate is Null "

Wc = ""

If Gen_DBType = "MDB" Then
   Wc = " And A.AuditDate Between #" & TxtFromDate.Value & "# And #" & TxtToDate.Value & "# "
Else
   Wc = " And A.AuditDate Between '" & TxtFromDate.Value & "' And '" & TxtToDate.Value & "' "
End If

If CmbSLocation.Text <> "" Then
   Wc = Wc & " And L.LocationName = '" & CmbSLocation.Text & "'  "
Else
   If Gen_WcLocation <> "" Then
      Wc = Wc & " And L." & Gen_WcLocation
   End If
End If

If CmbSAuditor.Text <> "" Then
   Wc = Wc & " AND EMP.EmployeeNo = '" & EmpNo_ & "' "
End If

If CmbSAuditType.Text <> "" Then
   Wc = Wc & " And AType.AuditTypeID= " & mSAuditTypeID & ""
End If

If Left(CmbStatus.Text, 1) = "O" Then
   Wc = Wc & " And AD.CloseDate IS NULL"
ElseIf Left(CmbStatus.Text, 1) = "C" Then
   Wc = Wc & " And AD.CloseDate IS NOT NULL"
End If

tmp = tmp & " " & Wc

tmp = tmp & " Order by A.AuditID,A.AuditDate,L.LocationName,S.SegmentType"

Call TmpTable

dblTmpAuditNo = 0

Screen.MousePointer = vbHourglass
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    'For c = 0 To MSHFlexGrid1.Cols - 1
    DoEvents
'    If dblTmpAuditNo <> TmpRs.Fields("AuditID") Then
       MSHFlexGrid1.TextMatrix(iu, 0) = IIf(IsNull(TmpRs.Fields("AuditID")), "", TmpRs.Fields("AuditID"))
       MSHFlexGrid1.TextMatrix(iu, 1) = IIf(IsNull(TmpRs.Fields("AuditDate")), "", TmpRs.Fields("AuditDate"))
       MSHFlexGrid1.TextMatrix(iu, 2) = IIf(IsNull(TmpRs.Fields("LocationName")), "", TmpRs.Fields("LocationName"))
       MSHFlexGrid1.TextMatrix(iu, 3) = IIf(IsNull(TmpRs.Fields("CMPName")), "", TmpRs.Fields("CMPName"))
       MSHFlexGrid1.TextMatrix(iu, 4) = IIf(IsNull(TmpRs.Fields("GodownNo")), "", TmpRs.Fields("GodownNo"))
       MSHFlexGrid1.TextMatrix(iu, 5) = IIf(IsNull(TmpRs.Fields("Designation")), "", TmpRs.Fields("Designation"))
       MSHFlexGrid1.TextMatrix(iu, 6) = IIf(IsNull(TmpRs.Fields("Auditor")), "", TmpRs.Fields("Auditor"))
       MSHFlexGrid1.TextMatrix(iu, 7) = IIf(IsNull(TmpRs.Fields("AuditType")), "", TmpRs.Fields("AuditType"))
       MSHFlexGrid1.TextMatrix(iu, 8) = IIf(IsNull(TmpRs.Fields("AuditPoint")), "", TmpRs.Fields("AuditPoint"))
       MSHFlexGrid1.TextMatrix(iu, 9) = IIf(IsNull(TmpRs.Fields("CloseDate")), "", TmpRs.Fields("CloseDate"))
'    End If
 '   dblTmpAuditNo = TmpRs.Fields("AuditID")
    'Next
    MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(iu, 9) = Format(MSHFlexGrid1.TextMatrix(iu, 9), "dd-mmm-yyyy")
    TmpRs.MoveNext
Next
Screen.MousePointer = vbNormal

Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub
Private Sub CmdGodown_Click()
FrmMstGodown.Show
End Sub

Private Sub CmdLocation_Click()
FrmLocationMaster.Show
End Sub

Private Sub CmdRemoveFrmList_Click()
If MSHFlexGrid2.Rows < 3 Then Exit Sub
Dim ans As Variant
If CmbAuditPoint.Text = "" Then
   MsgBox "Please Select atleast one row from grid!!!"
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   I = MSHFlexGrid2.RowSel
   If I > 0 Then
      If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
         For RO = I To MSHFlexGrid2.Rows - 2
             For c = 0 To MSHFlexGrid2.Cols - 1
                 MSHFlexGrid2.TextMatrix(RO, c) = MSHFlexGrid2.TextMatrix(RO + 1, c)
             Next
         Next
         MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
      End If
   End If
   Call ClearFrame
Else
   'CmbAuditPoint.SetFocus
   TxtObservation.SetFocus
End If
End Sub

Private Sub CmdWarehouse_Click()
FrmMst_CMP.Show
End Sub



Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
Frame3.Enabled = True
Frame5.Enabled = True

CmbAuditSegment.Visible = True
CmdAuditSegment.Visible = True
TxtAuditSegment.Visible = False
CmbAuditSegment.Text = TxtAuditSegment

CmbLocation.Visible = True
CmdLocation.Visible = True
TxtLocationName.Visible = False
CmbLocation.Text = TxtLocationName

CmbCMPID.Visible = True
CmdWarehouse.Visible = True
TxtCMPID.Visible = False
CmbCMPID.Text = TxtCMPID

CmbGodownName.Visible = True
CmdGodown.Visible = True
TxtGodownName.Visible = False
CmbGodownName.Text = TxtGodownName

CmbDesignationID.Visible = True
TxtDesignation.Visible = False
CmbDesignationID.Text = TxtDesignation

CmbAuditor.Visible = True
TxtAuditor.Visible = False
CmbAuditor.Text = TxtAuditor
CmdAuditor.Visible = True

CmbAuditType.Visible = True
TxtAuditType.Visible = False
CmdAuditType.Visible = True
CmbAuditType.Text = TxtAuditType

Frame1.Enabled = False

Call GButtonLock(Me, False)
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
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call TableTxn_Audit(" Where Flag ='Y' Order by AuditID ")

If RsTxn_Audit.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call GButtonLock(Me, True)
Call CmdClear_Click
Call Indata
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
Call TableMst_Location("")

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocation")
End If
End Sub

Private Sub CmbLocation_LostFocus()

If CmbLocation.Text = "" Then
   mLocationID = 0
   CmbCMPID.Clear
   Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocation.Text <> tmp Then
   CmbCMPID.Clear
   CmbCMPID.Text = ""
   CmbGodownName.Clear
   CmbGodownName.Text = ""
   TxtGodownAddress = ""
   mCMPID = 0
   mGodownID = 0
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID

End Sub


Private Sub CmbGodownName_Gotfocus()
If IsNull(mWarehoseID) Then Exit Sub
If CmbCMPID.Text = "" Then
   CmbGodownName.Clear
   Exit Sub
End If
tmp = CmbGodownName.Text
Call TableMst_Godown(" Where CMPID = " & mCMPID & " and CloseDate is Null ")
CmbGodownName.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Godown found  !!!!! "
   CmbGodownName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownName.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownName.Text = tmp

End Sub

Private Sub CmbGodownName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbGodownName")
End If
End Sub

Private Sub CmbGodownName_LostFocus()
If CmbGodownName.Text = "" Then
   mGodownID = 0
   Exit Sub
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownName.Text & "'"

If RsMst_Godown.EOF Then
   MsgBox "Invalid Godown Selection "
   CmbGodownName.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
TxtGodAddress = RsMst_Godown!Address

Call GetPreviousPoints(mSegmentTypeID, mGodownID, mAuditTypeID)

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant

'Dim strSplitString() As String
'Dim strObservation As String
'Dim strComment As String


If SaveCmd.Enabled = False Then Exit Sub
I = MSHFlexGrid2.RowSel
If I = 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit Sub

TxtAuditPointID = MSHFlexGrid2.TextMatrix(I, 13)
CmbAuditPoint = MSHFlexGrid2.TextMatrix(I, 1)
CmbAuditPoint.Enabled = False
CmdAuditPoint.Enabled = False
'If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
'   strSplitString = Split(MSHFlexGrid2.TextMatrix(i, 2), "~")
'   strObservation = strSplitString(0)
'   TxtObservation = strObservation
'   strComment = strSplitString(UBound(strSplitString))
'Else
'   TxtObservation = MSHFlexGrid2.TextMatrix(i, 3)
'End If

 'MSHFlexGrid2.TextMatrix(i, 3)

'Commented on 17 Nov 2010

TxtObservation = MSHFlexGrid2.TextMatrix(I, 3)
TxtComment = MSHFlexGrid2.TextMatrix(I, 5)

'If MSHFlexGrid2.TextMatrix(i, 4) <> "" Then
'   strSplitString = Split(MSHFlexGrid2.TextMatrix(i, 4), "~")
'   strComment = strSplitString(0)
'   TxtComment = strComment
'   TxtComment = strSplitString(UBound(strSplitString))
'Else
'   TxtComment = MSHFlexGrid2.TextMatrix(i, 5)
'End If


mRequired = MSHFlexGrid2.TextMatrix(I, 6)
TxtRequired = Val(MSHFlexGrid2.TextMatrix(I, 7))
TxtActual = Val(MSHFlexGrid2.TextMatrix(I, 8))
If MSHFlexGrid2.TextMatrix(I, 9) <> "" Then
   TxtRequiredDate = CDate(MSHFlexGrid2.TextMatrix(I, 9))
End If
TxtValue = Val(MSHFlexGrid2.TextMatrix(I, 10))
TxtStartDate.Value = CDate(MSHFlexGrid2.TextMatrix(I, 11))
If MSHFlexGrid2.TextMatrix(I, 12) <> "" Then
   Check1.Value = 1
   TxtCloseDate.Value = CDate(MSHFlexGrid2.TextMatrix(I, 12))
Else
   Check1.Value = 0
   'TxtCloseDate = CDate(MSHFlexGrid2.TextMatrix(i, 11))
End If

TxtOldAuditID = MSHFlexGrid2.TextMatrix(I, 15)
TxtAuditDetailsID = MSHFlexGrid2.TextMatrix(I, 14)
TxtObservation.SetFocus
CmdAdd2List.Caption = "Update List"
'CmbAuditPoint.SetFocus
Call CmbAuditPoint_GotFocus
Call CmbAuditPoint_LostFocus

TxtObservation.SetFocus

End Sub

Private Sub SaveCmd_Click()
Dim mObservation As Variant
Dim mComment As Variant
If TxtAuditDate.Value > Date Then
   MsgBox "Future date not allowed!!!"
   TxtAuditDate.SetFocus
   Exit Sub
End If

If CmbAuditSegment.Text = "" Then
   MsgBox "Blank Audit segment not allowed!!!"
   CmbAuditSegment.SetFocus
   Exit Sub
End If

If CmbLocation.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

If CmbGodownName.Text = "" Then
   MsgBox "Blank Godown not allowed!!!"
   CmbGodownName.SetFocus
   Exit Sub
End If

If CmbDesignationID.Text = "" Then
   MsgBox "Blank Designation not allowed!!!"
   CmbDesignationID.SetFocus
   Exit Sub
End If

If CmbAuditor.Text = "" Then
   MsgBox "Blank Auditor not allowed!!!"
   CmbAuditor.SetFocus
   Exit Sub
End If


If MSHFlexGrid2.Rows < 3 Then
   MsgBox "Blank grid not allowed!!!"
   CmbAuditPoint.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   tmp = "Select * from Txn_Audit Where Flag <> 'N' And  SegmentTypeID = " & mSegmentTypeID & " And GodownID = " & mGodownID & " And AuditTypeID = " & mAuditTypeID & " "
   Call Tmp4Table
   If TmpRs4.RecordCount > 0 Then
      RsTxn_Audit.MoveFirst
      RsTxn_Audit.Find "AuditID = " & TmpRs4!AuditID
      RsTxn_Audit!Flag = "N"
      RsTxn_Audit.Update
   End If
   RsTxn_Audit.AddNew
   RsTxn_Audit!AuditID = GetMaxAuditID()
   TxtAuditID = RsTxn_Audit!AuditID
   RsTxn_Audit!G_UID = GetGUID
Else
   RsTxn_Audit.MoveFirst
   RsTxn_Audit.Find "AuditID = " & TxtAuditID & ""
   If RsTxn_Audit.EOF Then
      Exit Sub
   End If
End If

RsTxn_Audit!AuditDate = TxtAuditDate.Value
RsTxn_Audit!SegmentTypeID = mSegmentTypeID
RsTxn_Audit!GodownID = mGodownID
RsTxn_Audit!DesignationID = mDesignationID
RsTxn_Audit!EmployeeID = mEmployeeID
RsTxn_Audit!AuditTypeID = mAuditTypeID
RsTxn_Audit!Flag = "Y"
RsTxn_Audit!Remarks = Trim(TxtRemarks.Text)

If IsNull(RsTxn_Audit!CreatedBy) = True Or RsTxn_Audit!CreatedBy = "" Then
   RsTxn_Audit!CreatedBy = Gen_UserLoginID
   RsTxn_Audit!CreatedDate = Now
Else
   RsTxn_Audit!UpdatedBy = Gen_UserLoginID
   RsTxn_Audit!UpdatedDate = Now
End If
RsTxn_Audit.Update

Call TableTxn_AuditDetails '(" Where AuditID = " & TxtAuditID & "")
For I = 1 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(I, 15) = "" Then
          'RsTxn_AuditDetails
          RsTxn_AuditDetails.AddNew
          RsTxn_AuditDetails!AuditDetailsID = GetMaxAuditDetailsID()
          RsTxn_AuditDetails!StartDate = TxtAuditDate
       Else
          RsTxn_AuditDetails.MoveFirst
          mAuditID = MSHFlexGrid2.TextMatrix(I, 15)
          RsTxn_AuditDetails.Filter = "AuditID = " & MSHFlexGrid2.TextMatrix(I, 15) & " And AuditPointID = " & MSHFlexGrid2.TextMatrix(I, 13) & ""
          
       End If
       If Not RsTxn_AuditDetails.EOF Then
          'RsTxn_AuditDetails!StartDate = TxtAuditDate
          
          RsTxn_AuditDetails!AuditPointID = MSHFlexGrid2.TextMatrix(I, 13)
          mObservation = MSHFlexGrid2.TextMatrix(I, 2) 'IIf(IsNull(RsTxn_AuditDetails!Observation), "", RsTxn_AuditDetails!Observation)
          If mObservation = "" Then
             RsTxn_AuditDetails!Observation = MSHFlexGrid2.TextMatrix(I, 3)
          Else
'             If MSHFlexGrid2.TextMatrix(i, 3) <> "" Then
'                RsTxn_AuditDetails!Observation = MSHFlexGrid2.TextMatrix(i, 3) & "~" & mObservation
'             Else
'                RsTxn_AuditDetails!Observation = mObservation
'             End If

             If MSHFlexGrid2.TextMatrix(I, 16) = 1 Then
                RsTxn_AuditDetails!Observation = MSHFlexGrid2.TextMatrix(I, 3) & "~" & mObservation
             Else
                RsTxn_AuditDetails!Observation = mObservation
             End If

          End If
          mComment = MSHFlexGrid2.TextMatrix(I, 4) 'IIf(IsNull(RsTxn_AuditDetails!Comment), "", RsTxn_AuditDetails!Comment) 'RsTxn_AuditDetails!Comment
          If mComment = "" Then
'             If MSHFlexGrid2.TextMatrix(i, 5) <> "" Then
             RsTxn_AuditDetails!Comment = MSHFlexGrid2.TextMatrix(I, 5)
'             Else
'                RsTxn_AuditDetails!Comment = ""
'             End If
          Else
'             If MSHFlexGrid2.TextMatrix(i, 5) <> "" Then
'                RsTxn_AuditDetails!Comment = MSHFlexGrid2.TextMatrix(i, 5) & "~" & mComment
'             Else
'                RsTxn_AuditDetails!Comment = mComment
'             End If
          
             If MSHFlexGrid2.TextMatrix(I, 16) = 1 Then
                RsTxn_AuditDetails!Comment = MSHFlexGrid2.TextMatrix(I, 5) & "~" & mComment
             Else
                RsTxn_AuditDetails!Comment = mComment
             End If
          End If
          
          RsTxn_AuditDetails!RequiredType = MSHFlexGrid2.TextMatrix(I, 6)
          RsTxn_AuditDetails!Required = Val(MSHFlexGrid2.TextMatrix(I, 7))
          RsTxn_AuditDetails!Actual = Val(MSHFlexGrid2.TextMatrix(I, 8))
          If MSHFlexGrid2.TextMatrix(I, 6) = "D" Then
             RsTxn_AuditDetails!RequiredDate = CDate(MSHFlexGrid2.TextMatrix(I, 9))
          End If
          RsTxn_AuditDetails!RequiredValue = MSHFlexGrid2.TextMatrix(I, 10)
          'RsTxn_AuditDetails!StartDate = CDate(MSHFlexGrid2.TextMatrix(i, 10))
          If MSHFlexGrid2.TextMatrix(I, 12) <> "" Then
             RsTxn_AuditDetails!CloseDate = MSHFlexGrid2.TextMatrix(I, 12)
             If MSHFlexGrid2.TextMatrix(I, 15) = "" Then
                RsTxn_AuditDetails!AuditID = TxtAuditID
             End If
             RsTxn_AuditDetails!Flag = "N"
          Else
             RsTxn_AuditDetails!AuditID = TxtAuditID
          End If
          'RsTxn_AuditDetails!BackAuditID = IIf(MSHFlexGrid2.TextMatrix(i, 12) <> "", MSHFlexGrid2.TextMatrix(i, 12), Null)
          If IsNull(RsTxn_AuditDetails!CreatedBy) = True Or RsTxn_AuditDetails!CreatedBy = "" Then
             RsTxn_AuditDetails!CreatedBy = Gen_UserLoginID
             RsTxn_AuditDetails!CreatedDate = Now
          Else
             RsTxn_AuditDetails!UpdatedBy = Gen_UserLoginID
             RsTxn_AuditDetails!UpdatedDate = Now
          End If
          
          RsTxn_AuditDetails.Update
       End If
       RsTxn_AuditDetails.Filter = ""
    End If
Next
RsTxn_Audit.Requery
RsTxn_Audit.MoveFirst
RsTxn_Audit.Find "AuditID = " & TxtAuditID
If RsTxn_Audit.EOF Then
   RsTxn_Audit.MoveLast
End If
Call Indata

End Sub
Private Sub Indata()
Dim strSplitString() As String
Dim strEmpNo As String
Dim strEmpName As String
Dim strObservation As String
Dim strComment As String

Call GButtonLock(Me, True)

TxtAuditID = RsTxn_Audit!AuditID
TxtAuditDate.Value = RsTxn_Audit!AuditDate
TxtAuditDate.Enabled = False

CmbAuditSegment.Visible = False
CmdAuditSegment.Visible = False
TxtAuditSegment.Visible = True
TxtAuditSegment.Text = GetAuditSegmentType(RsTxn_Audit!SegmentTypeID)
mSegmentTypeID = RsTxn_Audit!SegmentTypeID
CmbAuditSegment.Text = TxtAuditSegment

CmbGodownName.Visible = False
CmdGodown.Visible = False
TxtGodownName.Visible = True
mGodownID = RsTxn_Audit!GodownID
TxtGodownName = GetGodownName(mGodownID, "N")
TxtGodAddress = GetGodownName(mGodownID, "A")

CmbGodownName.Text = TxtGodownName

CmbCMPID.Visible = False
CmdWarehouse.Visible = False
TxtCMPID.Visible = True
TxtCMPID = GetCMPFromGodown(RsTxn_Audit!GodownID)
mCMPID = GetCMPID(TxtCMPID)
CmbCMPID.Text = TxtCMPID

CmbLocation.Visible = False
CmdLocation.Visible = False
TxtLocationName.Visible = True
'mLocationId = RsTxn_Audit!LocationID
'TxtLocationName = GetLocationName(RsTxn_Audit!LocationID)
TxtLocationName = GetLocationFromCMP(mCMPID)
mLocationID = GetLocationIDFromCMP(mCMPID)
CmbLocation.Text = TxtLocationName


CmbDesignationID.Visible = False
TxtDesignation.Visible = True
mDesignationID = RsTxn_Audit!DesignationID
TxtDesignation = GetDesignaion(mDesignationID)
CmbDesignationID.Text = TxtDesignation

CmbAuditor.Visible = False
TxtAuditor.Visible = True
CmdAuditor.Visible = False
mEmployeeID = RsTxn_Audit!EmployeeID
TxtAuditor = GetEmployeeNonName(mEmployeeID)
CmbAuditor.Text = TxtAuditor
'strSplitString = Split(TxtAuditor.Text, "~")
'strEmpName = strSplitString(1)
'strEmpNo = strSplitString(0)
'TxtAuditor = strEmpName & "~" & strEmpNo

CmbAuditType.Visible = False
CmdAuditType.Visible = False
TxtAuditType.Visible = True
mAuditTypeID = RsTxn_Audit!AuditTypeID
TxtAuditType = GetAuditType(mAuditTypeID)
CmbAuditType.Text = TxtAuditType
TxtRemarks.Text = IIf(IsNull(RsTxn_Audit!Remarks), "", RsTxn_Audit!Remarks)

TxtCreated = IIf(IsNull(RsTxn_Audit!CreatedBy), "", RsTxn_Audit!CreatedBy) & " :- " & IIf(IsNull(RsTxn_Audit!CreatedDate), "", RsTxn_Audit!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_Audit!UpdatedBy), "", RsTxn_Audit!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_Audit!UpdatedDate), "", RsTxn_Audit!UpdatedDate)


Call Grid_Head1

Call TableTxn_AuditDetails("Where AuditID = " & TxtAuditID & "")

If RsTxn_AuditDetails.RecordCount > 0 Then
   For I = 1 To RsTxn_AuditDetails.RecordCount
       MSHFlexGrid2.TextMatrix(I, 0) = ""
       MSHFlexGrid2.TextMatrix(I, 1) = GetAuditPoint(RsTxn_AuditDetails!AuditPointID)
       
       MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(RsTxn_AuditDetails!Observation), "", RsTxn_AuditDetails!Observation)

'       MSHFlexGrid2.TextMatrix(i, 3) = ""
       
       If IsNull(RsTxn_AuditDetails!Observation) <> True And (RsTxn_AuditDetails!Observation) <> "" Then
          strSplitString = Split(RsTxn_AuditDetails!Observation, "~")
          strObservation = strSplitString(0)
          MSHFlexGrid2.TextMatrix(I, 3) = strObservation
       Else
          MSHFlexGrid2.TextMatrix(I, 3) = MSHFlexGrid2.TextMatrix(I, 2)
       End If
        
       
       MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(RsTxn_AuditDetails!Comment), "", RsTxn_AuditDetails!Comment)
'       MSHFlexGrid2.TextMatrix(i, 5) = ""
       
       If IsNull(RsTxn_AuditDetails!Comment) <> True And (RsTxn_AuditDetails!Comment) <> "" Then
          strSplitString = Split(RsTxn_AuditDetails!Comment, "~")
          strComment = strSplitString(0)
          MSHFlexGrid2.TextMatrix(I, 5) = strComment
       Else
          MSHFlexGrid2.TextMatrix(I, 5) = MSHFlexGrid2.TextMatrix(I, 4)
       End If
       
       
       MSHFlexGrid2.TextMatrix(I, 6) = RsTxn_AuditDetails!RequiredType
       MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(RsTxn_AuditDetails!Required), "", RsTxn_AuditDetails!Required)
       MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(RsTxn_AuditDetails!Actual), "", RsTxn_AuditDetails!Actual)
       MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(RsTxn_AuditDetails!RequiredDate), "", RsTxn_AuditDetails!RequiredDate)
       MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(RsTxn_AuditDetails!RequiredValue), "", RsTxn_AuditDetails!RequiredValue)
       MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(RsTxn_AuditDetails!StartDate), "", RsTxn_AuditDetails!StartDate)
       MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(RsTxn_AuditDetails!CloseDate), "", RsTxn_AuditDetails!CloseDate)
       'MSHFlexGrid2.TextMatrix(i, 12) = IIf(IsNull(RsTxn_AuditDetails!OldAuditDetailsID), "", RsTxn_AuditDetails!OldAuditDetailsID)
       MSHFlexGrid2.TextMatrix(I, 13) = IIf(IsNull(RsTxn_AuditDetails!AuditPointID), "", RsTxn_AuditDetails!AuditPointID)
       MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(RsTxn_AuditDetails!AuditDetailsID), "", RsTxn_AuditDetails!AuditDetailsID)
       MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(RsTxn_AuditDetails!AuditID), "", RsTxn_AuditDetails!AuditID)
       MSHFlexGrid2.TextMatrix(I, 16) = 0
       RsTxn_AuditDetails.MoveNext
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   Next
End If

Frame3.Enabled = False
Frame5.Enabled = False
Call ClearFrame

'EditCmd.Enabled = False
'DeleteCmd.Enabled = False

End Sub
Private Function GetMaxAuditID() As Double
Dim Retval As Variant
Retval = 0
tmp = "select max(AuditID) from Txn_Audit "
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs4.Fields(0)), 0, TmpRs4.Fields(0))
End If
TmpRs4.Close
Retval = Retval + 1

GetMaxAuditID = Retval
End Function
Private Function GetMaxAuditDetailsID() As Double
Dim Retval As Variant
Retval = 0
tmp = "select max(AuditDetailsID) from Txn_AuditDetails "
Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs4.Fields(0)), 0, TmpRs4.Fields(0))
End If
TmpRs4.Close
Retval = Retval + 1
GetMaxAuditDetailsID = Retval
End Function

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_Audit.RecordCount = 0 Then Exit Sub

   Call GButtonLock(Me, True)
   If RsTxn_Audit.EOF Then
      RsTxn_Audit.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 600
Frame0.Visible = True
'MSHFlexGrid1.Clear
MSHFlexGrid1.Width = Frame0.Width - 2500
MSHFlexGrid1.Height = Frame0.Height - 1400

'Call CmdClear_Click

End Sub

Private Sub TxtActual_LostFocus()
If TxtActual = "" Then Exit Sub

If IsNumeric(TxtActual) = False Or Val(TxtActual) < 0 Then
   MsgBox "Invalid Numric format!!!"
   TxtActual.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtAuditDate_LostFocus()
TxtStartDate.Value = TxtAuditDate.Value
End Sub

Private Sub TxtCloseDate_LostFocus()
If TxtCloseDate.Value > Date Then
   MsgBox "Future Date not allowed!!!"
   TxtCloseDate.SetFocus
   Exit Sub
End If
End Sub
Private Sub TxtRequired_LostFocus()
If TxtRequired = "" Then Exit Sub

If IsNumeric(TxtRequired) = False Or Val(TxtRequired) < 0 Then
   MsgBox "Invalid Numric format!!!"
   TxtRequired.SetFocus
   Exit Sub
End If

End Sub
Private Sub TxtValue_LostFocus()
If TxtValue = "" Then Exit Sub

If IsNumeric(TxtValue) = False Or Val(TxtValue) < 0 Then
   MsgBox "Invalid Numric format!!!"
   TxtValue.SetFocus
   Exit Sub
End If
End Sub

Private Sub GetPreviousPoints(mSegment_ As Variant, mgodownID_ As Double, mAuditTypeID_ As Variant)
If IsNull(mSegment_) = True Then Exit Sub
If mgodownID_ = 0 Then Exit Sub
If IsNull(mAuditTypeID_) = True Then Exit Sub

'tmp = ""

tmp = "Select AD.AuditDetailsID,A.AuditID,AP.AuditPointID,AP.AuditPoint,AD.Observation,AD.Comment,AD.RequiredType, " & _
      "AD.Required , AD.Actual, AD.RequiredDate, AD.RequiredValue, AD.StartDate, AD.CloseDate " & _
      "from Txn_AuditDetails AD " & _
      "Inner Join Txn_Audit A On AD.AuditID = A.AuditID " & _
      "Inner Join Mst_AuditPoint AP On AD.AuditPointID = AP.AuditPointID " & _
      "Where A.SegmentTypeID = " & mSegment_ & " And A.GodownID = " & mgodownID_ & " And A.AuditTypeID = " & mAuditTypeID_ & " And AD.CloseDate is Null  "

Call Tmp4Table

Call Grid_Head1
If TmpRs4.RecordCount > 0 Then
   For I = 1 To TmpRs4.RecordCount
       MSHFlexGrid2.TextMatrix(I, 0) = ""
       MSHFlexGrid2.TextMatrix(I, 1) = TmpRs4!AuditPoint 'Audit Point"
       
       MSHFlexGrid2.TextMatrix(I, 2) = TmpRs4!Observation
       
       
'       MSHFlexGrid2.TextMatrix(I, 3) = "" 'TmpRs4!""
       
'       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs4!Comment
       
       
       If IsNull(TmpRs4!Observation) <> True And (TmpRs4!Observation) <> "" Then
          strSplitString = Split(TmpRs4!Observation, "~")
          strObservation = strSplitString(0)
          MSHFlexGrid2.TextMatrix(I, 3) = strObservation
       Else
          MSHFlexGrid2.TextMatrix(I, 3) = MSHFlexGrid2.TextMatrix(I, 2)
       End If
        
       
       MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs4!Comment), "", TmpRs4!Comment)
'       MSHFlexGrid2.TextMatrix(i, 5) = ""
       
       If IsNull(TmpRs4!Comment) <> True And (TmpRs4!Comment) <> "" Then
          strSplitString = Split(TmpRs4!Comment, "~")
          strComment = strSplitString(0)
          MSHFlexGrid2.TextMatrix(I, 5) = strComment
       Else
          MSHFlexGrid2.TextMatrix(I, 5) = MSHFlexGrid2.TextMatrix(I, 4)
       End If
       
       
'       MSHFlexGrid2.TextMatrix(I, 5) = "" 'TmpRs4!Comment
       MSHFlexGrid2.TextMatrix(I, 6) = TmpRs4!RequiredType
       MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs4!Required), "", TmpRs4!Required) '"Required"
       MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(TmpRs4!Actual), "", TmpRs4!Actual) '"Actual"
       MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!RequiredDate), "", TmpRs4!RequiredDate) '"Required Date"
       MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!RequiredValue), "", TmpRs4!RequiredValue) '"Value"
       MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs4!StartDate), "", TmpRs4!StartDate) '"Start Date"
       MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!CloseDate), "", TmpRs4!CloseDate) '"Close Date"
       MSHFlexGrid2.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!AuditPointID), "", TmpRs4!AuditPointID) '"AuditPointID"
       MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs4!AuditDetailsID), "", TmpRs4!AuditDetailsID) '"AuditDetailsID"
       MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(TmpRs4!AuditID), "", TmpRs4!AuditID) '"Old AuditID"
       MSHFlexGrid2.TextMatrix(I, 16) = 1
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       TmpRs4.MoveNext
   Next
End If


End Sub

Private Sub CmbCMPID_GotFocus()
If IsNull(mLocationID) Then Exit Sub
If CmbLocation.Text = "" Then Exit Sub
tmp = CmbCMPID.Text
Call TableMst_CMP("Where LocationID = " & mLocationID & " and CloseDate is Null ")

CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No CMP found"
   CmbCMPID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub

Private Sub CmbCMPID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbCMPID")
End If
End Sub

Private Sub CmbCMPID_LostFocus()

If CmbCMPID.Text = "" Then
   mCMPID = 0
   Exit Sub
End If

' Code for clearing Godown combox in case of change
If CmbCMPID.Text <> tmp Then
   CmbGodownName.Clear
   CmbGodownName.Text = ""
   mGodownID = 0
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid CMP selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_Audit.MoveNext
If RsTxn_Audit.EOF Then
   RsTxn_Audit.MoveLast
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
RsTxn_Audit.MovePrevious
If RsTxn_Audit.BOF Then
   RsTxn_Audit.MoveFirst
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
RsTxn_Audit.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_Audit.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = Fasle
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_Audit.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_Audit.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsTxn_Audit.MoveFirst

If RsTxn_Audit.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_Audit.MoveFirst
RsTxn_Audit.Find "AuditID = " & Val(mCaId) & ""
If RsTxn_Audit.EOF Then
   RsTxn_Audit.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub


Private Function CheckDuplicate() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(I, 13) = TxtAuditPointID Then
          If CmdAdd2List.Caption = "Add to List" Then
             Retval = True
             Exit For
          End If
       End If
    End If
Next I
CheckDuplicate = Retval
End Function
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
HScroll1.Max = Me.ScaleWidth - Frame1.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True
If Frame1.Width < Me.Width Then HScroll1.Enabled = False
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

