VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CITF_CR 
   Caption         =   "CITF Cash Receipt Mapping"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   14985
      TabIndex        =   16
      Top             =   30
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   300
      TabIndex        =   17
      Top             =   9720
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9615
      Left            =   -30
      TabIndex        =   18
      Top             =   -105
      Width           =   15300
      Begin VB.Frame Frame0 
         Height          =   2160
         Left            =   570
         TabIndex        =   19
         Top             =   4965
         Visible         =   0   'False
         Width           =   14775
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
            ItemData        =   "FrmTxn_CITF_CR.frx":0000
            Left            =   6240
            List            =   "FrmTxn_CITF_CR.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   62
            Top             =   457
            Width           =   2385
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   390
            Left            =   10230
            TabIndex        =   22
            Top             =   435
            Width           =   1065
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   390
            Left            =   9150
            TabIndex        =   21
            Top             =   435
            Width           =   1065
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
            Left            =   2910
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   457
            Width           =   3225
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   1455
            Left            =   105
            TabIndex        =   23
            Top             =   1005
            Width           =   14115
            _ExtentX        =   24897
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
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Begin MSComCtl2.DTPicker STo 
            Height          =   345
            Left            =   1507
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   465
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
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
            Format          =   80281601
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   345
            Left            =   105
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   465
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
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
            Format          =   80281601
            CurrentDate     =   36494
         End
         Begin VB.Label Label11 
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
            Left            =   7192
            TabIndex        =   63
            Top             =   165
            Width           =   480
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
            TabIndex        =   30
            Top             =   165
            Width           =   945
         End
         Begin VB.Label Label36 
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
            Left            =   4140
            TabIndex        =   29
            Top             =   165
            Width           =   765
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
            Left            =   1815
            TabIndex        =   28
            Top             =   165
            Width           =   735
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
            Left            =   13830
            TabIndex        =   27
            Top             =   510
            Width           =   90
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
            Left            =   11385
            TabIndex        =   26
            Top             =   503
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   9255
         Left            =   120
         TabIndex        =   31
         Top             =   135
         Width           =   14850
         Begin VB.Frame Frame2 
            Height          =   2520
            Left            =   180
            TabIndex        =   44
            Top             =   735
            Width           =   14550
            Begin VB.TextBox TxtBOE 
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
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   480
               Width           =   1785
            End
            Begin VB.CommandButton CmdCashReceipt 
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
               Left            =   1155
               TabIndex        =   59
               Top             =   1260
               Width           =   375
            End
            Begin VB.CommandButton CmdCITFNo 
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
               Left            =   1125
               TabIndex        =   58
               Top             =   165
               Width           =   375
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
               Height          =   405
               Left            =   12525
               TabIndex        =   6
               Top             =   2025
               Width           =   1755
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
               Height          =   405
               Left            =   10770
               TabIndex        =   5
               Top             =   2025
               Width           =   1755
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
               Height          =   375
               Left            =   1080
               MaxLength       =   1000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Top             =   2040
               Width           =   9615
            End
            Begin VB.TextBox TxtCITFNo 
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
               Left            =   75
               TabIndex        =   2
               Top             =   480
               Width           =   1425
            End
            Begin VB.TextBox TxtCRNo 
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
               Left            =   105
               TabIndex        =   3
               Top             =   1560
               Width           =   1425
            End
            Begin VB.TextBox TxtCITFDate 
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
               Left            =   1530
               Locked          =   -1  'True
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   480
               Width           =   1545
            End
            Begin VB.TextBox TxtCITFAgent 
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
               Left            =   4935
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   480
               Width           =   8925
            End
            Begin VB.TextBox TxtCRAmt 
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
               Left            =   3135
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   1560
               Width           =   1785
            End
            Begin VB.TextBox TxtCRDate 
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
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1560
               Width           =   1545
            End
            Begin VB.TextBox TxtCRAgent 
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
               Left            =   4965
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1560
               Width           =   8925
            End
            Begin VB.Label LblRRNQty 
               AutoSize        =   -1  'True
               Caption         =   "BOE Amount"
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
               Left            =   3442
               TabIndex        =   61
               Top             =   195
               Width           =   1140
            End
            Begin VB.Label Label42 
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
               Left            =   75
               TabIndex        =   57
               Top             =   2100
               Width           =   840
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Cash Receipt Amount"
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
               Left            =   3255
               TabIndex        =   56
               Top             =   990
               Width           =   1545
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "CITF Date"
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
               Left            =   1845
               TabIndex        =   55
               Top             =   195
               Width           =   915
            End
            Begin VB.Label LblISINNo 
               AutoSize        =   -1  'True
               Caption         =   "CITF No."
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
               Left            =   240
               TabIndex        =   54
               Top             =   195
               Width           =   795
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               Caption         =   "Cash Receipt No."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   75
               TabIndex        =   53
               Top             =   990
               Width           =   1230
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Caption         =   "Cash Receipt Date"
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
               Left            =   1575
               TabIndex        =   52
               Top             =   990
               Width           =   1515
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "CITF Agent"
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
               Left            =   8895
               TabIndex        =   51
               Top             =   195
               Width           =   1005
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "Cash Receipt Agent"
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
               Left            =   8797
               TabIndex        =   50
               Top             =   990
               Width           =   1260
            End
         End
         Begin VB.Frame ButtonFrm 
            Height          =   1125
            Left            =   180
            TabIndex        =   43
            Top             =   8025
            Width           =   14625
            Begin VB.CommandButton NextCmd 
               Caption         =   "&Next"
               Height          =   400
               Left            =   120
               TabIndex        =   11
               Top             =   600
               Width           =   2865
            End
            Begin VB.CommandButton SaveCmd 
               Caption         =   "Save"
               Height          =   400
               Left            =   3003
               TabIndex        =   7
               Top             =   195
               Width           =   2865
            End
            Begin VB.CommandButton PreviousCmd 
               Caption         =   "&Previous"
               Height          =   400
               Left            =   3003
               TabIndex        =   12
               Top             =   600
               Width           =   2865
            End
            Begin VB.CommandButton EditCmd 
               Caption         =   "&Edit"
               Height          =   400
               Left            =   5886
               TabIndex        =   8
               Top             =   195
               Width           =   2865
            End
            Begin VB.CommandButton FirstCmd 
               Caption         =   "&First"
               Height          =   400
               Left            =   5886
               TabIndex        =   13
               Top             =   600
               Width           =   2865
            End
            Begin VB.CommandButton DeleteCmd 
               Caption         =   "&Delete"
               Height          =   400
               Left            =   8769
               TabIndex        =   9
               Top             =   195
               Width           =   2865
            End
            Begin VB.CommandButton LastCmd 
               Caption         =   "&Last"
               Height          =   400
               Left            =   8769
               TabIndex        =   14
               Top             =   600
               Width           =   2865
            End
            Begin VB.CommandButton SearchCmd 
               Caption         =   "Search"
               Height          =   400
               Left            =   11655
               TabIndex        =   10
               Top             =   195
               Width           =   2865
            End
            Begin VB.CommandButton ExitCmd 
               Caption         =   "E&xit"
               Height          =   400
               Left            =   11655
               TabIndex        =   15
               Top             =   600
               Width           =   2865
            End
            Begin VB.CommandButton AddCmd 
               Caption         =   "&Add New"
               Height          =   400
               Left            =   120
               TabIndex        =   0
               Top             =   195
               Width           =   2865
            End
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
            Left            =   9210
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   7560
            Width           =   5475
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
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   7560
            Width           =   5205
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
            Left            =   7320
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   300
            Width           =   3630
         End
         Begin VB.TextBox TxtCITFNoTxn 
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
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   300
            Width           =   1815
         End
         Begin VB.TextBox TxtCITFDispalyDate 
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
            Left            =   4245
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   300
            Width           =   1815
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   3900
            Left            =   180
            TabIndex        =   38
            Top             =   3315
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   6879
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
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Begin VB.TextBox TxtLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   300
            Width           =   3630
         End
         Begin VB.Label Label6 
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
            Left            =   8085
            TabIndex        =   42
            Top             =   7620
            Width           =   975
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Left            =   150
            TabIndex        =   41
            Top             =   7620
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Txn ID"
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
            Left            =   180
            TabIndex        =   36
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Txn Date"
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
            Left            =   3225
            TabIndex        =   35
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label10 
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
            Left            =   6270
            TabIndex        =   34
            Top             =   360
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_CITF_CR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mSM_Prefix, mLocationName    As Variant
Dim mLocationID, mSLocationId As Double
Dim mCITFTxnNo  As Double
Dim mFlag As Variant
Private Sub AddCmd_Click()

Edit_Flg = "A"
mSM_Prefix = ""
Frame2.Enabled = True
CmdCITFNo.Visible = True
CmdCashReceipt.Visible = True
Call GButtonLock(Me, False)

Call Grid_Head
TxtLocation.Visible = False
CmbLocation.Visible = True
CmbLocation.Enabled = True
CmbLocation.Text = ""
'CmbLocation.Visible = True
TxtCITFNoTxn.Text = ""
TxtCITFDispalyDate.Text = Date 'Format(Now, "MM/dd/yyyy")
TxtCreated = ""
TxtUpdated = ""
mFlag = "N"
'Call ClearFrame

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
   MsgBox "No Record found !!!! "
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
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
Call ClearFrame
If tmp <> CmbLocation.Text Then
    Call Grid_Head
    Call GetCITFNo
    Call GetCRNo
End If

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
mSLocationId = RsMst_Location!LocationID
End Sub

Private Sub CmdAdd2List_Click()
Dim r As Variant

If TxtCITFNo.Text = "" Then
   MsgBox "Blank CITF No. not allowed!!!"
   TxtCITFNo.SetFocus
   Exit Sub
End If
If TxtCRNo.Text = "" Then
   MsgBox "Blank Cash Receipt No. not allowed!!!"
   TxtCRNo.SetFocus
   Exit Sub
End If
If CheckDuplicateCITF = True Then
   MsgBox "Duplicate 'CITF No.'  and  'Cash Receipt No.' not allowed."
   TxtCITFNo.SetFocus
   Exit Sub
End If
'If CheckDuplicateCRID = True Then
'   MsgBox "Duplicate CRID Not Allowed."
'   TxtCRNo.SetFocus
'   Exit Sub
'End If

If CheckCITFNo(Val(TxtCITFNo.Text), Val(TxtCRNo.Text)) = True Then
   MsgBox "This 'CITF No.'  and  'Cash Receipt No.' is already mapped in Transaction No. '" & mCITFTxnNo & "' !!!"
   TxtCITFNo.SetFocus
   Exit Sub
End If
'If CheckCRID(Val(TxtCRNo.Text)) = True Then
'   MsgBox "This CRID is already exists !!!"
'   TxtCRNo.SetFocus
'   Exit Sub
'End If
If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid1.RowSel
Else
   r = MSHFlexGrid1.Rows - 1
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   MSHFlexGrid1.TextMatrix(r, 0) = ""
End If


MSHFlexGrid1.TextMatrix(r, 0) = TxtCITFNo.Text
MSHFlexGrid1.TextMatrix(r, 1) = TxtCITFDate.Text
MSHFlexGrid1.TextMatrix(r, 2) = Val(TxtBOE)
MSHFlexGrid1.TextMatrix(r, 3) = Trim(TxtCITFAgent)
MSHFlexGrid1.TextMatrix(r, 4) = TxtCRNo.Text
MSHFlexGrid1.TextMatrix(r, 5) = TxtCRDate.Text
MSHFlexGrid1.TextMatrix(r, 6) = Val(TxtCRAmt)
MSHFlexGrid1.TextMatrix(r, 7) = Trim(TxtCRAgent)
MSHFlexGrid1.TextMatrix(r, 8) = TxtRemarks
Call ClearFrame
TxtCITFNo.SetFocus

End Sub

Private Sub CmdCashReceipt_Click()
FrmTxn_NewCashReceipt.Show
End Sub

Private Sub CmdCITFNo_Click()
FrmWithdrawal.Show
End Sub

Private Sub CmdDeleteList_Click()
I = MSHFlexGrid1.RowSel
If I > 0 Then
   'If MSHFlexGrid1.TextMatrix(I, 1) <> "" Then
   If Trim(TxtCITFNo.Text) <> "" Then
      For RO = I To MSHFlexGrid1.Rows - 2
          For c = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(RO, c) = MSHFlexGrid1.TextMatrix(RO + 1, c)
          Next
      Next
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
   Else
      MsgBox "Select record from grid to delete!!!"
   End If
End If

Call ClearFrame
End Sub

Private Sub ClearFrame()
TxtCITFNo.Text = ""
TxtCITFDate.Text = ""
TxtBOE.Text = ""
TxtCITFAgent.Text = ""
TxtCRNo.Text = ""
TxtCRDate.Text = ""
TxtCRAmt.Text = ""
TxtCRAgent.Text = ""
TxtRemarks.Text = ""
CmdAdd2List.Caption = "Add to List"
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CITFCR.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CITFCR.xls")
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

Private Sub CmdGo_Click()
Dim Wc As Variant

Call Grid_Head2

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

'wc = ""
'RsTxn_CITFCR.Filter = ""

tmp = "Select TC.CITFCRID, TC.TxnDate, ML.LocationName,TC.Flag " & _
      " From Txn_CITFCR TC " & _
      " INNER Join Mst_Location ML On TC.LocationID = ML.LocationID "


If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TxnDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
'   wc = " TxnDate >= #" & Format(SFrom, Gen_DatFormat) & "# And TxnDate <= #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where TxnDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
'   wc = " TxnDate >= '" & Format(SFrom, Gen_DatFormat) & "' And TxnDate <= '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " AND ML.LocationName = '" & CmbSLocation.Text & "' "
'   wc = wc & " And LocationID = " & mSLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " AND ML." & Gen_WcLocation '.LocationName = '" & CmbSLocation.Text & "' "
'      wc = wc & " And ML." & Gen_WcLocation 'LocationID = " & mSLocationID & ""
   End If
End If

If CmbSType.Text <> "" Then
   tmp = tmp & " And TC.Flag = '" & Left(CmbSType.Text, 1) & " '"
End If

Call TmpTable

Screen.MousePointer = vbHourglass

For iu = 1 To TmpRs.RecordCount
    For c = 0 To MSHFlexGrid2.Cols - 1
        MSHFlexGrid2.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
        If InStr(MSHFlexGrid2.TextMatrix(iu, c), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(iu, c)) Then
              MSHFlexGrid2.TextMatrix(iu, c) = Format(MSHFlexGrid2.TextMatrix(iu, c), "dd-mmm-yyyy")
           Else
              MSHFlexGrid2.TextMatrix(iu, c) = MSHFlexGrid2.TextMatrix(iu, c)
           End If
        Else
          MSHFlexGrid2.TextMatrix(iu, c) = MSHFlexGrid2.TextMatrix(iu, c)
        End If
    Next
    If LCase(MSHFlexGrid2.TextMatrix(iu, 3)) = "o" Then
       MSHFlexGrid2.TextMatrix(iu, 3) = "Old"
    ElseIf LCase(MSHFlexGrid2.TextMatrix(iu, 3)) = "n" Then
       MSHFlexGrid2.TextMatrix(iu, 3) = "New"
    End If
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    TmpRs.MoveNext
Next


Label22.Refresh
Label22.Caption = TmpRs.RecordCount
Screen.MousePointer = vbNormal

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If ChkCITFPrint = True Then
   MsgBox "One or more CITF is printed . Can't Delete!!!"
   Exit Sub
End If


ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   Call DeleteChildTableDetail
   
   RsTxn_CITFCR.Delete
   RsTxn_CITFCR.Update
   'MsgBox "Record Successfully Removed"
   If RsTxn_CITFCR.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CITFCR.MoveNext
   If RsTxn_CITFCR.EOF() Then
      RsTxn_CITFCR.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_CITFCR.CancelUpdate
End Sub
Private Function ChkCITFPrint() As Boolean
Dim RetVal As Boolean
RetVal = False

For I = 1 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(I, 9)) = "p" Then
       RetVal = True
       Exit For
    End If
Next

ChkCITFPrint = RetVal

End Function

Sub DeleteChildTableDetail()
tmp = "Delete from Txn_CITFCRDetail where CITFCRID = " & Val(TxtCITFNoTxn.Text) & ""
Call TmpTable
End Sub

Private Sub EditCmd_Click()

Edit_Flg = "E"
Frame2.Enabled = True
Call GButtonLock(Me, False)
CmdCITFNo.Visible = True
CmdCashReceipt.Visible = True
'TxtLocation.Visible = False
'CmbLocation.Visible = True
CmbLocation.Text = TxtLocation
'CmbLocation.Visible = True

'Call TableTxn_CxTF_Details(" WHERE TxnType = 'W' AND SM_Prefix = '" & mSM_Prefix & "'")
Call GetCITFNo
'Call TableTxn_CashReceipt(" WHERE SM_Prefix = '" & mSM_Prefix & "'")
Call GetCRNo

TxtCITFNo.Locked = False
TxtCRNo.Locked = False
TxtRemarks.Locked = False
Call ClearFrame


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim Wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)

TxtLocation.Visible = True
CmbLocation.Enabled = False
Frame2.Enabled = False

If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

CmbLocation.Enabled = True
Call Grid_Head

If Gen_WcLocation = "" Then
   Call TableTxn_CITFCR(" Order by CITFCRID")
Else
   Call TableTxn_CITFCR("  Where " & Gen_WcLocation & " Order by CITFCRID")
End If

If RsTxn_CITFCR.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Frame2.Enabled = True

Call Indata
Call GButtonLock(Me, True)
End Sub


Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid1_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid1.RowSel

If r < 1 Then Exit Sub

If MSHFlexGrid1.TextMatrix(r, 0) = "" Then
   Call ClearFrame
   Exit Sub
End If

If LCase(MSHFlexGrid1.TextMatrix(r, 9)) = "p" Then
   MsgBox "CITF already printed. Can't edit row!!!"
   Exit Sub
End If

TxtCITFNo.Text = MSHFlexGrid1.TextMatrix(r, 0)
TxtCITFDate.Text = MSHFlexGrid1.TextMatrix(r, 1)
TxtBOE = Val(MSHFlexGrid1.TextMatrix(r, 2))
TxtCITFAgent.Text = MSHFlexGrid1.TextMatrix(r, 3)
TxtCRNo.Text = MSHFlexGrid1.TextMatrix(r, 4)
TxtCRDate.Text = MSHFlexGrid1.TextMatrix(r, 5)
TxtCRAmt = Trim(MSHFlexGrid1.TextMatrix(r, 6))
TxtCRAgent = Trim(MSHFlexGrid1.TextMatrix(r, 7))
TxtRemarks = Trim(MSHFlexGrid1.TextMatrix(r, 8))
CmdAdd2List.Caption = "Update List"
TxtCITFNo.SetFocus

End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_CITFCR.MoveFirst
RsTxn_CITFCR.Find "CITFCRID = '" & MSHFlexGrid2.TextMatrix(I, 0) & "'"
If Not RsTxn_CITFCR.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CITFCR.MoveNext
If RsTxn_CITFCR.EOF Then
   RsTxn_CITFCR.MoveLast
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
RsTxn_CITFCR.MovePrevious
If RsTxn_CITFCR.BOF Then
   RsTxn_CITFCR.MoveFirst
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
RsTxn_CITFCR.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsTxn_CITFCR.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Public Sub Indata()

Call ClearFrame
Frame2.Enabled = False
CmdCITFNo.Visible = False
CmdCashReceipt.Visible = False
Call Grid_Head

mLocationID = 0
TxtCITFNoTxn.Text = ""
TxtCITFDispalyDate.Text = Format(Now, "MM/dd/yyyy")
TxtLocation.Visible = True
CmbLocation.Visible = False


TxtCITFNoTxn = RsTxn_CITFCR!CITFCRID
TxtCITFDispalyDate.Text = RsTxn_CITFCR!TxnDate
mSM_Prefix = IIf(IsNull(RsTxn_CITFCR!SM_Prefix), "", RsTxn_CITFCR!SM_Prefix)
mLocationID = IIf(IsNull(RsTxn_CITFCR!LocationID), 0, RsTxn_CITFCR!LocationID)
TxtLocation.Text = GetLocationName(RsTxn_CITFCR!LocationID)

mFlag = IIf(IsNull(RsTxn_CITFCR!Flag), "O", RsTxn_CITFCR!Flag)

TxtCreated = IIf(IsNull(RsTxn_CITFCR!CreatedBy), "", RsTxn_CITFCR!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CITFCR!CreatedDate), "", RsTxn_CITFCR!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CITFCR!UpdatedBy), "", RsTxn_CITFCR!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CITFCR!UpdatedDate), "", RsTxn_CITFCR!UpdatedDate)

Call InDataCITFCRDetail

'If RsTxn_CITFCR.RecordCount = 1 Then
'   Call GButtonLock_1(Me, True)
'Else
Call GButtonLock(Me, True)
'End If

End Sub

Private Sub InDataCITFCRDetail()

Call Grid_Head
Call TableTxn_CITFCRDetail(" Where CITFCRID= " & Val(TxtCITFNoTxn.Text) & " ") 'AND SM_Prefix = '" & TxtPreFix & "'")

Call GetCITFNo
Call GetCRNo
'tmp = "SELECT TCD.CxTFNo, TCD.CxTFDate, TCB.BOEAmount, MA.AgentName FROM  Txn_CxTF_Details TCD"
'tmp = tmp & " LEFT Join Txn_CxTF_BOE TCB On TCB.CxTFNo = TCD.CxTFNo AND TCD.SM_Prefix = TCB.SM_Prefix"
'tmp = tmp & " LEFT Join Mst_Agent MA On MA.AgentID = TCD.AgentID"
'tmp = tmp & " WHERE TCD.TxnType = 'W' AND TCD.SM_Prefix = '" & mSM_Prefix & "'"
'Call Tmp1Table

'tmp = "SELECT TCR.CRID, TCR.CRDate, TCR.InstrumentAmount, MA.AgentName FROM  Txn_CashReceipt TCR"
'tmp = tmp & " INNER Join Mst_Agent MA On MA.AgentID = TCR.AgentID"
'tmp = tmp & " WHERE TCR.SM_Prefix = '" & mSM_Prefix & "'"
'Call Tmp2Table

If RsTxn_CITFCRDetail.RecordCount > 0 Then
   For I = 1 To RsTxn_CITFCRDetail.RecordCount
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
       End If
       TmpRs1.Find "CxTFNo = " & Val(RsTxn_CITFCRDetail!CITFNo)
       
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(I, 0) = RsTxn_CITFCRDetail!CITFNo
          MSHFlexGrid1.TextMatrix(I, 1) = Format(TmpRs1!CxtfDate, "dd-MMM-yyyy")
          MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!BOEAmount), "0", TmpRs1!BOEAmount)
          MSHFlexGrid1.TextMatrix(I, 3) = TmpRs1!AgentName
       End If
       
       MSHFlexGrid1.TextMatrix(I, 4) = RsTxn_CITFCRDetail!CRID
       If TmpRs2.RecordCount > 0 Then
          TmpRs2.MoveFirst
       End If
       TmpRs2.Find "CRID = " & Val(RsTxn_CITFCRDetail!CRID)
        
       If Not TmpRs2.EOF Then
          TmpRs2.MoveFirst
          TmpRs2.Find "CRID = " & Val(RsTxn_CITFCRDetail!CRID)
          If Not TmpRs2.EOF Then
              MSHFlexGrid1.TextMatrix(I, 5) = Format(TmpRs2!CRDate, "dd-MMM-yyyy")
              MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!InstrumentAmount), "0", TmpRs2!InstrumentAmount)
              If LCase(mFlag) = "o" Then
                 MSHFlexGrid1.TextMatrix(I, 7) = TmpRs2!AgentName
              ElseIf LCase(mFlag) = "n" Then
                 MSHFlexGrid1.TextMatrix(I, 7) = ""
              End If
          End If
       End If
       MSHFlexGrid1.TextMatrix(I, 8) = RsTxn_CITFCRDetail!Remarks
       MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
       RsTxn_CITFCRDetail.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Next
End If

End Sub

Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If CmbLocation.Text = "" Then
      MsgBox "Blank Location not allowed!!!"
      CmbLocation.SetFocus
      Exit Sub
   End If
End If

If MSHFlexGrid1.Rows < 3 Then
   MsgBox "Map at least one CITF No. and Cash Receipt No. !!!"
   TxtCITFNo.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_CITFCR.AddNew
   RsTxn_CITFCR!CITFCRID = GetMaxCITFCRID
   RsTxn_CITFCR!G_UID = GetGUID
   TxtCITFNoTxn.Text = RsTxn_CITFCR!CITFCRID
   RsTxn_CITFCR!TxnDate = TxtCITFDispalyDate.Text
   RsTxn_CITFCR!SM_Prefix = mSM_Prefix
   RsTxn_CITFCR!LocationID = mLocationID
   RsTxn_CITFCR!Flag = mFlag
Else
   RsTxn_CITFCR.MoveFirst
   RsTxn_CITFCR.Find "CITFCRID= '" & TxtCITFNoTxn.Text & "'"
End If


If IsNull(RsTxn_CITFCR!CreatedBy) = True Or RsTxn_CITFCR!CreatedBy = "" Then
   RsTxn_CITFCR!CreatedBy = Gen_UserLoginID
   RsTxn_CITFCR!CreatedDate = Now
Else
   RsTxn_CITFCR!UpdatedBy = Gen_UserLoginID
   RsTxn_CITFCR!UpdatedDate = Now
End If

RsTxn_CITFCR.Update

Call SaveCITFCRDetail

TxtCreated = IIf(IsNull(RsTxn_CITFCR!CreatedBy), "", RsTxn_CITFCR!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CITFCR!CreatedDate), "", RsTxn_CITFCR!CreatedDate)
If Edit_Flg = "E" Then
   TxtUpdated = IIf(IsNull(RsTxn_CITFCR!UpdatedBy), "", RsTxn_CITFCR!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CITFCR!UpdatedDate), "", RsTxn_CITFCR!UpdatedDate)
End If

RsTxn_CITFCR.Filter = ""
RsTxn_CITFCR.Requery
RsTxn_CITFCR.MoveFirst
RsTxn_CITFCR.Find "CITFCRID = '" & TxtCITFNoTxn.Text & "'"

Call Indata

End Sub

Private Sub SaveCITFCRDetail()
Call DeleteCITFCRDetail
Call TableTxn_CITFCRDetail("") '(" Where ClaimID = " & TxtClaimID & " AND SM_Prefix = '" & TxtPreFix & "'")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       RsTxn_CITFCRDetail.AddNew
       RsTxn_CITFCRDetail!CITFCRDetailID = GetMaxCITFCRIDDetails
       RsTxn_CITFCRDetail!CITFCRID = TxtCITFNoTxn.Text
       RsTxn_CITFCRDetail!SM_Prefix = mSM_Prefix
       RsTxn_CITFCRDetail!CITFNo = MSHFlexGrid1.TextMatrix(I, 0)
       RsTxn_CITFCRDetail!CRID = Val(MSHFlexGrid1.TextMatrix(I, 4))
'       MSHFlexGrid1.TextMatrix(I, 8) = Replace(MSHFlexGrid1.TextMatrix(I, 8), ",", "-")
       RsTxn_CITFCRDetail!Remarks = MSHFlexGrid1.TextMatrix(I, 8)
       RsTxn_CITFCRDetail!G_UID = GetGUID
       RsTxn_CITFCRDetail!CreatedBy = Gen_UserLoginID
       RsTxn_CITFCRDetail!CreatedDate = Now
       RsTxn_CITFCRDetail.Update
       RsTxn_CITFCRDetail.Requery
    End If
Next
End Sub
Private Sub DeleteCITFCRDetail()
tmp = "Delete From Txn_CITFCRDetail Where CITFCRID=" & TxtCITFNoTxn.Text & " "
Call Tmp4Table
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CITFCR.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_CITFCR.MoveFirst
 
   RsTxn_CITFCR.Find "G_UID = '" & TxtG_UID & "'"
   If RsTxn_CITFCR.EOF Then
      RsTxn_CITFCR.MoveLast
   End If
   Call Indata
   Exit Sub
End If
'
'Frame0.Height = Frame1.Height + 800 'Me.Height - 500
'Frame0.Left = Frame4.Left + 100
'Frame0.Top = Frame1.Top
'Frame0.Width = Frame1.Width + 100 'Me.Width - 200


Frame0.Height = Frame4.Height - 100
Frame0.Left = Frame1.Left '- 100
Frame0.Top = Frame1.Top '- 100
Frame0.Width = Frame4.Width - 350

Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 200
MSHFlexGrid2.Height = Frame0.Height - 1400

RsTxn_CITFCR.MoveFirst

Call Grid_Head2

SFrom.Value = CDate(From_GenDate) - 900
STo.Value = Date 'To_GenDate
CmbSLocation.Text = ""
Label22.Caption = ""

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 10


MSHFlexGrid1.TextMatrix(0, 0) = "CITF No"
MSHFlexGrid1.TextMatrix(0, 1) = "CITF Date"
MSHFlexGrid1.TextMatrix(0, 2) = "BOE Amt"
MSHFlexGrid1.TextMatrix(0, 3) = "CITF Agent"
MSHFlexGrid1.TextMatrix(0, 4) = "Cash Receipt No."
MSHFlexGrid1.TextMatrix(0, 5) = "Cash Receipt Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Cash Receipt Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Cash Receipt Agent"
MSHFlexGrid1.TextMatrix(0, 8) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 9) = "CITFPrint"



MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1100
MSHFlexGrid1.ColWidth(3) = 2300
MSHFlexGrid1.ColWidth(4) = 800
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 2350
MSHFlexGrid1.ColWidth(8) = 3300
MSHFlexGrid1.ColWidth(9) = 0

MSHFlexGrid1.RowHeight(0) = 750

End Sub

Private Sub Grid_Head2()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.Cols = 4


MSHFlexGrid2.TextMatrix(0, 0) = "Txn No"
MSHFlexGrid2.TextMatrix(0, 1) = "Txn Date"
MSHFlexGrid2.TextMatrix(0, 2) = "Location"
MSHFlexGrid2.TextMatrix(0, 3) = "Flag"

MSHFlexGrid2.ColWidth(0) = 1200
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 2500
MSHFlexGrid2.ColWidth(3) = 1000

MSHFlexGrid2.RowHeight(0) = 600

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
VScroll1.Max = Me.ScaleHeight - Frame1.Height '4
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame1.Width '4
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame1.Height >= Me.Height Then VScroll1.Enabled = True ''4
If Frame1.Height < Me.Height Then VScroll1.Enabled = False '4
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True '4
If Frame1.Width < Me.Width Then HScroll1.Enabled = False '4
End Sub


Private Sub TxtCITFNo_GotFocus()
'Call TableTxn_CxTF_Details(" WHERE TxnType = 'W' AND SM_Prefix = '" & mSM_Prefix & "'")
End Sub

Private Sub TxtCITFNo_LostFocus()
If Trim(TxtCITFNo.Text) = "" Then
   TxtCITFDate.Text = ""
   TxtBOE.Text = ""
   TxtCITFAgent.Text = ""
   Exit Sub
End If
If CmbLocation.Text = "" Then
    MsgBox "Select Location !!!"
    CmbLocation.SetFocus
    Exit Sub
End If

If IsNumeric(TxtCITFNo.Text) = False Or Val(TxtCITFNo.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtCITFNo.SetFocus
   Exit Sub
End If
If TmpRs1.RecordCount > 0 Then
   TmpRs1.MoveFirst
   TmpRs1.Find "CxTFNo = " & Val(TxtCITFNo.Text)
   If Not TmpRs1.EOF Then
      TxtCITFDate.Text = TmpRs1!CxtfDate
      TxtBOE.Text = IIf(IsNull(TmpRs1!BOEAmount), "0", TmpRs1!BOEAmount)
      TxtCITFAgent.Text = TmpRs1!AgentName
   Else
      TxtCITFDate.Text = ""
      TxtBOE.Text = ""
      TxtCITFAgent.Text = ""
      MsgBox "No Record found !!!! "
      TxtCITFNo.SetFocus
   End If
Else
   TxtCITFDate.Text = ""
   TxtBOE.Text = ""
   TxtCITFAgent.Text = ""
   MsgBox "No Record found !!!! "
   TxtCITFNo.SetFocus
End If
End Sub

Private Sub TxtCRNo_GotFocus()
'Call TableTxn_CashReceipt(" WHERE SM_Prefix = '" & mSM_Prefix & "'")
End Sub

Private Sub TxtCRNo_LostFocus()
If Trim(TxtCRNo.Text) = "" Then
   TxtCRDate.Text = ""
   TxtCRAmt.Text = ""
   TxtCRAgent.Text = ""
   Exit Sub
End If
If CmbLocation.Text = "" Then
    MsgBox "Select Location !!!"
    CmbLocation.SetFocus
    Exit Sub
End If

If IsNumeric(TxtCRNo.Text) = False Or Val(TxtCRNo.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtCRNo.SetFocus
   Exit Sub
End If

If TmpRs2.RecordCount > 0 Then
   TmpRs2.MoveFirst
   TmpRs2.Find "CRID = " & Val(TxtCRNo.Text)
   If Not TmpRs2.EOF Then
      TxtCRDate.Text = TmpRs2!CRDate
      TxtCRAmt.Text = IIf(IsNull(TmpRs2!InstrumentAmount), "0", TmpRs2!InstrumentAmount)
      If LCase(mFlag) = "o" Then
         TxtCRAgent.Text = TmpRs2!AgentName
      ElseIf LCase(mFlag) = "n" Then
         TxtCRAgent.Text = "" 'TmpRs2!AgentName
      End If
   Else
      TxtCRDate.Text = ""
      TxtCRAmt.Text = ""
      TxtCRAgent.Text = ""
      MsgBox "No Record found !!!! "
      TxtCRNo.SetFocus
   End If
Else
   TxtCRDate.Text = ""
   TxtCRAmt.Text = ""
   TxtCRAgent.Text = ""
   MsgBox "No Record found !!!! "
   TxtCRNo.SetFocus
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

Private Function GetMaxCITFCRID() As Double
Dim RetVal As Double
tmp = "Select max(CITFCRID) from Txn_CITFCR "
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetMaxCITFCRID = RetVal
End Function

Function GetMaxCITFCRIDDetails() As Double
Dim RetVal As Double
RetVal = 0
tmp = "Select max(CITFCRDetailID) from Txn_CITFCRDetail"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
RetVal = RetVal + 1
TmpRs.Close
GetMaxCITFCRIDDetails = RetVal

End Function

Function CheckDuplicateCITF() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid1.Rows > 1 Then
   For u = 1 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(u, 0) = Trim(TxtCITFNo.Text) And MSHFlexGrid1.TextMatrix(u, 4) = Trim(TxtCRNo.Text) Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             RetVal = True
             Exit For
          Else
             If u <> MSHFlexGrid1.RowSel Then
                RetVal = True
             End If
          End If
       End If
   Next
End If
CheckDuplicateCITF = RetVal
End Function

Function CheckDuplicateCRID() As Boolean
'Dim retval As Boolean
'retval = False
'If MSHFlexGrid1.Rows > 1 Then
'   For u = 1 To MSHFlexGrid1.Rows - 1
'       If MSHFlexGrid1.TextMatrix(u, 4) = Trim(TxtCRNo.Text) Then
'          If LCase(CmdAdd2List.Caption) = "add to list" Then
'             retval = True
'             Exit For
'          Else
'             If u <> MSHFlexGrid1.RowSel Then
'                retval = True
'             End If
'          End If
'       End If
'   Next
'End If
'CheckDuplicateCRID = retval
End Function

Function CheckCITFNo(CITF_ As Double, CRID_ As Double) As Boolean
Dim RetVal As Boolean

RetVal = False
tmp = "Select TC.CITFCRID,TCD.CITFNo,TCD.CRID From Txn_CITFCR TC "
tmp = tmp & " INNER Join Txn_CITFCRDetail TCD On TCD.CITFCRID = TC.CITFCRID "
tmp = tmp & " Where TC.SM_Prefix = '" & mSM_Prefix & "' AND TCD.CITFNo = " & CITF_ & " AND TCD.CRID = " & CRID_
Call TmpTable
mCITFTxnNo = 0
If TmpRs.RecordCount > 0 Then
   If TmpRs.Fields(0) <> TxtCITFNoTxn.Text Then
       RetVal = True
       mCITFTxnNo = TmpRs.Fields("CITFCRID")
   End If
End If
CheckCITFNo = RetVal
End Function

Function CheckCRID(CRID_ As Double) As Boolean
Dim RetVal As Boolean

'retval = False
'tmp = "Select TC.CITFCRID,CRID From Txn_CITFCR TC " & _
'      " INNER Join Txn_CITFCRDetail TCD On TC.CITFCRID = TCD.CITFCRID "
'tmp = tmp & " Where TC.SM_Prefix = '" & mSM_Prefix & "' AND TCD.CRID = " & CRID_
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   If TmpRs.Fields(0) <> TxtCITFNoTxn.Text Then
'       retval = True
'   End If
'End If
'CheckCRID = retval
End Function

Private Sub GetCITFNo()

tmp = "SELECT TCD.CxTFNo, TCD.CxTFDate, TCB.BOEAmount, MA.AgentName,TCD.Flag FROM  Txn_CxTF_Details TCD"
tmp = tmp & " LEFT Join Txn_CxTF_BOE TCB On TCB.CxTFNo = TCD.CxTFNo AND TCD.SM_Prefix = TCB.SM_Prefix"
tmp = tmp & " LEFT Join Mst_Agent MA On MA.AgentID = TCD.AgentID"
tmp = tmp & " WHERE TCD.TxnType = 'W' AND TCD.SM_Prefix = '" & mSM_Prefix & "'"

Call Tmp1Table

End Sub

Private Sub GetCRNo()

If LCase(mFlag) = "o" Then
   tmp = "SELECT TCR.CRID, TCR.CRDate, TCR.InstrumentAmount, MA.AgentName FROM  Txn_CashReceipt TCR"
   tmp = tmp & " INNER Join Mst_Agent MA On MA.AgentID = TCR.AgentID"
   tmp = tmp & " WHERE TCR.SM_Prefix = '" & mSM_Prefix & "'"
ElseIf LCase(mFlag) = "n" Then
   tmp = "SELECT TCR.CRID, TCR.CRDate, TCR.InstrumentAmount, '' FROM  Txn_NewCashReceipt TCR "
   'tmp = tmp & " INNER Join Mst_Agent MA On MA.AgentID = TCR.AgentID"
   tmp = tmp & " WHERE TCR.SM_Prefix = '" & mSM_Prefix & "' " 'And TCR.AdvanceFlag = 0 And"
End If


Call Tmp2Table

End Sub

