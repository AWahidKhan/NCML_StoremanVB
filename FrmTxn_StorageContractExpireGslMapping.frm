VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_StorageContractExpireGslMapping 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   9495
      Left            =   45
      TabIndex        =   22
      Top             =   30
      Width           =   14775
      Begin VB.Frame Frame0 
         Height          =   2775
         Left            =   840
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   12735
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   10485
            TabIndex        =   20
            Top             =   720
            Width           =   1005
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   345
            Left            =   9480
            TabIndex        =   19
            Top             =   720
            Width           =   1005
         End
         Begin VB.ComboBox CmbSClient 
            Height          =   315
            Left            =   5475
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   330
            Width           =   3525
         End
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   3195
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   330
            Width           =   2205
         End
         Begin VB.ComboBox CmbSCommodity 
            Height          =   315
            Left            =   9075
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   330
            Width           =   2325
         End
         Begin VB.TextBox TxtSearchContract 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4035
            TabIndex        =   18
            Top             =   720
            Width           =   5145
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1620
            Left            =   120
            TabIndex        =   24
            Top             =   1140
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   2858
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1665
            TabIndex        =   14
            Top             =   330
            Width           =   1410
            _ExtentX        =   2487
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
            Format          =   107675649
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   120
            TabIndex        =   13
            Top             =   330
            Width           =   1410
            _ExtentX        =   2487
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
            Format          =   107675649
            CurrentDate     =   36494
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "_"
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
            Left            =   13920
            TabIndex        =   32
            Top             =   765
            Width           =   135
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
            Left            =   11520
            TabIndex        =   31
            Top             =   765
            Width           =   2415
         End
         Begin VB.Label Label3 
            Caption         =   "Client"
            Height          =   225
            Left            =   6990
            TabIndex        =   30
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label42 
            Caption         =   "Location"
            Height          =   225
            Left            =   3930
            TabIndex        =   29
            Top             =   120
            Width           =   720
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1935
            TabIndex        =   28
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   375
            TabIndex        =   27
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label39 
            Caption         =   "Commodity"
            Height          =   210
            Left            =   9810
            TabIndex        =   26
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Description"
            Height          =   195
            Left            =   3075
            TabIndex        =   25
            Top             =   810
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4515
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   14535
         Begin VB.TextBox TxtContractID 
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
            Left            =   9615
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   217
            Width           =   2160
         End
         Begin VB.TextBox TxtBookedTo 
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
            Left            =   13275
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   635
            Width           =   1080
         End
         Begin VB.TextBox TxtBookedFrom 
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
            Left            =   13275
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   1080
         End
         Begin VB.CommandButton CmdStorageReservation 
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
            Left            =   7440
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   375
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1005
            Width           =   6000
         End
         Begin VB.TextBox TxtBookedRate 
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1320
         End
         Begin VB.TextBox TxtStorageCharges 
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1410
            Width           =   2850
         End
         Begin VB.TextBox TxtLocationID 
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
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   600
            Width           =   6000
         End
         Begin VB.Frame Frame2 
            Caption         =   "Mapped Godown Details"
            Height          =   2595
            Left            =   120
            TabIndex        =   42
            Top             =   1800
            Width           =   14295
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   2175
               Left            =   240
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   240
               Width           =   13905
               _ExtentX        =   24527
               _ExtentY        =   3836
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
         End
         Begin VB.TextBox TxtCommodityID 
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
            Left            =   9615
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   635
            Width           =   2160
         End
         Begin VB.TextBox TxtBookedQty 
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
            Left            =   9615
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1425
            Width           =   2130
         End
         Begin VB.TextBox TxtBookedAsPer 
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
            Left            =   9615
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1030
            Width           =   2130
         End
         Begin VB.TextBox TxtExtendedUpto 
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
            Left            =   13275
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1030
            Width           =   1080
         End
         Begin VB.TextBox TxtSMPrefix 
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
            Left            =   10800
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.ComboBox CmbStorageDescription 
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
            Top             =   210
            Width           =   5520
         End
         Begin VB.Label Label5 
            Caption         =   "Contract ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   8160
            TabIndex        =   63
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label Label2 
            Caption         =   "Reservation Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   1455
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
            Height          =   225
            Left            =   120
            TabIndex        =   60
            Top             =   1080
            Width           =   1800
         End
         Begin VB.Label LblLocationID 
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
            Height          =   225
            Left            =   120
            TabIndex        =   59
            Top             =   675
            Width           =   1800
         End
         Begin VB.Label LblCommodityID 
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
            Height          =   225
            Left            =   8160
            TabIndex        =   58
            Top             =   645
            Width           =   1200
         End
         Begin VB.Label LblBookedFrom 
            Caption         =   "Booked From"
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
            Left            =   11880
            TabIndex        =   57
            Top             =   270
            Width           =   1320
         End
         Begin VB.Label LblBookedTo 
            Caption         =   "Booked To"
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
            Left            =   11880
            TabIndex        =   56
            Top             =   660
            Width           =   1200
         End
         Begin VB.Label LblBookedAsPer 
            Caption         =   "Booked As Per"
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
            Left            =   8160
            TabIndex        =   55
            Top             =   1080
            Width           =   1425
         End
         Begin VB.Label LblBookedRate 
            Caption         =   "Booked Rate Per Month"
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
            Left            =   5190
            TabIndex        =   54
            Top             =   1365
            Width           =   1215
         End
         Begin VB.Label LblStorageCharges 
            Caption         =   "Storage Charges"
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
            Left            =   105
            TabIndex        =   53
            Top             =   1470
            Width           =   1800
         End
         Begin VB.Label LblBookedQty 
            Caption         =   "Booked Qty"
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
            Left            =   8160
            TabIndex        =   52
            Top             =   1455
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Extended Upto"
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
            Left            =   11880
            TabIndex        =   51
            Top             =   1065
            Width           =   1395
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1065
         Left            =   120
         TabIndex        =   35
         Top             =   8280
         Width           =   14415
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
            Left            =   105
            TabIndex        =   8
            Top             =   585
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
            Left            =   2940
            TabIndex        =   9
            Top             =   585
            Width           =   2835
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
            Left            =   11445
            TabIndex        =   7
            Top             =   180
            Width           =   2715
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
            Left            =   8610
            TabIndex        =   6
            Top             =   180
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
            Left            =   5775
            TabIndex        =   5
            Top             =   180
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
            Left            =   2940
            TabIndex        =   4
            Top             =   180
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
            Left            =   105
            TabIndex        =   3
            Top             =   180
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
            Left            =   11445
            TabIndex        =   12
            Top             =   585
            Width           =   2715
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
            Left            =   8610
            TabIndex        =   11
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
            Left            =   5775
            TabIndex        =   10
            Top             =   585
            Width           =   2835
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3570
         Left            =   120
         TabIndex        =   33
         Top             =   4725
         Width           =   14535
         Begin VB.CommandButton CmdGSLExcel 
            Caption         =   "GSL Excel"
            Height          =   350
            Left            =   13080
            TabIndex        =   66
            Top             =   132
            Width           =   1005
         End
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
            TabIndex        =   2
            Top             =   180
            Width           =   1320
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2865
            Left            =   120
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   600
            Width           =   14295
            _ExtentX        =   25215
            _ExtentY        =   5054
            _Version        =   393216
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "GSL mapped to another contract"
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
            Left            =   2880
            TabIndex        =   65
            Top             =   187
            Width           =   2880
         End
         Begin VB.Label Label6 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   2280
            TabIndex        =   64
            Top             =   180
            Width           =   375
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9975
      Left            =   15165
      TabIndex        =   21
      Top             =   30
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   9390
      Width           =   15015
   End
End
Attribute VB_Name = "FrmTxn_StorageContractExpireGslMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mLocationID, mSM_Prefix As Variant
Dim mCMPID, mGodownID, mStorageReservationDate, mMappingID, mExchnageTypeID As Variant
Dim mContractID As Double
Dim mClientIDRow As Double
Dim mSrearchfirst As Boolean

Public Sub Grid_Head()
With MSHFlexGrid2
   .Clear
   .Cols = 8
   .Rows = 3
   .FixedRows = 2
   .WordWrap = True
   
   .TextMatrix(1, 0) = "CMP"
   .TextMatrix(1, 1) = "Godown"
   .TextMatrix(1, 2) = "Unit Type"
   .TextMatrix(1, 3) = "Address"
   .TextMatrix(1, 4) = "Capacity in Bags"
   .TextMatrix(1, 5) = "Capacity in MT"
   .TextMatrix(1, 6) = "Reservation Qty"
   .TextMatrix(1, 7) = "GodownID"
   
   .ColWidth(0) = 2500
   .ColWidth(1) = 1400
   .ColWidth(2) = 1600
   .ColWidth(3) = 4500
   .ColWidth(4) = 1300
   .ColWidth(5) = 1200
   .ColWidth(6) = 1200
   .ColWidth(7) = 0
   
   .RowHeight(1) = 500
End With
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
CmbStorageDescription.Visible = True
CmbStorageDescription.Enabled = True
CmbStorageDescription.Text = ""
CmdStorageReservation.Visible = True
Frame3.Enabled = True
Check1.Enabled = True
Check1.Value = 0
TxtContractID.Text = ""
Call ClearContractDetails
Call ClearGodownDetails

Call Grid_Head
Call Grid_Head2
Call GButtonLock(Me, False)

End Sub

Private Sub Check1_Click()
If SaveCmd.Enabled = False Then Exit Sub
If Check1.Value = 1 Then
   MSHFlexGrid3.TextMatrix(0, 7) = 0 '- -Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(i, 7))
   MSHFlexGrid3.TextMatrix(0, 8) = 0 '- -Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(i, 8))

   For I = 2 To MSHFlexGrid3.Rows - 1
      If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
         MSHFlexGrid3.TextMatrix(I, 0) = "*"
         MSHFlexGrid3.TextMatrix(I, 11) = mContractID
         MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 7))
         MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(I, 8))
      End If
   Next
Else
   For I = 2 To MSHFlexGrid3.Rows - 1
      MSHFlexGrid3.TextMatrix(I, 0) = ""
      MSHFlexGrid3.TextMatrix(I, 11) = ""
      MSHFlexGrid3.TextMatrix(0, 7) = 0 'Val(MSHFlexGrid3.TextMatrix(0, 7)) - Val(MSHFlexGrid3.TextMatrix(i, 7))
      MSHFlexGrid3.TextMatrix(0, 8) = 0 'Val(MSHFlexGrid3.TextMatrix(0, 8)) - Val(MSHFlexGrid3.TextMatrix(i, 8))
   Next
End If
End Sub

Private Sub CmbStorageDescription_GotFocus()
tmp = CmbStorageDescription.Text

If Gen_WcLocation <> "" Then
   Call TableMst_StorageContract(" Where ContractID in(Select ContractID from Txn_SRGodownMapping) And " & Gen_WcLocation)  '(" Where CloseDate is null And LocationID = " & mLocationID)
Else
   Call TableMst_StorageContract(" Where ContractID in(Select ContractID from Txn_SRGodownMapping)")  '(" Where CloseDate is null And LocationID = " & mLocationID)
End If

CmbStorageDescription.Clear
If RsMst_StorageContract.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_StorageContract.RecordCount
   CmbStorageDescription.AddItem RsMst_StorageContract!ContractName
   RsMst_StorageContract.MoveNext
Next
CmbStorageDescription.Text = tmp
End Sub

Private Sub CmbStorageDescription_LostFocus()
TxtContractID.Text = ""
If CmbStorageDescription.Text = "" Then
   Call ClearContractDetails
   Call ClearGodownDetails
   mContractID = 0
   Exit Sub
End If

RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractName = '" & CmbStorageDescription.Text & "'"
If RsMst_StorageContract.EOF Then
   MsgBox "Invalid selection "
   CmbStorageDescription.SetFocus
   Exit Sub
End If
If mContractID <> RsMst_StorageContract!ContractID Then
   Call ClearContractDetails
   Call ClearGodownDetails
   Call Grid_Head
   Call Grid_Head2
End If
mContractID = RsMst_StorageContract!ContractID
TxtContractID.Text = mContractID
tmp = "Select SRGodownMappingID From Txn_SRGodownMapping Where ContractID=" & mContractID
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mMappingID = IIf(IsNull(TmpRs!SRGodownMappingID), 0, TmpRs!SRGodownMappingID)
End If
TxtSMPrefix = GetPreFixByLocation(RsMst_StorageContract!LocationID, "P")
Call GetContractDetails(mContractID)
Call Grid_Head
Call Indata_MappingDetails
Call FillGrid3
End Sub

Private Sub GetContractDetails(mContractID_ As Double)

'If CmbStorageDescription.Text = "" Then Exit Sub

tmp = "Select MSC.ContractID,MC.ClientName,ML.LocationID,ML.LocationName,MComm.Commodity,MSC.BookedFrom,MSC.BookedTo,MSC.BookedAsPer, "
tmp = tmp & " MSC.BookedRate , MSC.StorageCharges, MSC.BookedQty, MSC.ContractName,MSC.ContractDate,MSC.ClientIDRow,MSC.ExtendedPeriod "
tmp = tmp & " From Mst_StorageContract MSC "
tmp = tmp & " Inner Join Mst_Client MC On MSC.ClientIDRow = MC.ClientIDRow "
tmp = tmp & " Inner Join Mst_Location ML On MSC.LocationID = ML.LocationID "
tmp = tmp & " Left Join Mst_Commodity MComm On MSC.CommodityID = MComm.CommodityID "
tmp = tmp & " Where MSC.ContractID = " & mContractID_
Call TmpTable

If TmpRs.RecordCount > 0 Then
   TxtLocationID = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
   TxtClientID = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
   TxtCommodityID = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
   TxtBookedAsPer = IIf(IsNull(TmpRs!BookedAsPer), "", TmpRs!BookedAsPer)
   TxtBookedRate = IIf(IsNull(TmpRs!BookedRate), 0, TmpRs!BookedRate)
   TxtBookedQty = IIf(IsNull(TmpRs!BookedQty), "", TmpRs!BookedQty)
   TxtStorageCharges = IIf(IsNull(TmpRs!StorageCharges), "", TmpRs!StorageCharges)
   TxtBookedFrom = IIf(IsNull(TmpRs!BookedFrom), "", TmpRs!BookedFrom)
   TxtBookedTo = IIf(IsNull(TmpRs!BookedTo), "", TmpRs!BookedTo)
   TxtExtendedUpto = IIf(IsNull(TmpRs!ExtendedPeriod), "", TmpRs!ExtendedPeriod)
   mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
   mStorageReservationDate = IIf(IsNull(TmpRs!ContractDate), Date, TmpRs!ContractDate)
   mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
End If


End Sub
Private Sub FillGrid3()
Dim wc, mGodownMapped, mCMPMapped As Variant

Call Grid_Head2
'
'mCMPMapped = ""
'mGodownMapped = ""
'
'For i = 2 To MSHFlexGrid2.Rows - 1
'    If MSHFlexGrid2.TextMatrix(i, 0) = "" Then Exit For
'    If mCMPMapped = "" Then
'       mCMPMapped = "'" & MSHFlexGrid2.TextMatrix(i, 0) & "'"
'    Else
'       mCMPMapped = mCMPMapped & ", " & "'" & MSHFlexGrid2.TextMatrix(i, 0) & "'"
'    End If
'
'    If mGodownMapped = "" Then
'       mGodownMapped = "'" & MSHFlexGrid2.TextMatrix(i, 1) & "'"
'    Else
'       mGodownMapped = mGodownMapped & ", " & "'" & MSHFlexGrid2.TextMatrix(i, 1) & "'"
'    End If
'Next

wc = " Where MGSL.Flag = 'E' And MCMP.LocationID = " & mLocationID
wc = wc & " And MGSL.ExchangeTypeID = " & mExchnageTypeID
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   wc = wc & " And MGSL.ClientIDRow = " & mClientIDRow
End If

'wc = wc & " And MCMP.CMPName in (" & mCMPMapped & ") And MG.GodownNo in (" & mGodownMapped & ") "
'wc = wc & " And (MGSL.nContractID = " & mContractID & " or MGSL.nContractID is Null or MGSL.nContractID = 0) "

If TxtCommodityID.Text <> "" Then
   mCommodityID = GetCommodityID(Trim(TxtCommodityID.Text))
   wc = wc & " And MGSL.CommodityID=" & mCommodityID
End If


tmp = " SELECT MCMP.CMPName, MG.GodownNo, MGSL.StackNo, MGSL.LotNo, MCom.Commodity, MCl.ClientName, MGSL.BalanceBags, "
tmp = tmp & " MGSL.BalanceWeight, MGSL.GSLID, MGSL.SM_Prefix, MGSL.NContractID, MGSL.CDTFDate, MGSL.CDTFNo"
tmp = tmp & " FROM Mst_GSL MGSL"
tmp = tmp & " Inner JOIN Mst_Godown MG ON MG.GodownID = MGSL.GodownID"
tmp = tmp & " Left Join Mst_Commodity MCom ON MCom.CommodityID = MGSL.CommodityID"
tmp = tmp & " Inner Join Mst_CMP MCMP ON MCMP.CMPID = MGSL.CMPID"
tmp = tmp & " Inner Join Mst_Client MCL ON MGSL.ClientIDRow = MCL.ClientIDRow"

tmp1 = " Order By MCL.ClientName, MCom.Commodity, MCMP.CMPName, MG.GodownNo, MGSL.StackNo, MGSL.LotNo "

tmp = tmp & wc & tmp1

Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       For j = 1 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.TextMatrix(x, j) = IIf(IsNull(TmpRs.Fields(j - 1)), "", TmpRs.Fields(j - 1))
           If IIf(IsNull(TmpRs!NContractID), 0, TmpRs!NContractID) > 0 And IIf(IsNull(TmpRs!NContractID), 0, TmpRs!NContractID) <> mContractID Then
              MSHFlexGrid3.row = x
              MSHFlexGrid3.Col = j
              MSHFlexGrid3.CellBackColor = vbRed
           End If
       Next
       If Val(MSHFlexGrid3.TextMatrix(x, 11)) <> 0 Then
          MSHFlexGrid3.row = x
          MSHFlexGrid3.Col = 1
          If MSHFlexGrid3.CellBackColor <> vbRed Then
             MSHFlexGrid3.TextMatrix(x, 0) = "*"
          End If
          MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(x, 7))
          MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(x, 8))
       End If
       
       
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       TmpRs.MoveNext
       x = x + 1
   Next
End If
End Sub

Private Sub ClearContractDetails()

TxtLocationID = ""
TxtClientID = ""
TxtCommodityID = ""
TxtBookedAsPer = ""
TxtBookedRate = ""
TxtBookedQty = ""
TxtStorageCharges = ""
TxtBookedFrom = ""
TxtBookedTo = ""
TxtExtendedUpto = ""
mLocationID = Null
mClientIDRow = 0
End Sub

Private Sub ClearGodownDetails()

'CmbCMPName.Text = ""
' CmbGodownNo.Text = "" 'delete it
TxtUnitTypeID = ""
TxtGodownAddress = ""
TxtGodownCapacity = ""
TxtBagsCapacity = ""
TxtReservationQty = ""
' CmdAddToList.Caption = "Add To List" 'delete it
' CmdRemoveFromList.Enabled = False  'delete it

End Sub
Private Sub CmdCMPName_Click()
FrmMst_CMP.Show
End Sub

Private Sub CmdGo_Click()
Dim wc As String
Dim tmp1 As String
Call Grid_Head1
wc = ""

If SFrom.Value > STo.Value Then
   MsgBox "'From date' must be always less than 'To date'"
   SFrom.SetFocus
   Exit Sub
End If

tmp = " SELECT Distinct Mst_GSL.nContractID ,Txn_SRGodownMapping.SRGodownMappingID,"
tmp = tmp & "  Txn_SRGodownMapping.SRGodownMappingDate,Mst_StorageContract.ContractName,Mst_Client.ClientName,Mst_Location.LocationName,"
tmp = tmp & "  Mst_Commodity.Commodity ,"
tmp = tmp & "  Mst_StorageContract.BookedFrom,Mst_StorageContract.BookedTo,Mst_StorageContract.BookedAsPer,"
tmp = tmp & "  Mst_StorageContract.BookedQty,Mst_StorageContract.BookedRate,"
tmp = tmp & "  Mst_StorageContract.StorageCharges, Sum(Mst_GSL.BalanceWeight) as ReservedStockQty"
tmp = tmp & "  From Mst_GSL"
tmp = tmp & "  Inner Join Txn_SRGodownMapping On Mst_GSL.nContractID = Txn_SRGodownMapping.ContractID"
tmp = tmp & "  Inner Join Mst_StorageContract On Txn_SRGodownMapping.ContractID =  Mst_StorageContract.ContractID"
tmp = tmp & "  Inner Join Mst_Client On Mst_StorageContract.ClientIDRow = Mst_Client.ClientIDRow"
tmp = tmp & "  Inner Join Mst_Location On Mst_StorageContract.LocationID = Mst_Location.LocationID"
tmp = tmp & "  Left Join  Mst_Commodity On Mst_StorageContract.CommodityID = Mst_Commodity.CommodityID"
tmp = tmp & "  Where  Mst_GSL.NContractID is not Null AND Mst_GSL.NContractID <> '0'  And (Mst_GSL.Flag = 'E')"

If Gen_DBType = "MDB" Then
   wc = " And Txn_SRGodownMapping.SRGodownMappingDate Between #" & SFrom.Value & "# And  #" & STo.Value & "#"
Else
   wc = " And Txn_SRGodownMapping.SRGodownMappingDate Between '" & SFrom.Value & "' And  '" & STo.Value & "'"
End If

If CmbSLocation.Text <> "" Then
   wc = GenWc(wc, mSM_Prefix, "Mst_GSL.SM_Prefix", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "Mst_Location.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, CmbSCommodity.Text, "Mst_Commodity.Commodity", "S", "N")
End If

If CmbSClient.Text <> "" Then
   wc = GenWc(wc, CmbSClient.Text, "Mst_Client.ClientName", "S", "N")
End If

If TxtSearchContract <> "" Then
   wc = GenWc(wc, TxtSearchContract, "Mst_StorageContract.ContractName", "S", "L")
End If

'--And Mst_StorageContract.ContractName = 'Muktsar - LD'
'--And LocationName = 'Davangere'

tmp1 = "  Group by  Mst_GSL.nContractID,Txn_SRGodownMapping.SRGodownMappingID,"
tmp1 = tmp1 & "  Txn_SRGodownMapping.SRGodownMappingDate,Mst_StorageContract.ContractName,"
tmp1 = tmp1 & "  Mst_Client.ClientName,Mst_Location.LocationName,"
tmp1 = tmp1 & "  Mst_Commodity.Commodity ,"
tmp1 = tmp1 & "  Mst_StorageContract.BookedFrom,Mst_StorageContract.BookedTo,"
tmp1 = tmp1 & "  Mst_StorageContract.BookedQty,Mst_StorageContract.BookedRate,"
tmp1 = tmp1 & "  Mst_StorageContract.StorageCharges,Mst_StorageContract.BookedAsPer"
tmp1 = tmp1 & "  Order By Mst_GSL.NContractID,Mst_StorageContract.ContractName,Mst_Client.ClientName,Mst_Location.LocationName,Mst_Commodity.Commodity "

tmp = tmp & wc & tmp1

Call TmpTable



Call Grid_Head1
'original code

'For iu = 1 To TmpRs.RecordCount
'    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'    For c = 0 To MSHFlexGrid1.Cols - 1
'        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'    Next
'    MSHFlexGrid1.TextMatrix(iu, 2) = Format(MSHFlexGrid1.TextMatrix(iu, 2), "dd-mmm-yyyy")
'    TmpRs.MoveNext
'Next

 If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!NContractID), "", TmpRs!NContractID)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!SRGodownMappingID), "", TmpRs!SRGodownMappingID)
       'mLocationId = TmpRs1!LocationID

       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!SRGodownMappingDate), "", Format(TmpRs!SRGodownMappingDate, "dd-mm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ContractName), "", TmpRs!ContractName)
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!BookedFrom), "", Format(TmpRs!BookedFrom, "dd-mm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!BookedTo), "", Format(TmpRs!BookedTo, "dd-mm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!BookedAsPer), "", TmpRs!BookedAsPer)
       MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs!BookedQty), "", TmpRs!BookedQty)
       MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!BookedRate), "", TmpRs!BookedRate)
       MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!StorageCharges), "", TmpRs!StorageCharges)
       MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!ReservedStockQty), "", TmpRs!ReservedStockQty)
       
       TmpRs.MoveNext
   Next

End If

Label22.Refresh
Label22.Caption = TmpRs.RecordCount


End Sub


Private Sub CmdGSLExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_SContractExpireGSLDetail.xls")
Call Xls_Detail_Rpt1
End Sub

Public Sub Xls_Detail_Rpt1()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_SContractExpireGSLDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid3.Cols - 1
         
        If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
        End If
        MSHFlexGrid3.row = I
        MSHFlexGrid3.Col = C
        If MSHFlexGrid3.CellBackColor = vbRed Then
            oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdStorageReservation_Click()
FrmMst_StorageContract.Show
End Sub

Private Sub DeleteCmd_Click()
'Dim Ans As Variant
'On Error GoTo msgError
'Ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
'If Ans = vbYes Then
'   tmp = "Delete from Txn_SRGodownMapDetails Where SRGodownMappingID = " & TxtMappingID
'   Call TmpTable
'
''   tmp = "SELECT * from Mst_EmpCMPMapping where EmployeeID = " & TxtEmpID
''   Call TmpTable
'
'   RsExpiredGSL_Contract.Requery
'   RsExpiredGSL_Contract.Find "SRGodownMappingID = " & TxtMappingID
'   If RsExpiredGSL_Contract.EOF Then Exit Sub
'   RsExpiredGSL_Contract.Delete
'   RsExpiredGSL_Contract.Update
'   If RsExpiredGSL_Contract.RecordCount = 0 Then
'      Call AddCmd_Click
'      Exit Sub
'   End If
'   RsExpiredGSL_Contract.MoveNext
'   If RsExpiredGSL_Contract.EOF() Then
'      RsExpiredGSL_Contract.MoveLast
'   End If
'   Call InData
'   'TmpRs.Close
'End If
'
'Exit Sub
'msgError:
'MsgBox "Child Record Present "
'RsExpiredGSL_Contract.CancelUpdate
End Sub

Private Sub EditCmd_Click()

Edit_Flg = "E"
Frame3.Enabled = True
Check1.Enabled = True
Call FillGrid3
Call GButtonLock(Me, False)



End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If


SFrom.Value = Date - 365
STo.Value = Date
CmbSLocation.Text = ""
CmbSCommodity.Text = ""
CmbSClient.Text = ""
TxtSearchDescription = ""
Label22.Caption = ""
Call Grid_Head1

Call Grid_Head
Call Grid_Head2
mExchnageTypeID = 1
'Call Grid_Head
Call TableExpiredGSL_Contract(" Where Mst_GSL.nContractID is not Null AND Mst_GSL.nContractID <> '0' ", "Order By Mst_GSL.nContractID")
If RsExpiredGSL_Contract.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
mSrearchfirst = True
Call GButtonLock(Me, True)
Call Indata
If Gen_ShowGSLMapping = True Then
   Call SearchCmd_Click
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsExpiredGSL_Contract.MoveFirst

If RsExpiredGSL_Contract.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsExpiredGSL_Contract.MoveFirst
RsExpiredGSL_Contract.Find "nContractID = " & Val(mCaId) & ""
If RsExpiredGSL_Contract.EOF Then
   RsExpiredGSL_Contract.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub MSHFlexGrid2_Click()
'If SaveCmd.Enabled = False Then Exit Sub
'
'Dim i As Variant
'
'i = MSHFlexGrid2.RowSel
'If MSHFlexGrid2.TextMatrix(i, 0) = "" Then Exit Sub
'CmbCMPName.Text = MSHFlexGrid2.TextMatrix(i, 0)
'mCMPID = GetCMPID(MSHFlexGrid2.TextMatrix(i, 0))
'CmbGodownNo.Text = MSHFlexGrid2.TextMatrix(i, 1)
'TxtUnitTypeID = MSHFlexGrid2.TextMatrix(i, 2)
'TxtBagsCapacity = Val(MSHFlexGrid2.TextMatrix(i, 3))
'TxtGodownCapacity = Val(MSHFlexGrid2.TextMatrix(i, 4))
'TxtReservationQty = Val(MSHFlexGrid2.TextMatrix(i, 5))
'mGodownID = MSHFlexGrid2.TextMatrix(i, 6)
''TxtDetailsID = MSHFlexGrid2.TextMatrix(i, 7)
'TxtGodownAddress = MSHFlexGrid2.TextMatrix(i, 7)
'
'
'CmdAddToList.Caption = "Update List"
'CmdRemoveFromList.Enabled = True
'CmbCMPName.SetFocus
 
End Sub

Private Function Validate_DuplicateGrid() As Boolean
Dim Retval As Boolean
Retval = True
If MSHFlexGrid2.Rows = 2 Then
   Retval = True
Else
    For I = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
           If CmdAddtoList.Caption = "Update List" Then
              If I <> MSHFlexGrid2.RowSel Then
                 If Val(MSHFlexGrid2.TextMatrix(I, 6)) = mGodownID Then
                    Retval = False
                 End If
              End If
           Else
              If Val(MSHFlexGrid2.TextMatrix(I, 6)) = mGodownID Then
                 Retval = False
              End If
           End If
        End If
    Next
End If
Validate_DuplicateGrid = Retval
End Function



Private Sub MSHFlexGrid3_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim mRow As Variant
If MSHFlexGrid3.Rows < 4 Then Exit Sub
'If MSHFlexGrid3.Col = 1 Then
   mRow = MSHFlexGrid3.row
   If MSHFlexGrid3.TextMatrix(mRow, 0) = "*" Then
      MSHFlexGrid3.TextMatrix(mRow, 11) = ""
      MSHFlexGrid3.TextMatrix(mRow, 0) = ""
      MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) - Val(MSHFlexGrid3.TextMatrix(mRow, 7))
      MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) - Val(MSHFlexGrid3.TextMatrix(mRow, 8))
   Else
      MSHFlexGrid3.TextMatrix(mRow, 11) = mContractID
      MSHFlexGrid3.TextMatrix(mRow, 0) = "*"
      MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(mRow, 7))
      MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(mRow, 8))
   End If
'End If

End Sub
Private Function ChkGSLSelection() As Boolean
Dim Retval As Boolean
Retval = False

For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 1) <> "" Then
       If MSHFlexGrid3.TextMatrix(I, 0) = "*" Then
          Retval = True
          Exit For
       End If
    End If
Next

ChkGSLSelection = Retval
End Function

Private Sub SaveCmd_Click()

If CmbStorageDescription.Text = "" Then
   MsgBox "Blank Storage Description not allowed!!!"
   CmbStorageDescription.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 4 Then
   MsgBox "No godown map to this storage reservation !!!"
   CmbStorageDescription.SetFocus
   Exit Sub
End If

If MSHFlexGrid3.Rows < 3 Then
   MsgBox "No GSL Details in GRID !!!"
   CmbStorageDescription.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   If ChkGSLSelection = False Then
      MsgBox "Please select atleast one GSL from Grid!!!"
      Exit Sub
   End If
End If

Call TableMst_GSL(" Where sm_prefix = '" & TxtSMPrefix & "'")

For x = 1 To MSHFlexGrid3.Rows - 1
    MSHFlexGrid3.row = x
    MSHFlexGrid3.Col = 1
    If MSHFlexGrid3.CellBackColor <> vbRed Then
      RsMst_GSL.MoveFirst
      RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid3.TextMatrix(x, 9))
      If Not RsMst_GSL.EOF Then
         RsMst_GSL!NContractID = 0
         RsMst_GSL!ContractID = 0
         RsMst_GSL.Update
      End If
   End If
Next

''''''''''''''For Inserting Log Data in Log_ContractGSL
tmp = "Select * From Log_ContractGSL"
Call TmpTable
''''''''''''''For Inserting Log Data in Log_ContractGSL

For x = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(x, 0) = "*" Then
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "GSLID = " & Val(MSHFlexGrid3.TextMatrix(x, 9))
       If Not RsMst_GSL.EOF Then
          RsMst_GSL!ContractID = mContractID
          RsMst_GSL!NContractID = mContractID
          RsMst_GSL.Update
          ''''''''''''''For Inserting Log Data in Log_ContractGSL
          TmpRs.AddNew
          TmpRs!GSLID = Val(MSHFlexGrid3.TextMatrix(x, 9))
          TmpRs!SM_Prefix = TxtSMPrefix.Text
          TmpRs!ContractID = mContractID
          TmpRs!UpdatedOn = Now
          TmpRs!UpdatedBy = Gen_UserLoginID
          TmpRs!EntryFlag = "E"
          TmpRs.Update
          ''''''''''''''For Inserting Log Data in Log_ContractGSL
       End If
    End If
Next

Call TableExpiredGSL_Contract(" Where Mst_GSL.NContractID is not Null AND Mst_GSL.NContractID <> '0' ", "Order By Mst_GSL.NContractID")

RsExpiredGSL_Contract.Find "NContractID = " & mContractID & ""
If RsExpiredGSL_Contract.RecordCount > 0 Then
   If RsExpiredGSL_Contract.EOF Then
      RsExpiredGSL_Contract.MoveFirst
   End If
   Call GButtonLock(Me, True)
   Call Indata
End If

'Call InData

End Sub


Private Sub Indata()

CmbStorageDescription.Visible = True
CmbStorageDescription.Enabled = False
CmdStorageReservation.Enabled = False

mContractID = RsExpiredGSL_Contract!NContractID
TxtContractID.Text = mContractID
CmbStorageDescription = GetContract(mContractID)
'Call GetContractDetails(mContractID)

Check1.Enabled = False
Call CmbStorageDescription_LostFocus

Check1.Value = 0
Frame3.Enabled = False
'Call Grid_Head
'Call Grid_Head2
'Call Indata_MappingDetails
'Call FillGrid3

Call GButtonLock(Me, True)

'SearchCmd.Enabled = False
DeleteCmd.Enabled = False


End Sub

Private Sub Indata_MappingDetails()

'CmbCMPName.Text = "" 'delete it
Call ClearGodownDetails

tmp = " Select MC.CMPName,MG.GodownNo,MU.UnitType, MG.Address,MG.BagsCapacity,MG.GodownCapacity, TSRGD.ReservationQty , "
tmp = tmp & " MG.GodownID From Txn_SRGodownMapDetails TSRGD "
tmp = tmp & " Inner Join Mst_Godown MG On TSRGD.GodownID = MG.GodownID "
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID "
tmp = tmp & " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID "
tmp = tmp & " Where TSRGD.SRGodownMappingID = " & mMappingID
Call TmpTable

If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(x, 4))
       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(x, 5))
       MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(x, 6))
       TmpRs.MoveNext
       x = x + 1
   Next
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsExpiredGSL_Contract.MoveNext
If RsExpiredGSL_Contract.EOF Then
   RsExpiredGSL_Contract.MoveLast
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
RsExpiredGSL_Contract.MovePrevious
If RsExpiredGSL_Contract.BOF Then
   RsExpiredGSL_Contract.MoveFirst
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
RsExpiredGSL_Contract.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsExpiredGSL_Contract.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsExpiredGSL_Contract.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame1.Height - 200 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width '- 200
Frame0.Visible = True
'MSHFlexGrid1.Clear
MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1200

''===========comment on 7 july 2012 for not clearing search grid=============
'If mSrearchfirst = True Then
'   SFrom.Value = Date - 365
'   STo.Value = Date
'   CmbSLocation.Text = ""
'   CmbSCommodity.Text = ""
'   CmbSClient.Text = ""
'   TxtSearchDescription = ""
'   Label22.Caption = ""
'   Call Grid_Head1
'End If
'
'mSrearchfirst = False
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Public Sub Grid_Head1()
With MSHFlexGrid1
   .Clear
   .Cols = 14
   .Rows = 2
   .FixedRows = 1
   
   .TextMatrix(0, 0) = "Contract ID"
   .TextMatrix(0, 1) = "Mapping ID"
   .TextMatrix(0, 2) = "Mapping Date"
   .TextMatrix(0, 3) = "Storage Description"
   .TextMatrix(0, 4) = "Client"
   .TextMatrix(0, 5) = "Location"
   .TextMatrix(0, 6) = "Commodity"
   
   .TextMatrix(0, 7) = "Booked From"
   .TextMatrix(0, 8) = "Booked To"
   .TextMatrix(0, 9) = "Booked As Per"
   .TextMatrix(0, 10) = "Qty"
   .TextMatrix(0, 11) = "Rate"
   .TextMatrix(0, 12) = "Storage Amount"
   
   .TextMatrix(0, 13) = "Reserved Stock Qty"
   
   
   
   .ColWidth(0) = 1000
   .ColWidth(1) = 1000
   .ColWidth(2) = 1200
   .ColWidth(3) = 3600
   .ColWidth(4) = 2000
   .ColWidth(5) = 2200
   .ColWidth(6) = 2200
   
   .ColWidth(7) = 1100
   .ColWidth(8) = 1000
   .ColWidth(9) = 1400
   .ColWidth(10) = 900
   .ColWidth(11) = 1200
   .ColWidth(12) = 1500
   .ColWidth(13) = 1200
   
   .RowHeight(0) = 600
End With
End Sub

Public Sub Grid_Head2()
With MSHFlexGrid3
   .Clear
   .Cols = 14 '12
   .Rows = 3
   .FixedRows = 2
   
   .TextMatrix(1, 0) = "Select"
   .TextMatrix(1, 1) = "CMP"
   .TextMatrix(1, 2) = "Godown"
   .TextMatrix(1, 3) = "Stack"
   .TextMatrix(1, 4) = "Lot"
   .TextMatrix(1, 5) = "Commodity"
   .TextMatrix(1, 6) = "Client"
   .TextMatrix(1, 7) = "No of Bags"
   .TextMatrix(1, 8) = "Quantity"
   .TextMatrix(1, 9) = "GSLID"
   .TextMatrix(1, 10) = "SM_Prefix"
   .TextMatrix(1, 11) = "Mapped"
   .TextMatrix(1, 12) = "CDTF Date"
   .TextMatrix(1, 13) = "CDTF No"
   
   .ColWidth(0) = 1000
   .ColWidth(1) = 2500
   .ColWidth(2) = 1000
   .ColWidth(3) = 1000
   .ColWidth(4) = 1000
   .ColWidth(5) = 1500
   .ColWidth(6) = 1500
   .ColWidth(7) = 1000
   .ColWidth(8) = 1000
   .ColWidth(9) = 1000
   .ColWidth(10) = 1000
   .ColWidth(11) = 1000
   
   .RowHeight(1) = 400
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_SContractExpireGSLMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_SContractExpireGSLMapping.xls")
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

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
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
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbSClient_GotFocus()

tmp = CmbSClient.Text
Call TableMst_Client("Where ExchangeTypeID = " & GetExchangeID("Non NCDEX"))
CmbSClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
   CmbSClient.AddItem RsMst_Client!ClientName
   RsMst_Client.MoveNext
Next
CmbSClient.Text = tmp

End Sub

Private Sub CmbSClient_LostFocus()
If CmbSClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClient.SetFocus
   Exit Sub
End If

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

Private Sub SetScrollBars()
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame1.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame1.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame1.Height >= Me.Height Then VScroll1.Enabled = True
If Frame1.Height < Me.Height Then VScroll1.Enabled = False
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True
If Frame1.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame1.Height
   l = (a * x) / 1000
   Frame1.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = VScroll1.Value
   a = Me.Height - Frame1.Height
   l = (a * x) / 1000
   Frame1.Top = -l
'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = HScroll1.Value
   a = Me.Width - Frame1.Width
   l = (a * x) / 1000
   Frame1.Left = -l
End Sub

Private Sub HScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = HScroll1.Value
   a = Me.Width - Frame1.Width
   l = (a * x) / 1000
   Frame1.Left = -l
End Sub
