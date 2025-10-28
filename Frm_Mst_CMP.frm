VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_CMP 
   Caption         =   "CMP Master"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   101
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   100
      Top             =   10080
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15375
      Begin VB.Frame Frame0 
         Height          =   1125
         Left            =   240
         TabIndex        =   82
         Top             =   6720
         Visible         =   0   'False
         Width           =   14040
         Begin VB.CheckBox ChkNoOpenGodown 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6360
            TabIndex        =   116
            Top             =   675
            Width           =   195
         End
         Begin VB.CheckBox ChkNoGodown 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2160
            TabIndex        =   112
            Top             =   675
            Width           =   195
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   10440
            TabIndex        =   87
            Top             =   195
            Width           =   1000
         End
         Begin VB.TextBox TxtSearchCMP 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   83
            Top             =   195
            Width           =   3600
         End
         Begin VB.ComboBox CmbLocationSearch 
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
            Left            =   6810
            Sorted          =   -1  'True
            TabIndex        =   84
            Top             =   195
            Width           =   2400
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   9360
            TabIndex        =   85
            Top             =   195
            Width           =   1000
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "All Godown Closed but CMP Open"
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
            Left            =   3120
            TabIndex        =   117
            Top             =   630
            Width           =   3045
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "No Godown Created"
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
            TabIndex        =   113
            Top             =   630
            Width           =   1830
         End
         Begin VB.Label Label1 
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
            Left            =   6030
            TabIndex        =   91
            Top             =   255
            Width           =   780
         End
         Begin VB.Label Label17 
            Caption         =   "Search by CMP Name"
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
            TabIndex        =   90
            Top             =   240
            Width           =   2220
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
            Left            =   11535
            TabIndex        =   88
            Top             =   255
            Width           =   2415
         End
         Begin VB.Label Label22 
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
            Left            =   13920
            TabIndex        =   86
            Top             =   240
            Width           =   75
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2295
         Left            =   285
         TabIndex        =   81
         Top             =   6105
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   4048
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1200
         Left            =   120
         TabIndex        =   92
         Top             =   8520
         Width           =   15045
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   225
            TabIndex        =   99
            Top             =   240
            Width           =   2880
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   225
            TabIndex        =   98
            Top             =   645
            Width           =   2880
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3120
            TabIndex        =   33
            Top             =   240
            Width           =   2880
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3105
            TabIndex        =   97
            Top             =   645
            Width           =   2880
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   6000
            TabIndex        =   96
            Top             =   240
            Width           =   2880
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5985
            TabIndex        =   95
            Top             =   645
            Width           =   2880
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   8865
            TabIndex        =   94
            Top             =   240
            Width           =   2880
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   8865
            TabIndex        =   93
            Top             =   645
            Width           =   2880
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   11745
            TabIndex        =   35
            Top             =   240
            Width           =   2880
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   11745
            TabIndex        =   34
            Top             =   645
            Width           =   2880
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8415
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   15030
         Begin VB.TextBox TxtBlockDays 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   12270
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   114
            Top             =   1932
            Width           =   2640
         End
         Begin VB.TextBox TxtFaxNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   12270
            MaxLength       =   50
            TabIndex        =   10
            Top             =   990
            Width           =   2640
         End
         Begin VB.TextBox TxtMobileNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   12270
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1455
            Width           =   2640
         End
         Begin VB.TextBox TxtTelephoneNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   12270
            MaxLength       =   50
            TabIndex        =   9
            Top             =   585
            Width           =   2640
         End
         Begin VB.TextBox TxtAddress 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   1965
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   1404
            Width           =   8760
         End
         Begin VB.TextBox TxtCMPID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   1965
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   180
            Width           =   1320
         End
         Begin VB.Frame Frame2 
            Caption         =   "Godown Details"
            Height          =   3150
            Left            =   120
            TabIndex        =   39
            Top             =   2820
            Width           =   14850
            Begin VB.CheckBox ChkYearlyBilling 
               Height          =   195
               Left            =   13890
               TabIndex        =   110
               Top             =   2055
               Width           =   195
            End
            Begin VB.TextBox TxtCOEN 
               Appearance      =   0  'Flat
               Height          =   510
               Left            =   8100
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   1755
               Width           =   2955
            End
            Begin MSComCtl2.DTPicker TxtCENDate 
               Height          =   360
               Left            =   5400
               TabIndex        =   22
               Top             =   1830
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   635
               _Version        =   393216
               Format          =   6488065
               CurrentDate     =   39867
            End
            Begin VB.TextBox TxtCEN 
               Appearance      =   0  'Flat
               Height          =   510
               Left            =   1500
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Top             =   1755
               Width           =   3075
            End
            Begin VB.ComboBox CmbCategoryID 
               Height          =   315
               Left            =   1485
               Sorted          =   -1  'True
               TabIndex        =   13
               Top             =   600
               Width           =   2685
            End
            Begin VB.CommandButton CmdCategoryID 
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
               Left            =   4320
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   600
               Width           =   375
            End
            Begin VB.CommandButton CmdUnitTypeID 
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
               Left            =   4320
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   210
               Width           =   375
            End
            Begin VB.TextBox TxtBagsCapacity 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   12195
               MaxLength       =   50
               TabIndex        =   19
               Top             =   615
               Width           =   2475
            End
            Begin VB.TextBox TxtGodownCapacity 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   12195
               TabIndex        =   20
               Top             =   1005
               Width           =   2460
            End
            Begin VB.TextBox TxtGodownArea 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   12195
               MaxLength       =   100
               TabIndex        =   18
               Top             =   210
               Width           =   2460
            End
            Begin VB.TextBox TxtGodownTelephoneNo 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   6180
               MaxLength       =   50
               TabIndex        =   17
               Top             =   990
               Width           =   4050
            End
            Begin VB.ComboBox CmbUnitTypeID 
               Height          =   315
               Left            =   1485
               Sorted          =   -1  'True
               TabIndex        =   12
               Top             =   210
               Width           =   2685
            End
            Begin VB.TextBox TxtGodownAddress 
               Appearance      =   0  'Flat
               Height          =   750
               Left            =   6180
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Top             =   210
               Width           =   4035
            End
            Begin VB.CommandButton CmdLicenseNo 
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
               Left            =   4320
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1440
               Width           =   375
            End
            Begin VB.TextBox TxtLicenseDesc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   6180
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   1365
               Width           =   8460
            End
            Begin VB.TextBox TxtUnitTypeID 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1500
               TabIndex        =   43
               Top             =   210
               Width           =   2640
            End
            Begin VB.TextBox TxtCategoryID 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1500
               TabIndex        =   102
               Top             =   600
               Width           =   2640
            End
            Begin VB.TextBox TxtGodownNo 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1500
               MaxLength       =   25
               TabIndex        =   14
               Top             =   960
               Width           =   2640
            End
            Begin VB.ComboBox CmbLicenseNo 
               Height          =   315
               Left            =   1500
               Sorted          =   -1  'True
               TabIndex        =   15
               Top             =   1365
               Width           =   2640
            End
            Begin VB.CommandButton CmdAdd2List 
               Caption         =   "Add to List"
               Height          =   375
               Left            =   13545
               TabIndex        =   32
               Top             =   2640
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker TxtCOENDate 
               Height          =   360
               Left            =   12000
               TabIndex        =   24
               Top             =   1830
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   635
               _Version        =   393216
               Format          =   6488065
               CurrentDate     =   39867
            End
            Begin MSComCtl2.DTPicker TxtLeasedUpto 
               Height          =   360
               Left            =   6240
               TabIndex        =   104
               Top             =   2610
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   635
               _Version        =   393216
               Format          =   6488065
               CurrentDate     =   39867
            End
            Begin VB.CheckBox Check1 
               Height          =   195
               Left            =   3930
               TabIndex        =   27
               Top             =   2700
               Width           =   195
            End
            Begin VB.TextBox TxtAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   7800
               MaxLength       =   50
               TabIndex        =   29
               Top             =   2610
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.TextBox TxtAdvAmount 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   9690
               MaxLength       =   50
               TabIndex        =   30
               Top             =   2610
               Visible         =   0   'False
               Width           =   1605
            End
            Begin MSComCtl2.DTPicker GodownStartDate 
               Height          =   360
               Left            =   270
               TabIndex        =   25
               Top             =   2610
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   635
               _Version        =   393216
               Format          =   6488065
               CurrentDate     =   39867
            End
            Begin MSComCtl2.DTPicker GodownCloseDate 
               Height          =   360
               Left            =   4605
               TabIndex        =   28
               Top             =   2610
               Visible         =   0   'False
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   635
               _Version        =   393216
               Format          =   6488065
               CurrentDate     =   39867
            End
            Begin MSComCtl2.DTPicker TxtAgreementEndDate 
               Height          =   360
               Left            =   2010
               TabIndex        =   26
               Top             =   2610
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   635
               _Version        =   393216
               Format          =   6488065
               CurrentDate     =   39867
            End
            Begin VB.ComboBox CmbNCDEXAcccreditedQty 
               Height          =   315
               Left            =   11400
               Sorted          =   -1  'True
               TabIndex        =   103
               Top             =   2640
               Width           =   1875
            End
            Begin VB.TextBox TxtAccreditedQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   13320
               MaxLength       =   50
               TabIndex        =   31
               Top             =   2640
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.TextBox TxtStock 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   14160
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   40
               Top             =   2280
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Yearly Billing"
               Height          =   195
               Left            =   13560
               TabIndex        =   111
               Top             =   1800
               Width           =   885
            End
            Begin VB.Label Label21 
               Caption         =   "Commodity Endorement no"
               Height          =   435
               Left            =   6840
               TabIndex        =   109
               Top             =   1800
               Width           =   1290
            End
            Begin VB.Label Label20 
               Caption         =   "COPE Date"
               Height          =   435
               Left            =   11280
               TabIndex        =   108
               Top             =   1800
               Width           =   585
            End
            Begin VB.Label Label19 
               Caption         =   "Cotton Endorsement no "
               Height          =   435
               Left            =   120
               TabIndex        =   107
               Top             =   1800
               Width           =   1290
            End
            Begin VB.Label Label 
               Caption         =   "CPE Date"
               Height          =   435
               Left            =   4800
               TabIndex        =   106
               Top             =   1800
               Width           =   585
            End
            Begin VB.Label LblCategoryID 
               Caption         =   "Category"
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
               Left            =   120
               TabIndex        =   62
               Top             =   645
               Width           =   1320
            End
            Begin VB.Label LblUnitTypeID 
               Caption         =   "Unit Type"
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
               Left            =   150
               TabIndex        =   61
               Top             =   240
               Width           =   840
            End
            Begin VB.Label LblGodownNo 
               Caption         =   "Godown No"
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
               Left            =   135
               TabIndex        =   60
               Top             =   990
               Width           =   1200
            End
            Begin VB.Label LblBagsCapacity 
               Caption         =   "Capacity in Bags"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   10560
               TabIndex        =   59
               Top             =   705
               Width           =   1560
            End
            Begin VB.Label LblGodownCapacity 
               Caption         =   "Capacity in MT"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   10560
               TabIndex        =   58
               Top             =   1110
               Width           =   1440
            End
            Begin VB.Label LblArea 
               Caption         =   "Area In Sq.Ft."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   10560
               TabIndex        =   57
               Top             =   300
               Width           =   1440
            End
            Begin VB.Label Label3 
               Caption         =   "Telephone No"
               Height          =   300
               Left            =   4890
               TabIndex        =   56
               Top             =   990
               Width           =   1320
            End
            Begin VB.Label Label2 
               Caption         =   "Godown Address"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4890
               TabIndex        =   53
               Top             =   435
               Width           =   1200
            End
            Begin VB.Label Label6 
               Caption         =   "License No"
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
               Left            =   120
               TabIndex        =   52
               Top             =   1410
               Width           =   1200
            End
            Begin VB.Label Label7 
               Caption         =   "Description"
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
               Left            =   4890
               TabIndex        =   51
               Top             =   1440
               Width           =   1200
            End
            Begin VB.Label lblleaseduptodate 
               AutoSize        =   -1  'True
               Caption         =   "Leased Upto Date"
               Height          =   195
               Left            =   6270
               TabIndex        =   105
               Top             =   2325
               Width           =   1305
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Godown Start Date"
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
               TabIndex        =   55
               Top             =   2325
               Width           =   1635
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Close Date"
               Height          =   195
               Left            =   4815
               TabIndex        =   54
               Top             =   2325
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Agreement end date"
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
               Left            =   1815
               TabIndex        =   50
               Top             =   2325
               Width           =   1725
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Close Godown "
               Height          =   195
               Left            =   3600
               TabIndex        =   49
               Top             =   2325
               Width           =   1080
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Amount"
               Height          =   195
               Left            =   8445
               TabIndex        =   48
               Top             =   2325
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Advance Amount"
               Height          =   195
               Left            =   9840
               TabIndex        =   47
               Top             =   2325
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "NCDEX Accredited Qty"
               Height          =   195
               Left            =   11460
               TabIndex        =   46
               Top             =   2325
               Width           =   1650
            End
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Show Close Date"
            Height          =   195
            Left            =   9015
            TabIndex        =   3
            Top             =   270
            Width           =   1695
         End
         Begin VB.TextBox TxtEHoPCMPID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   4695
            MaxLength       =   15
            TabIndex        =   1
            Top             =   180
            Width           =   1440
         End
         Begin VB.TextBox TxtCMPDispalyName 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1965
            MaxLength       =   100
            TabIndex        =   6
            Top             =   996
            Width           =   8760
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
            Left            =   8895
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   2340
            Width           =   5985
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
            Left            =   1965
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   2340
            Width           =   5625
         End
         Begin MSComCtl2.DTPicker StartDate 
            Height          =   360
            Left            =   7575
            TabIndex        =   2
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            Format          =   6488065
            CurrentDate     =   39864
         End
         Begin MSComCtl2.DTPicker CloseDate 
            Height          =   360
            Left            =   12270
            TabIndex        =   4
            Top             =   180
            Visible         =   0   'False
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   635
            _Version        =   393216
            Format          =   6488065
            CurrentDate     =   39864
         End
         Begin VB.TextBox TxtFlag 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   66
            Top             =   3390
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox TxtCMPName 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1965
            MaxLength       =   100
            TabIndex        =   5
            Top             =   588
            Width           =   8760
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
            Left            =   10350
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1965
            Width           =   375
         End
         Begin VB.ComboBox CmbLocationID 
            Height          =   315
            Left            =   1965
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   1932
            Width           =   8400
         End
         Begin VB.TextBox TxtLocationID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1965
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1932
            Width           =   8760
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Block No Of Days"
            Height          =   195
            Left            =   10920
            TabIndex        =   115
            Top             =   2015
            Width           =   1275
         End
         Begin VB.Label LblFlag 
            Caption         =   "Flag"
            Enabled         =   0   'False
            Height          =   240
            Left            =   600
            TabIndex        =   80
            Top             =   3450
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label LblCloseDate 
            AutoSize        =   -1  'True
            Caption         =   "Close Date"
            Height          =   195
            Left            =   10920
            TabIndex        =   79
            Top             =   270
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label LblStartDate 
            AutoSize        =   -1  'True
            Caption         =   "CMP Start Date"
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
            Left            =   6150
            TabIndex        =   78
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label LblFaxNo 
            AutoSize        =   -1  'True
            Caption         =   "Fax No"
            Height          =   195
            Left            =   10920
            TabIndex        =   77
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label LblMobileNo 
            AutoSize        =   -1  'True
            Caption         =   "Mobile No"
            Height          =   195
            Left            =   10920
            TabIndex        =   76
            Top             =   1545
            Width           =   720
         End
         Begin VB.Label LblTelephoneNo 
            AutoSize        =   -1  'True
            Caption         =   "Telephone No"
            Height          =   195
            Left            =   10920
            TabIndex        =   75
            Top             =   675
            Width           =   1020
         End
         Begin VB.Label LblAddress 
            AutoSize        =   -1  'True
            Caption         =   "Address"
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
            Left            =   270
            TabIndex        =   74
            Top             =   1545
            Width           =   690
         End
         Begin VB.Label LblCMPName 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   270
            TabIndex        =   72
            Top             =   675
            Width           =   945
         End
         Begin VB.Label LblCMPID 
            AutoSize        =   -1  'True
            Caption         =   "CMP ID"
            Height          =   195
            Left            =   270
            TabIndex        =   71
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "EHoP CMP ID"
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
            Left            =   3390
            TabIndex        =   70
            Top             =   270
            Width           =   1200
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "CMP Display Name"
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
            Left            =   270
            TabIndex        =   69
            Top             =   1080
            Width           =   1620
         End
         Begin VB.Label Label15 
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
            Left            =   7860
            TabIndex        =   68
            Top             =   2400
            Width           =   885
         End
         Begin VB.Label Label16 
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
            Left            =   270
            TabIndex        =   67
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label LblLocationID 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   270
            TabIndex        =   73
            Top             =   2025
            Width           =   750
         End
      End
   End
End
Attribute VB_Name = "FrmMst_CMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, I As Variant
Dim mAssayerID, mLocationID, mUnitTypeID, mCategoryID, mCMPID, mGodownID, mLicenseID As Double
Dim mAccreditedID, mCapacity  As Variant
Dim mSLocationID As Double
Dim mGridMsg  As Variant
Dim mChkCloseGodown As Boolean
Dim mCxTFDate, mSpillageDate, mAdjTxnDate, mGSLTransferDate, mAgreeExpiryDate, mMsgforDate As Variant
Dim mSM_Prefix As Variant
Private Sub Check1_Click()
Label5.Visible = False
GodownCloseDate.Visible = False
GodownCloseDate.Enabled = False
If Check1.Value = 1 Then
   Label5.Visible = True
   GodownCloseDate.Visible = True
   GodownCloseDate.Enabled = True
   GodownCloseDate.Value = Date
End If
End Sub
Private Sub Check2_Click()
If SaveCmd.Enabled = False Then Exit Sub
 LblCloseDate.Visible = False
 CloseDate.Visible = False
 If Check2.Value = 1 Then
    For I = 2 To MSHFlexGrid2.Rows - 2
        If MSHFlexGrid2.TextMatrix(I, 10) = "" Then
           MsgBox "All godown is not Closed, You can't close this CMP"
           Check2.Value = 0
           Exit Sub
        End If
    Next
    LblCloseDate.Visible = True
    CloseDate.Visible = True
    CloseDate.Enabled = True
    CloseDate.Value = Date
 End If
End Sub
Private Sub CloseDate_LostFocus()
If CloseDate.Value < StartDate.Value Then
   MsgBox "CloseDate Must Be Greater Than Start Date !!!"
   CloseDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CloseDate.Value, "yyyy-mm") Then
   MsgBox "CMP Close Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CloseDate.SetFocus
   Exit Sub
End If

End Sub
'---
Function CheckDuplicateGodown() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 3) = TxtGodownNo Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateGodown = Retval
End Function
'---
Private Sub CmbNCDEXAcccreditedQty_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location !!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp1 = CmbNCDEXAcccreditedQty.Text

tmp = " Select"
tmp = tmp & " AccreditedID , Capacity, AccreditedDescription"
tmp = tmp & " From Mst_NCDEXAccreditedCapacity"
tmp = tmp & " Where AccreditedID in (Select max(AccreditedID) From Mst_NCDEXAccreditedCapacity Where LocationID = " & mLocationID & ")"
Call TmpTable

CmbNCDEXAcccreditedQty.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbNCDEXAcccreditedQty.AddItem TmpRs!AccreditedDescription
    TmpRs.MoveNext
Next
CmbNCDEXAcccreditedQty.Text = tmp1
End Sub

Private Sub CmbNCDEXAcccreditedQty_LostFocus()
If CmbNCDEXAcccreditedQty.Text = "" Then Exit Sub
TmpRs.MoveFirst
TmpRs.Find "AccreditedDescription = '" & CmbNCDEXAcccreditedQty.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid selection "
   CmbLicenseNo.SetFocus
   Exit Sub
End If
mAccreditedID = TmpRs!AccreditedID
mCapacity = TmpRs!Capacity
End Sub

Private Sub CmbLicenseNo_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location !!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
'mLicenseID = Null

tmp = CmbLicenseNo.Text
Call TableMst_License(" Where LocationID = 0 or LicenseID in (Select Distinct LicenseID from Txn_LicenseAllocationDetail Where SLGID = " & mLocationID & " And Flag = 'L' )")
CmbLicenseNo.Clear
If RsMst_License.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_License.RecordCount
    CmbLicenseNo.AddItem RsMst_License!LicenseNo
    RsMst_License.MoveNext
Next
CmbLicenseNo.Text = tmp
End Sub
'----
Private Sub CmbLicenseNo_LostFocus()
If CmbLicenseNo.Text = "" Then
   Exit Sub
End If
RsMst_License.MoveFirst
RsMst_License.Find "LicenseNo = '" & CmbLicenseNo.Text & "'"
If RsMst_License.EOF Then
   MsgBox "Invalid selection "
   CmbLicenseNo.SetFocus
   Exit Sub
End If
mLicenseID = RsMst_License!LicenseID
TxtLicenseDesc = RsMst_License!LicenseDesc
End Sub

Private Sub CmbLocationSearch_GotFocus()
tmp = CmbLocationSearch.Text
Call TableMst_Location(" Where LocationID <> 0 ")
CmbLocationSearch.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationSearch.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationSearch.Text = tmp
End Sub
'----
Private Sub CmbLocationSearch_LostFocus()
If CmbLocationSearch.Text = "" Then
   RsMst_CMP.Filter = ""
   RsMst_CMP.MoveFirst
   mSLocationID = 0
   'Call SearchCmd_Click
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationSearch.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationSearch.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
'RsMst_CMP.Filter = "LocationId = " & RsMst_Location!LocationID
'Call SearchCmd_Click
End Sub

Private Sub CmdAdd2List_Click()
 
If CmbUnitTypeID = "" Then
  MsgBox "Unit Type Can't Be Left Balnk !!!"
  CmbUnitTypeID.SetFocus
  Exit Sub
End If

If CmbCategoryID = "" Then
  MsgBox "Category Can't Be Left Balnk !!!"
  CmbCategoryID.SetFocus
  Exit Sub
End If

If TxtGodownNo = "" Then
  MsgBox "Godown No Can't Be Left Balnk !!!"
  TxtGodownNo.SetFocus
  Exit Sub
End If

If TxtGodownAddress = "" Then
  MsgBox "Godown Address Can't Be Left Balnk !!!"
  TxtGodownAddress.SetFocus
  Exit Sub
End If

'If TxtGodownTelephoneNo = "" Then
'  MsgBox "Godown TelephoneNo Can't Be Left Balnk !!!"
'  TxtGodownTelephoneNo.SetFocus
'  Exit Sub
'End If

If TxtGodownArea = "" Then
  MsgBox "GodownArea Can't Be Left Balnk !!!"
  TxtGodownArea.SetFocus
  Exit Sub
End If

If TxtBagsCapacity = "" Then
  MsgBox "Godown Bags Capacity Can't Be Left Balnk !!!"
  TxtBagsCapacity.SetFocus
  Exit Sub
End If

If TxtGodownCapacity = "" Then
  MsgBox "GodownCapacity Can't Be Left Balnk !!!"
  TxtGodownCapacity.SetFocus
  Exit Sub
End If
If CmbLicenseNo.Text = "" Then
  MsgBox "License No Can't Be Left Balnk !!!"
  CmbLicenseNo.SetFocus
  Exit Sub
End If

If GodownStartDate > TxtAgreementEndDate Then
   MsgBox "Agreement End Date must be greater than Start Date !!!! "
   TxtAgreementEndDate.SetFocus
   Exit Sub
End If


If TxtLeasedUpto.Enabled = True Then
   If LCase(CmbUnitTypeID.Text) = "leased franchisee" Then
      If TxtLeasedUpto.Value < GodownStartDate Or TxtLeasedUpto.Value > TxtAgreementEndDate Then
         MsgBox "Leased Upto Date must be between Godown Start Date and Agreement End Date!!!! "
         TxtLeasedUpto.SetFocus
         Exit Sub
      End If
   End If
End If

If CheckDuplicateGodown = True Then
   MsgBox "Duplicate Godown No not allowd !!!"
   TxtGodownNo.SetFocus
   Exit Sub
End If

If Val(TxtStock) <> 0 And Check1.Value = 1 Then
  MsgBox "Godown Contain Bags Can't Close This Godown"
  Exit Sub
End If



If LCase(CmdAdd2List.Caption) = "update list" And Check1.Value = 1 Then
   tmp = " Select GSLID from Mst_GSL where Flag = 'E' and WHRBags < DepositeBags And GodownID = " & Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0))
   Call Tmp7Table
   
   If TmpRs7.RecordCount > 0 Then
      If MsgBox("Some expired lots are pending for WHR, Still do you want to close this godown?", vbYesNo) = vbNo Then
         GodownCloseDate.SetFocus
         Exit Sub
      End If
   End If
End If

''For NCDEX Accredited Capacity'''''''''''''''''''''''''''''''''''
'If Val(TxtGodownCapacity) < Val(TxtAccreditedQty) Then
'  MsgBox "Godown Capacity can't be less than NCDEX Accredited Qty"
'  TxtAccreditedQty.SetFocus
'  Exit Sub
'End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


If CmdAdd2List.Caption = "Update List" Then
   I = MSHFlexGrid2.RowSel
   MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(I, 6))
   MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(I, 7))
   MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(I, 8))
Else
    I = MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If

'MSHFlexGrid2.TextMatrix(i, 0) = "" 'mGodownID
MSHFlexGrid2.TextMatrix(I, 1) = CmbUnitTypeID.Text
If LCase(MSHFlexGrid2.TextMatrix(I, 1)) = "leased franchisee" Then
    MSHFlexGrid2.TextMatrix(I, 23) = TxtLeasedUpto.Value
Else
    MSHFlexGrid2.TextMatrix(I, 23) = ""
End If

MSHFlexGrid2.TextMatrix(I, 2) = CmbCategoryID.Text
MSHFlexGrid2.TextMatrix(I, 3) = TxtGodownNo.Text
MSHFlexGrid2.TextMatrix(I, 4) = TxtGodownAddress.Text
MSHFlexGrid2.TextMatrix(I, 5) = TxtGodownTelephoneNo.Text
MSHFlexGrid2.TextMatrix(I, 6) = TxtGodownArea.Text
MSHFlexGrid2.TextMatrix(I, 7) = TxtBagsCapacity.Text
MSHFlexGrid2.TextMatrix(I, 8) = TxtGodownCapacity.Text
MSHFlexGrid2.TextMatrix(I, 9) = GodownStartDate.Value

If Check1.Value = 1 Then
   MSHFlexGrid2.TextMatrix(I, 10) = GodownCloseDate.Value
Else
   MSHFlexGrid2.TextMatrix(I, 10) = ""
End If
MSHFlexGrid2.TextMatrix(I, 11) = CmbLicenseNo.Text
MSHFlexGrid2.TextMatrix(I, 12) = mLicenseID

If TxtAgreementEndDate.Visible = True Then
   MSHFlexGrid2.TextMatrix(I, 15) = TxtAgreementEndDate
Else
   MSHFlexGrid2.TextMatrix(I, 15) = ""
End If

MSHFlexGrid2.TextMatrix(I, 16) = TxtAmount
MSHFlexGrid2.TextMatrix(I, 17) = TxtAdvAmount
MSHFlexGrid2.TextMatrix(I, 18) = TxtAccreditedQty

If CmbNCDEXAcccreditedQty.Text <> "" Then
    MSHFlexGrid2.TextMatrix(I, 20) = CmbNCDEXAcccreditedQty
    MSHFlexGrid2.TextMatrix(I, 21) = mCapacity
    MSHFlexGrid2.TextMatrix(I, 22) = mAccreditedID
Else
    MSHFlexGrid2.TextMatrix(I, 20) = ""
    MSHFlexGrid2.TextMatrix(I, 21) = ""
    MSHFlexGrid2.TextMatrix(I, 22) = ""
End If

MSHFlexGrid2.TextMatrix(I, 24) = TxtCEN
MSHFlexGrid2.TextMatrix(I, 25) = TxtCENDate.Value
MSHFlexGrid2.TextMatrix(I, 26) = TxtCOEN
MSHFlexGrid2.TextMatrix(I, 27) = TxtCOENDate.Value
MSHFlexGrid2.TextMatrix(I, 28) = ChkYearlyBilling.Value

MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 6))
MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(I, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(I, 8))
StartDate.Enabled = False
CmbLocationID.Enabled = False
CmbUnitTypeID.Enabled = True
Call ClearControl
End Sub

Sub ClearControl()
mGodownID = ""
CmbUnitTypeID.Text = ""
CmbCategoryID.Text = ""
TxtGodownNo.Text = ""
TxtGodownAddress.Text = ""
TxtGodownTelephoneNo.Text = ""
TxtGodownArea.Text = ""
TxtBagsCapacity.Text = ""
TxtGodownCapacity.Text = ""
GodownStartDate.Value = Date
GodownCloseDate.Value = Date
TxtAgreementEndDate.Value = Date
GodownCloseDate.Value = Date
GodownStartDate.Value = Date
TxtCENDate.Value = Date
TxtCOENDate.Value = Date
TxtCEN = ""
TxtCOEN = ""
CmbLicenseNo = ""
TxtLicenseDesc = ""
TxtAmount = ""
TxtAdvAmount = ""
TxtAccreditedQty = ""
CmbNCDEXAcccreditedQty = ""
TxtLeasedUpto.Visible = False
lblleaseduptodate.Visible = False
mLicenseID = 0
Check1.Value = 0
Check1_Click
ChkYearlyBilling.Value = 0
''added after godown inspection module on 19 apr 2013
If LCase(Gen_UserRole) = "power" Then
   CmdAdd2List.Caption = "Add to List"
   CmdAdd2List.Enabled = True
Else
   CmdAdd2List.Caption = "Update List"
   CmdAdd2List.Enabled = False
End If


End Sub

Private Sub CmdGo_Click()

tmp = "Select MC.CMPID,MC.CMPName,ML.LocationName,MC.Address,MC.TelephoneNo,MC.MobileNo,MC.FaxNo,MC.StartDate,MC.CloseDate,MC.Flag, "
tmp = tmp & " MC.CurrentDate , MC.EHoPCMPID "
tmp = tmp & " From Mst_CMP MC "
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID "
If ChkNoOpenGodown.Value = 1 Then
   tmp = tmp & " Left Join Mst_Godown MG On MC.CMPID = MG.CMPID"
End If
Dim wc, Wc1 As Variant
wc = ""
Wc1 = ""

If TxtSearchCMP <> "" Then
   wc = " Where MC.CMPName Like '%" & TxtSearchCMP & "%' "
   Wc1 = "CMPName Like '%" & TxtSearchCMP & "%'"
End If

If CmbLocationSearch.Text <> "" Then
   If TxtSearchCMP = "" Then
      wc = " Where ML.LocationName = '" & CmbLocationSearch.Text & "'"
      Wc1 = "LocationID = " & mSLocationID & ""
   Else
      wc = wc & " And ML.LocationName = '" & CmbLocationSearch.Text & "'"
      'wc1 = wc1 & "  LocationID = " & mSLocationID & "" 'changed by mayur on 3rd June 2013. AND was missing in condition
      Wc1 = Wc1 & " And LocationID = " & mSLocationID & ""
     
   End If
End If

If ChkNoGodown.Value = 1 Then
   If wc <> "" Then
      wc = wc & " And MC.CMPID Not in (Select Distinct CMPID from Mst_Godown) "
   Else
      wc = " Where MC.CMPID Not in (Select Distinct CMPID from Mst_Godown) "
   End If
End If

If ChkNoOpenGodown.Value = 1 Then
   If wc <> "" Then
      wc = wc & " And MC.CloseDate is null And MC.CMPID in (Select Distinct CMPID from Mst_Godown)"
      wc = wc & " Group by MC.CMPID,MC.CMPName,ML.LocationName,MC.Address,MC.TelephoneNo,MC.MobileNo,MC.FaxNo,MC.StartDate,MC.CloseDate,MC.Flag,"
      wc = wc & " mC.CurrentDate , mC.EHoPCMPID"
      wc = wc & " Having Count(MG.GodownID) = Count(MG.CloseDate)"
   Else
      wc = wc & " Where MC.CloseDate is null And MC.CMPID in (Select Distinct CMPID from Mst_Godown)"
      wc = wc & " Group by MC.CMPID,MC.CMPName,ML.LocationName,MC.Address,MC.TelephoneNo,MC.MobileNo,MC.FaxNo,MC.StartDate,MC.CloseDate,MC.Flag,"
      wc = wc & " mC.CurrentDate , mC.EHoPCMPID"
      wc = wc & " Having Count(MG.GodownID) = Count(MG.CloseDate)"
   End If
End If


tmp = tmp & wc

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        If InStr(MSHFlexGrid1.TextMatrix(iu, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(iu, C)) Then
              MSHFlexGrid1.TextMatrix(iu, C) = Format(MSHFlexGrid1.TextMatrix(iu, C), "dd-mmm-yyyy")
           Else
              MSHFlexGrid1.TextMatrix(iu, C) = MSHFlexGrid1.TextMatrix(iu, C)
           End If
        Else
          MSHFlexGrid1.TextMatrix(iu, C) = MSHFlexGrid1.TextMatrix(iu, C)
        End If
    Next
    If MSHFlexGrid1.TextMatrix(iu, 8) = "" Then
       MSHFlexGrid1.TextMatrix(iu, 9) = "Open"
    Else
       MSHFlexGrid1.TextMatrix(iu, 9) = "Close"
    End If
   
'    MSHFlexGrid1.TextMatrix(iu, 8) = IIf(IsNull(TmpRs.Fields(9)), "", TmpRs.Fields(9))
'    MSHFlexGrid1.TextMatrix(iu, 9) = IIf(IsNull(TmpRs!CurrentDate), TmpRs!StartDate, TmpRs!CurrentDate)
'    MSHFlexGrid1.TextMatrix(iu, 10) = IIf(IsNull(TmpRs!EHOPCMPID), TmpRs!StartDate, TmpRs!EHOPCMPID)
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount


RsMst_CMP.Filter = Wc1

End Sub

Private Sub CmdLicenseNo_Click()
FrmMst_License.Show
End Sub


Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

'Check1.Height = Label5.Height
'Check1.Left = Label5.Left
'Check1.Top = Label5.Top
'Check1.Width = Label5.Width
Call Grid_Head
TxtSearchCMP = ""
CmbLocationSearch.Text = ""
mSLocationID = 0
Label22.Caption = ""
ChkNoGodown.Value = 0

Call TableMst_CMP
Call GridHead2
If RsMst_CMP.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If

Call GButtonLock(Me, True)
RsMst_CMP.MoveLast
Call Indata
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_CMP.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_CMP.xls")
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

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub GodownCloseDate_LostFocus()

If GodownCloseDate <= GodownStartDate Then
   MsgBox "Close Date must be greater than Start Date !!! "
   GodownCloseDate.SetFocus
   Exit Sub
End If

If GodownCloseDate > Date Then
   MsgBox "Must not be future Date !!! "
   GodownCloseDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(GodownCloseDate.Value, "yyyy-mm") Then
   MsgBox "Godown Close Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   GodownCloseDate.SetFocus
   Exit Sub
End If

If mGodownID <> "" Then
   If ChkTransactionDate(mGodownID) = False Then
      MsgBox mMsgforDate
      GodownCloseDate.SetFocus
      Exit Sub
   End If
End If


End Sub
Private Function ChkTransactionDate(mGodownID_ As Variant) As Boolean
ChkTransactionDate = True
mMsgforDate = ""
'--CxTF Dates
tmp = " Select b.GodownID,Max(c.CxTFDate) 'TxnDate'"
tmp = tmp & " From Txn_CxTF_GSL a"
tmp = tmp & " Inner Join Txn_CxTF_Details c On a.CxTFNo = c.CxTFNo and a.SM_Prefix = c.SM_Prefix and a.TxnType = c.TxnType"
tmp = tmp & " Inner Join Mst_GSL b On a.GSLID = b.GSLID and a.SM_Prefix = b.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MG On b.GodownID = MG.GodownID"
tmp = tmp & " Where MG.CloseDate Is Null And MG.GodownID = " & mGodownID_ & "  "
tmp = tmp & " Group By b.GodownID "

Call TmpTable

If TmpRs.RecordCount > 0 Then
   If IsNull(TmpRs!TxnDate) = False Then
      If GodownCloseDate.Value < TmpRs!TxnDate Then
         mCxTFDate = TmpRs!TxnDate
         'MsgBox "One or More CDTF/CITF exist for selected godown beyond provided Close date. Can't Close on provided date!!!"
         mMsgforDate = "One or More CDTF/CITF exist for selected godown is " & mCxTFDate & " which is beyond provided Close date. Can't Close on provided date!!!" & vbCrLf
         ChkTransactionDate = False
         'Exit Function
      End If
   End If
End If

'mCxTFDate , mSpillageDate, mAdjTxnDate, mGSLTransferDate, mAgreeExpiryDate

'--Check of Spillage Date
tmp = " Select b.GodownID,Max(a.SpillageDate) 'SpillageDate'"
tmp = tmp & " From Txn_Spillage a "
tmp = tmp & " Inner Join Mst_GSL b On a.GSLID = b.GSLID and a.SM_Prefix = b.SM_Prefix "
tmp = tmp & " Inner Join Mst_Godown MG On b.GodownID = MG.GodownID "
tmp = tmp & " Where MG.CloseDate Is Null And MG.GodownID = " & mGodownID_ & " "
tmp = tmp & " Group By b.GodownID "

Call TmpTable

If TmpRs.RecordCount > 0 Then
   If IsNull(TmpRs!SpillageDate) = False Then
      If GodownCloseDate.Value < TmpRs!SpillageDate Then
         'MsgBox "One or More Spillage / Made up transaction exist for selected godown beyond provided Close date. Can't Close on provided date!!!"
         mSpillageDate = TmpRs!SpillageDate
         mMsgforDate = mMsgforDate & "One or More Spillage / Made up transaction exist for selected godown is " & mSpillageDate & " which is beyond provided Close date. Can't Close on provided date!!!" & vbCrLf
         ChkTransactionDate = False
         'Exit Function
      End If
   End If
End If


'---Check on Adjustment Date
tmp = " Select MG.GodownID,Max(TGAD.TxnGSLAdjDat) 'AdjDate'"
tmp = tmp & " From Txn_GSLAdjTransaction TGAD"
tmp = tmp & " Inner Join Txn_AdjTransaction TA On TGAD.TxnID = TA.TxnID"
tmp = tmp & " Inner Join Mst_GSL MG On TA.SM_Prefix = MG.SM_Prefix And TGAD.GSLID = MG.GSLID"
tmp = tmp & " Inner Join Mst_Godown G On MG.GodownID = G.GodownID"
tmp = tmp & " Where G.CloseDate Is Null And G.GodownID = " & mGodownID_ & ""
tmp = tmp & " Group By MG.GodownID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   If IsNull(TmpRs!AdjDate) = False Then
      If GodownCloseDate.Value < TmpRs!AdjDate Then
         'MsgBox "One or More Adjustment transaction exist for selected godown beyond provided Close date. Can't Close on provided date!!!"
         mAdjTxnDate = TmpRs!AdjDate
         mMsgforDate = mMsgforDate & "One or More Adjustment transaction exist for selected godown is " & mAdjTxnDate & " which is beyond provided Close date. Can't Close on provided date!!!" & vbCrLf
         ChkTransactionDate = False

         'Exit Function
      End If
   End If
End If



'--Check on Transfer Date
tmp = " Select b.GodownID,Max(a.GSLTransferDate) 'GSLTransferDate'"
tmp = tmp & " From Txn_GSL_Transfer a"
tmp = tmp & " Inner Join Mst_GSL b On a.OldGSLID = b.GSLID and a.SM_Prefix = b.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown G On b.GodownID = G.GodownID"
tmp = tmp & " Where G.CloseDate Is Null And G.GodownID = " & mGodownID_ & ""
tmp = tmp & " Group By b.GodownID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   If IsNull(TmpRs!GSLTransferDate) = False Then
      If GodownCloseDate.Value < TmpRs!GSLTransferDate Then
         'MsgBox "One or More GSL Transfer transaction exist for selected godown beyond provided Close date. Can't Close on provided date!!!"
         mGSLTransferDate = TmpRs!GSLTransferDate
         mMsgforDate = mMsgforDate & "One or More GSL Transfer transaction exist for selected godown is " & mGSLTransferDate & " which is beyond provided Close date. Can't Close on provided date!!!" & vbCrLf
         ChkTransactionDate = False
         
         'Exit Function
      End If
   End If
End If


'--Check on Lessor Agreement Date
tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA"
tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID"
tmp = tmp & " Where TA.Flag not in ('R') And TGA.GodownID = " & mGodownID_ & ""
tmp = tmp & " Group BY TGA.GodownID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   If IsNull(TmpRs!ExpDate) = False Then
      If GodownCloseDate.Value < TmpRs!ExpDate Then
'         MsgBox "One or More Live Lessor Agreement exist for selected godown beyond provided Close date. Can't Close on provided date!!!"
         mAgreeExpiryDate = TmpRs!ExpDate
         mMsgforDate = mMsgforDate & "One or More Live Lessor Agreement exist for selected godown is " & mAgreeExpiryDate & " which is beyond provided Close date. Can't Close on provided date!!!" & vbCrLf
         ChkTransactionDate = False
         
         'Exit Function
      End If
   End If
End If

End Function

Private Sub GodownStartDate_LostFocus()
If GodownStartDate.Value < StartDate.Value Then
  MsgBox "Godown Start Date Must Be After CMP Start Date"
  GodownStartDate.SetFocus
  Exit Sub
End If
If Check1.Value = 1 Then
   If GodownCloseDate <= GodownStartDate Then
      MsgBox "Start  Date must be Less than Close Date !!! "
      GodownStartDate.SetFocus
      Exit Sub
   End If
End If



'TxtAgreementEndDate.Value = GodownStartDate.Value + 330
TxtAgreementEndDate.Value = DateAdd("M", 11, GodownStartDate.Value)
TxtLeasedUpto.Value = DateAdd("M", 3, GodownStartDate.Value)
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
'Dim i As Variant
I = MSHFlexGrid2.RowSel
If I <= 1 Then Exit Sub

If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit Sub
'RsMst_Godown.MoveFirst
'RsMst_CMP.Find "CMPID = " & MSHFlexGrid1.TextMatrix(i, 0)
'If Not RsMst_CMP.EOF Then
   'Call InData
   'Frame0.Visible = False
 If MSHFlexGrid2.TextMatrix(I, 0) <> "" Then
    mGodownID = MSHFlexGrid2.TextMatrix(I, 0)
 Else
    mGodownID = ""
 End If
 If LCase(Gen_UserRole) <> "power" Then
    If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
       MsgBox "Godown closed you can not edit. Please contact Power users!!!"
       Exit Sub
    End If
 End If
   
   CmbUnitTypeID.Text = MSHFlexGrid2.TextMatrix(I, 1)
   If LCase(MSHFlexGrid2.TextMatrix(I, 1)) = "leased franchisee" Then
      lblleaseduptodate.Visible = True
      TxtLeasedUpto.Visible = True
      If MSHFlexGrid2.TextMatrix(I, 23) <> "" Then
         TxtLeasedUpto.Value = MSHFlexGrid2.TextMatrix(I, 23)
      Else
         TxtLeasedUpto.Value = Date
      End If
   Else
      TxtLeasedUpto.Visible = False
      lblleaseduptodate.Visible = False
   End If
     
   mUnitTypeID = GetUnitID(CmbUnitTypeID.Text)
   If Edit_Flg = "E" Then
      CmbUnitTypeID.Enabled = False
   ElseIf Edit_Flg = "A" Then
      CmbUnitTypeID.Enabled = True
   End If
   CmbCategoryID.Text = MSHFlexGrid2.TextMatrix(I, 2)
   mCategoryID = GetCategoryID(CmbCategoryID.Text)
   TxtGodownNo.Text = MSHFlexGrid2.TextMatrix(I, 3)
   TxtGodownAddress.Text = MSHFlexGrid2.TextMatrix(I, 4)
   TxtGodownTelephoneNo.Text = MSHFlexGrid2.TextMatrix(I, 5)
   TxtGodownArea.Text = MSHFlexGrid2.TextMatrix(I, 6)
   TxtBagsCapacity.Text = MSHFlexGrid2.TextMatrix(I, 7)
   TxtGodownCapacity.Text = MSHFlexGrid2.TextMatrix(I, 8)
   TxtStock.Text = MSHFlexGrid2.TextMatrix(I, 19)
   GodownStartDate = MSHFlexGrid2.TextMatrix(I, 9)
   Check1.Value = 0
   If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
      Check1.Value = 1
      GodownCloseDate.Value = MSHFlexGrid2.TextMatrix(I, 10)
   End If
   CmbLicenseNo.Text = MSHFlexGrid2.TextMatrix(I, 11)
   mLicenseID = IIf((MSHFlexGrid2.TextMatrix(I, 12)) = "", 0, MSHFlexGrid2.TextMatrix(I, 12))
      
   If MSHFlexGrid2.TextMatrix(I, 15) <> "" Then
      TxtAgreementEndDate = MSHFlexGrid2.TextMatrix(I, 15)
      If mUnitTypeID = 1 Then
         TxtAgreementEndDate.Enabled = False
      Else
         TxtAgreementEndDate.Enabled = True
      End If
   Else
      If mUnitTypeID = 1 Then
         TxtAgreementEndDate.Visible = False
         Label9.Visible = False
      Else
         TxtAgreementEndDate.Visible = True
         Label9.Visible = True
      End If
   End If
   TxtAmount = MSHFlexGrid2.TextMatrix(I, 16)
   TxtAdvAmount = MSHFlexGrid2.TextMatrix(I, 17)
   TxtAccreditedQty = MSHFlexGrid2.TextMatrix(I, 18)
   If MSHFlexGrid2.TextMatrix(I, 22) <> "" Then
       CmbNCDEXAcccreditedQty = MSHFlexGrid2.TextMatrix(I, 20)
       mCapacity = Val(MSHFlexGrid2.TextMatrix(I, 21))
       mAccreditedID = Val(MSHFlexGrid2.TextMatrix(I, 22))
   Else
       CmbNCDEXAcccreditedQty = ""
       mCapacity = ""
       mAccreditedID = ""
   End If
   TxtCEN.Text = MSHFlexGrid2.TextMatrix(I, 24)
   If MSHFlexGrid2.TextMatrix(I, 25) <> "" Then
      TxtCENDate = MSHFlexGrid2.TextMatrix(I, 25)
   Else
      TxtCENDate.Value = Date
   End If
   TxtCOEN.Text = MSHFlexGrid2.TextMatrix(I, 26)
   If MSHFlexGrid2.TextMatrix(I, 27) <> "" Then
      TxtCOENDate = MSHFlexGrid2.TextMatrix(I, 27)
   Else
      TxtCOENDate.Value = Date
   End If
   ChkYearlyBilling.Value = Val(MSHFlexGrid2.TextMatrix(I, 28))
   
   Call CmbLicenseNo_GotFocus
   Call CmbLicenseNo_LostFocus
'End If

CmdAdd2List.Caption = "Update List"
CmdAdd2List.Enabled = True

'added after godown inspection on 19 apr 2013
If LCase(Gen_UserRole) = "power" Then
   CmbCategoryID.Enabled = True
   TxtGodownNo.Enabled = True
   CmbLicenseNo.Enabled = True
   GodownStartDate.Enabled = True
   TxtAgreementEndDate.Enabled = True
   TxtLeasedUpto.Enabled = True
   TxtGodownArea.Enabled = True
   TxtBagsCapacity.Enabled = True
   TxtGodownCapacity.Enabled = True
Else
   CmbCategoryID.Enabled = False
   TxtGodownNo.Enabled = False
   CmbLicenseNo.Enabled = True
   GodownStartDate.Enabled = False
   TxtAgreementEndDate.Enabled = False
   TxtLeasedUpto.Enabled = False
   TxtGodownArea.Enabled = False
   TxtBagsCapacity.Enabled = False
   TxtGodownCapacity.Enabled = False
End If
'Call Check1_Click
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_CMP.MoveNext
If RsMst_CMP.EOF Then
   RsMst_CMP.MoveLast
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
RsMst_CMP.MovePrevious
If RsMst_CMP.BOF Then
   RsMst_CMP.MoveFirst
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
RsMst_CMP.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_CMP.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsMst_CMP.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If
 
Frame0.Height = Frame4.Height - 400
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 200
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 400
MSHFlexGrid1.Height = Frame0.Height - 1400
If RsMst_CMP.RecordCount <= 0 Then Exit Sub
RsMst_CMP.Filter = ""
RsMst_CMP.MoveFirst

'MSHFlexGrid1.Rows = RsMst_CMP.RecordCount + 1
'MSHFlexGrid1.ColWidth(0) = 1000

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchCMP = ""
'CmbLocationSearch.Text = ""
'mSLocationId = 0
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub StartDate_LostFocus()

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(StartDate.Value, "yyyy-mm") Then
   MsgBox "CMP Start Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   StartDate.SetFocus
   Exit Sub
End If


End Sub

Private Sub TxtAccreditedQty_LostFocus()
If TxtAccreditedQty = "" Then Exit Sub
If Not IsNumeric(TxtAccreditedQty) Then
   MsgBox "Invalid Number Foramt !!!!"
   TxtAccreditedQty.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtAdvAmount_LostFocus()
If TxtAdvAmount = "" Then Exit Sub
If Not IsNumeric(TxtAdvAmount) Then
   MsgBox "Invalid Number Foramt !!!!"
   TxtAdvAmount.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtAmount_LostFocus()
If TxtAmount = "" Then Exit Sub
If Not IsNumeric(TxtAmount) Then
   MsgBox "Invalid Number Foramt !!!!"
   TxtAmount.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtCMPName_LostFocus()
TxtCMPName = Trim(TxtCMPName)
End Sub
Private Sub TxtCMPDispalyName_LostFocus()
TxtCMPDispalyName = Trim(TxtCMPDispalyName)
End Sub
Private Sub TxtGodownNo_LostFocus()
TxtGodownNo = Trim(TxtGodownNo)
End Sub

Private Sub TxtSearchCMP_LostFocus()
'Dim x As Integer
'RsMst_CMP.MoveFirst
'Call Grid_Head
'If TxtSearchCMP <> "" Then
'    RsMst_CMP.Find ("CMPID = " & Trim(TxtSearchCMP))
'    If RsMst_CMP.EOF Then
'      MsgBox "No Such CMP Exist in DataBase!!!"
'      TxtSearchCMP.SetFocus
'      Exit Sub
'    End If
'  x = 1
'  Call ShowGrid(x)
'Else
' 'RsMst_ISIN.Sort
' MSHFlexGrid1.Rows = RsMst_CMP.RecordCount + 1
'
' For x = 1 To RsMst_CMP.RecordCount
'     Call ShowGrid(x)
'     RsMst_CMP.MoveNext
' Next

'End If

End Sub


Public Sub Indata()
mChkCloseGodown = False
TxtLeasedUpto.Visible = False
lblleaseduptodate.Visible = False
Check1.Visible = False
Label5.Visible = False
GodownCloseDate.Visible = False
Check2.Visible = False
LblCloseDate.Visible = False
CloseDate.Visible = False
Frame2.Enabled = False
TxtCMPID.Locked = True
TxtCMPName.Locked = True
TxtCMPDispalyName.Locked = True
TxtLocationID.Locked = True
CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True
TxtAddress.Locked = True
TxtTelephoneNo.Locked = True
TxtMobileNo.Locked = True
TxtFaxNo.Locked = True
StartDate.Enabled = False
CloseDate.Enabled = False
TxtUnitTypeID.Locked = True
CmbUnitTypeID.Visible = False
CmdUnitTypeID.Visible = False
TxtUnitTypeID.Visible = True
TxtCategoryID.Locked = True
CmbCategoryID.Visible = False
CmdCategoryID.Visible = False
CmbLicenseNo.Locked = True
CmdLicenseNo.Visible = False
TxtCategoryID.Visible = True
TxtGodownNo.Locked = True
TxtBagsCapacity.Locked = True
TxtGodownCapacity.Locked = True
TxtGodownAddress.Locked = True
TxtGodownArea.Locked = True
TxtGodownTelephoneNo.Locked = True
GodownStartDate.Enabled = False
Label10.Visible = False
GodownCloseDate.Enabled = False
TxtEHoPCMPID.Locked = True
'TxtFlag.Locked = True
TxtAgreementEndDate.Visible = True
Label9.Visible = True
TxtCMPID = IIf(IsNull(RsMst_CMP!CMPID), "", RsMst_CMP!CMPID)
TxtEHoPCMPID = IIf(IsNull(RsMst_CMP!EHOPCMPID), "", RsMst_CMP!EHOPCMPID)

mCMPID = TxtCMPID
TxtCMPName = IIf(IsNull(RsMst_CMP!CMPName), "", RsMst_CMP!CMPName)
TxtCMPDispalyName = IIf(IsNull(RsMst_CMP!CMPAlias), "", RsMst_CMP!CMPAlias)
'mAssayerID = IIf(IsNull(RsMst_CMP!AssayerID), "", RsMst_CMP!AssayerID)
mLocationID = IIf(IsNull(RsMst_CMP!LocationID), "", RsMst_CMP!LocationID)
mSM_Prefix = GetPreFix(RsMst_CMP!CMPID, "P") 'a

TxtLocationID = GetLocationName(RsMst_CMP!LocationID) '(IsNull(RsMst_CMP!LocationName), "", RsMst_CMP!LocationName)

TxtAddress = IIf(IsNull(RsMst_CMP!Address), "", RsMst_CMP!Address)
TxtTelephoneNo = IIf(IsNull(RsMst_CMP!TelephoneNo), "", RsMst_CMP!TelephoneNo)
TxtMobileNo = IIf(IsNull(RsMst_CMP!MobileNo), "", RsMst_CMP!MobileNo)
TxtFaxNo = IIf(IsNull(RsMst_CMP!FaxNo), "", RsMst_CMP!FaxNo)
StartDate.Value = IIf(IsNull(RsMst_CMP!StartDate), "", RsMst_CMP!StartDate)
TxtBlockDays = IIf(IsNull(RsMst_CMP!BlockDays), "", RsMst_CMP!BlockDays)
TxtFlag = IIf(IsNull(RsMst_CMP!Flag), "", RsMst_CMP!Flag)

TxtCreated = IIf(IsNull(RsMst_CMP!CreatedBy), "", RsMst_CMP!CreatedBy) & " :- " & IIf(IsNull(RsMst_CMP!CreatedDate), "", RsMst_CMP!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_CMP!UpdatedBy), "", RsMst_CMP!UpdatedBy) & " :- " & IIf(IsNull(RsMst_CMP!UpdatedDate), "", RsMst_CMP!UpdatedDate)

Call TableMst_Godown("where CMPID = " & mCMPID)
'Call TableMst_GSL

'tmp = "SELECT Sum(Mst_GSL.DepositeBags) AS DepositeBags, Sum(Mst_GSL.DepositeWeight) AS DepositeWeight, Sum(Mst_GSL.WHRBags) AS WHRBags, " & _
'      "Sum(Mst_GSL.WHRWeight) AS WHRWeight, Sum(Mst_GSL.CDTFBags) AS CDTFBags, Sum(Mst_GSL.CDTFWeight) AS CDTFWeight " & _
'      "FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
      
      
tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA "
tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
tmp = tmp & " Where TA.Flag not in ('R') "
tmp = tmp & " Group BY TGA.GodownID "
      
Call Tmp3Table

Call ClearControl
Call GridHead2
Dim I As Variant
I = 1
For ix = 1 To RsMst_Godown.RecordCount
    I = I + 1
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    MSHFlexGrid2.TextMatrix(I, 0) = RsMst_Godown!GodownID
    MSHFlexGrid2.TextMatrix(I, 1) = GetUnitName(RsMst_Godown!UnitTypeID)
    MSHFlexGrid2.TextMatrix(I, 2) = GetCategoryName(RsMst_Godown!CategoryID)
    MSHFlexGrid2.TextMatrix(I, 3) = RsMst_Godown!GodownNo
    MSHFlexGrid2.TextMatrix(I, 4) = RsMst_Godown!Address
    MSHFlexGrid2.TextMatrix(I, 5) = RsMst_Godown!TelephoneNo
    MSHFlexGrid2.TextMatrix(I, 6) = RsMst_Godown!Area
    MSHFlexGrid2.TextMatrix(I, 7) = RsMst_Godown!BagsCapacity
    MSHFlexGrid2.TextMatrix(I, 8) = RsMst_Godown!GodownCapacity
    MSHFlexGrid2.TextMatrix(I, 9) = RsMst_Godown!StartDate
    MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(RsMst_Godown!CloseDate), "", RsMst_Godown!CloseDate)
    MSHFlexGrid2.TextMatrix(I, 11) = GetLicenseNo(IIf(IsNull(RsMst_Godown!LicenseID), 0, RsMst_Godown!LicenseID))
    MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(RsMst_Godown!LicenseID), "", RsMst_Godown!LicenseID)
    MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(RsMst_Godown!LicenseID), "", RsMst_Godown!LicenseID)
    MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(RsMst_Godown!AgreementEndDate), "", RsMst_Godown!AgreementEndDate)
    If RsMst_Godown!UnitTypeID = 1 Then
       If TmpRs3.RecordCount > 0 Then
          TmpRs3.MoveFirst
          TmpRs3.Find "GodownID = " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & ""
          If TmpRs3.EOF = False Then
             MSHFlexGrid2.TextMatrix(I, 15) = Format(TmpRs3!ExpDate, "dd-mmm-yyyy")
          Else
             MSHFlexGrid2.TextMatrix(I, 15) = "" 'Format(TmpRs3!ExpDate, "dd-mmm-yyyy")
          End If
       End If
    End If
      
    MSHFlexGrid2.TextMatrix(I, 16) = IIf(IsNull(RsMst_Godown!Amount), "", RsMst_Godown!Amount)
    MSHFlexGrid2.TextMatrix(I, 17) = IIf(IsNull(RsMst_Godown!AdvAmount), "", RsMst_Godown!AdvAmount)
    MSHFlexGrid2.TextMatrix(I, 18) = IIf(IsNull(RsMst_Godown!NCDEXAccreditedQty), "", RsMst_Godown!NCDEXAccreditedQty)
     
    MSHFlexGrid2.TextMatrix(I, 24) = IIf(IsNull(RsMst_Godown!CPEN), "", RsMst_Godown!CPEN)
    MSHFlexGrid2.TextMatrix(I, 25) = IIf(IsNull(RsMst_Godown!CPED), "", RsMst_Godown!CPED)
    MSHFlexGrid2.TextMatrix(I, 26) = IIf(IsNull(RsMst_Godown!COPEN), "", RsMst_Godown!COPEN)
    MSHFlexGrid2.TextMatrix(I, 27) = IIf(IsNull(RsMst_Godown!COPED), "", RsMst_Godown!COPED)
    MSHFlexGrid2.TextMatrix(I, 28) = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), 0, RsMst_Godown!YearlyBillingFlg)
    
    tmp = "SELECT Sum(Mst_GSL.DepositeBags) AS DepositeBags, Sum(Mst_GSL.DepositeWeight) AS DepositeWeight, Sum(Mst_GSL.WHRBags) AS WHRBags, " & _
          "Sum(Mst_GSL.WHRWeight) AS WHRWeight, Sum(Mst_GSL.CDTFBags) AS CDTFBags, Sum(Mst_GSL.CDTFWeight) AS CDTFWeight " & _
          "FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
          " WHERE ((Mst_Godown.GodownID)=" & RsMst_Godown!GodownID & ") And Mst_GSL.Flag <> 'E' And Mst_GSL.Flag <> 'Z' "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
       If IsNull(TmpRs!DepositeWeight) = False Then
          MSHFlexGrid2.TextMatrix(I, 13) = Val(TmpRs!DepositeWeight) - Val(IIf(IsNull(TmpRs!CDTFWeight), 0, TmpRs!CDTFWeight))
       Else
          MSHFlexGrid2.TextMatrix(I, 13) = 0
       End If
       
       If IsNull(TmpRs!DepositeWeight) = False Then
          MSHFlexGrid2.TextMatrix(I, 19) = Val(TmpRs!DepositeBags) - Val(IIf(IsNull(TmpRs!CDTFBags), 0, TmpRs!CDTFBags))
       Else
          MSHFlexGrid2.TextMatrix(I, 19) = 0
       End If
    End If
     
    If IsNull(RsMst_Godown!AccreditedID) = False Then
       tmp = "Select AccreditedDescription, Capacity from Mst_NCDEXAccreditedCapacity Where AccreditedID = " & RsMst_Godown!AccreditedID
       Call TmpTable
       If TmpRs.RecordCount > 0 Then
          MSHFlexGrid2.TextMatrix(I, 20) = TmpRs!AccreditedDescription
          MSHFlexGrid2.TextMatrix(I, 21) = TmpRs!Capacity
          MSHFlexGrid2.TextMatrix(I, 22) = RsMst_Godown!AccreditedID
       End If
    End If
    MSHFlexGrid2.TextMatrix(I, 23) = IIf(IsNull(RsMst_Godown!LeasedUptoDate), "", RsMst_Godown!LeasedUptoDate)
    RsMst_Godown.MoveNext
Next

For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 10) = "" Then
       MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 6))
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(I, 7))
       MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(I, 8))
       MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(0, 13)) + Val(MSHFlexGrid2.TextMatrix(I, 13))
       MSHFlexGrid2.TextMatrix(0, 19) = Val(MSHFlexGrid2.TextMatrix(0, 19)) + Val(MSHFlexGrid2.TextMatrix(I, 19))
    End If
Next
Check2.Value = 0
If Not IsNull(RsMst_CMP!CloseDate) Then
   CloseDate.Value = IIf(IsNull(RsMst_CMP!CloseDate), "", RsMst_CMP!CloseDate)
   Check2.Value = 1
   Check2.Visible = True
   Check2.Enabled = False
   CloseDate.Enabled = False
   CloseDate.Visible = True
End If

'MSHFlexGrid2.Enabled = False
If RsMst_CMP.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
CmdAdd2List.Enabled = False
Call ClearGodownContent
End Sub
Private Sub TxtBagsCapacity_LostFocus()
  If TxtBagsCapacity <> "" Then
  If IsNumeric(TxtBagsCapacity) = False Then
     MsgBox "Invalid numeric format !!!!"
     TxtBagsCapacity.SetFocus
     Exit Sub
   End If
  End If
End Sub

Private Sub TxtCMPID_LostFocus()
If TxtCMPID = "" Then Exit Sub
If IsNumeric(TxtCMPID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtCMPID.SetFocus
   Exit Sub
End If
TxtCMPID = Val(TxtCMPID)
End Sub
Private Sub TxtStartDate_LostFocus()
End Sub
Private Sub TxtCloseDate_LostFocus()
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
mChkCloseGodown = False
Call GButtonLock(Me, False)
Check1.Visible = False
TxtLeasedUpto.Visible = False
lblleaseduptodate.Visible = False
Check2.Visible = False
LblCloseDate.Visible = False
CloseDate.Visible = False
Label10.Visible = False
TxtCMPID.Locked = True
TxtCMPName.Locked = False
TxtCMPDispalyName.Locked = False
TxtLocationID.Locked = False
CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
TxtAddress.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False
TxtFaxNo.Locked = False
StartDate.Enabled = True
'CloseDate.Enabled = True
Frame2.Enabled = True
TxtUnitTypeID.Locked = False
CmbUnitTypeID.Visible = True
CmbUnitTypeID.Enabled = True
CmdUnitTypeID.Visible = True
TxtUnitTypeID.Visible = False
TxtCategoryID.Locked = False
CmbCategoryID.Visible = True
CmdCategoryID.Visible = True
TxtCategoryID.Visible = False
CmbLicenseNo.Locked = False
CmdLicenseNo.Visible = True
TxtGodownNo.Locked = False
TxtEHoPCMPID.Locked = False
TxtBagsCapacity.Locked = False
TxtGodownCapacity.Locked = False
TxtGodownAddress.Locked = False
TxtGodownArea.Locked = False
TxtGodownTelephoneNo.Locked = False
GodownStartDate.Enabled = True
GodownCloseDate.Enabled = True
''comment after adding godown inspection
If LCase(Gen_UserRole) = "power" Then
   CmdAdd2List.Enabled = True
Else
   CmdAdd2List.Caption = "Update List"
   CmdAdd2List.Enabled = False
End If

'
TxtBlockDays = 7
Call GridHead2
'MSHFlexGrid2.Enabled = True
TxtFlag.Locked = False

StartDate.Enabled = True
CmbLocationID.Enabled = True

StartDate.Value = Date
TxtCMPID = ""
TxtCMPName = ""
TxtCMPDispalyName = ""
TxtLocationID = ""
TxtAddress = ""
TxtTelephoneNo = ""
TxtMobileNo = ""
TxtFaxNo = ""

TxtCreated = ""
TxtUpdated = ""

CmbLocationID.Text = ""
TxtEHoPCMPID = ""
'StartDate.SetFocus
Call ClearGodownContent
Check2.Value = 0
If RsMst_CMP.RecordCount > 0 Then
   TxtEHoPCMPID.SetFocus
End If
TxtFlag = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
If Check2.Value = 1 Then
   MsgBox "CMP is Closed !!!"
   Exit Sub
End If
Call GButtonLock(Me, False)
Frame2.Enabled = True
Check1.Visible = True
Check2.Enabled = True
Label10.Visible = True
Check2.Visible = True
TxtCMPID.Locked = True
TxtCMPName.Locked = False
TxtCMPDispalyName.Locked = False
TxtLocationID.Locked = False
CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
TxtEHoPCMPID.Locked = False
TxtAddress.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False
TxtFaxNo.Locked = False
StartDate.Enabled = True
'CloseDate.Enabled = True
TxtUnitTypeID.Locked = False
CmbUnitTypeID.Visible = True
CmdUnitTypeID.Visible = True
TxtUnitTypeID.Visible = False
TxtCategoryID.Locked = False
CmbCategoryID.Visible = True
CmdCategoryID.Visible = True
TxtCategoryID.Visible = False
CmbLicenseNo.Locked = False
CmdLicenseNo.Visible = True
TxtGodownNo.Locked = False
TxtBagsCapacity.Locked = False
TxtGodownCapacity.Locked = False
TxtGodownAddress.Locked = False
TxtGodownArea.Locked = False
TxtGodownTelephoneNo.Locked = False
GodownStartDate.Enabled = True
'GodownCloseDate.Enabled = True

'comment after adding godown inspection on 19 apr 2013
If LCase(Gen_UserRole) = "power" Then
   CmdAdd2List.Caption = "Add to List"
   CmdAdd2List.Enabled = True
Else
   CmdAdd2List.Caption = "Update List"
   CmdAdd2List.Enabled = False
End If


'MSHFlexGrid2.Enabled = True
TxtFlag.Locked = False
CmbLocationID.Text = TxtLocationID.Text

StartDate.Enabled = False

If MSHFlexGrid2.Rows > 3 Then
   CmbLocationID.Enabled = False
End If

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   
   tmp = "select *  from Mst_ISIN where CMPID = " & TxtCMPID
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "You can't delete this CMP,used in ISIN Master !!!! "
      Exit Sub
   End If
   
   tmp = "Delete from Mst_Godown where CMPID = " & TxtCMPID
   Call TmpTable
   
   tmp = "SELECT * from Mst_Godown where CMPID = " & TxtCMPID
   Call TmpTable
   TmpRs.Close
   
   RsMst_CMP.Requery
   RsMst_CMP.Find "cmpid = " & TxtCMPID
   If RsMst_CMP.EOF Then Exit Sub
   RsMst_CMP.Delete
   RsMst_CMP.Update
   If RsMst_CMP.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_CMP.MoveNext
   If RsMst_CMP.EOF() Then
      RsMst_CMP.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_CMP.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtEHoPCMPID.Text = "" Then
   MsgBox "Blank EHoP CMP ID Not Allowed !!! "
   TxtEHoPCMPID.SetFocus
   Exit Sub
End If
If TxtCMPName.Text = "" Then
   MsgBox "Blank CMPName Not Allowed !!! "
   TxtCMPName.SetFocus
   Exit Sub
End If
If TxtCMPDispalyName.Text = "" Then
   MsgBox "Blank CMP Dispaly Name Not Allowed !!! "
   TxtCMPDispalyName.SetFocus
   Exit Sub
End If
If CmbLocationID.Text = "" Then
   MsgBox "Blank LocationID Not Allowed !!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
If TxtAddress.Text = "" Then
   MsgBox "Blank Address Not Allowed !!! "
   TxtAddress.SetFocus
   Exit Sub
End If

'If TxtTelephoneNo.Text = "" Then
'   MsgBox "Blank TelephoneNo Not Allowed !!! "
'   TxtTelephoneNo.SetFocus
'   Exit Sub
'End If
'
'
'If TxtMobileNo.Text = "" Then
'   MsgBox "Blank MobileNo Not Allowed !!! "
'   TxtMobileNo.SetFocus
'   Exit Sub
'End If
If StartDate.Value = "" Then
   MsgBox "Blank StartDate Not Allowed !!! "
   TxtStartDate.SetFocus
   Exit Sub
End If
If CloseDate.Value = "" Then
   MsgBox "Blank CloseDate Not Allowed !!! "
   TxtCloseDate.SetFocus
   Exit Sub
End If
'If Edit_Flg = "E" And Check2.Value = 1 And (StartDate.Value > CloseDate) Then
'   MsgBox "Close Date Must Always Be Greater Than Close Date Always!!!"
'   CloseDate.SetFocus
'   Exit Sub
'End If
''''comment after adding godown inspection module on 19 apr 2013
'If MSHFlexGrid2.Rows <= 3 Then
'   MsgBox "CMP Should Have Atleast One Godown !!! "
'   CmbUnitTypeID.SetFocus
'   Exit Sub
'End If


'If ChkGodownCloseDate = False Then
'   MsgBox "One or more godown having close date more than cmp close date provided!!!"
'   CloseDate.SetFocus
'   Exit Sub
'End If

If Check2.Value = 1 And Val(MSHFlexGrid2.TextMatrix(0, 19)) <> 0 Then
   MsgBox "Can't Close This CMP it contains Bags Stock!!!"
   Check2.SetFocus
   Exit Sub
End If

If Check2.Value = 1 Then
   tmp = " select cmpid from txn_cxtf_details where txntype = 'W' and isnull(flag, '') <> 'P' and cmpid  = " & Val(TxtCMPID.Text)
   Call Tmp7Table
   If TmpRs7.RecordCount > 0 Then
      MsgBox "One or more CITFs not printed. Can't close this CMP!!!"
      Exit Sub
   End If
End If


mChkCloseGodown = False
If Edit_Flg = "E" Then
   If ChkAllGodownCloseStatus = True Then
      mChkCloseGodown = True
   End If
End If

If Edit_Flg = "A" Then
   If RsMst_CMP.RecordCount > 0 Then
      RsMst_CMP.MoveFirst
      RsMst_CMP.Find "CMPName = '" & TxtCMPName.Text & "'"
      If Not RsMst_CMP.EOF Then
        MsgBox "Duplicate EHoPCMPID Not Allowed !!! "
         TxtCMPName.SetFocus
         Exit Sub
      End If
      RsMst_CMP.MoveFirst
      RsMst_CMP.Find "EHoPCMPID = '" & TxtEHoPCMPID.Text & "'"
      If Not RsMst_CMP.EOF Then
         MsgBox "Duplicate EHoP CMP ID Not Allowed !!!"
         TxtEHoPCMPID.SetFocus
         Exit Sub
      End If
   End If
   RsMst_CMP.AddNew
   RsMst_CMP!CMPID = GetCMPID
   RsMst_CMP!G_UID = GetGUID
   RsMst_CMP!CMPLockedDate = GetMaxCMPLockedDate
   RsMst_CMP!CurrentDate = StartDate.Value
   RsMst_CMP!BlockDays = Val(TxtBlockDays)
   TxtCMPID.Text = RsMst_CMP!CMPID

Else
    RsMst_CMP.MoveFirst
    RsMst_CMP.Find "CMPName = '" & TxtCMPName.Text & "'"
    If Not RsMst_CMP.EOF Then
       If RsMst_CMP!CMPID <> Val(TxtCMPID.Text) Then
          MsgBox "Duplicate CMP Name Not Allowed !!! "
          TxtCMPID.SetFocus
          Exit Sub
       End If
       RsMst_CMP.MoveFirst
       RsMst_CMP.Find "EHoPCMPID = '" & TxtEHoPCMPID.Text & "'"
       If Not RsMst_CMP.EOF Then
          If RsMst_CMP!CMPID <> Val(TxtCMPID.Text) Then
             MsgBox "Duplicate EHoP CMP ID Not Allowed !!! "
             TxtEHoPCMPID.SetFocus
             Exit Sub
          End If
       End If
    End If
    RsMst_CMP.MoveFirst
    RsMst_CMP.Find " CMPID= " & TxtCMPID.Text
End If

'RsMst_CMP!CMPID = TxtCMPID.Text
RsMst_CMP!EHOPCMPID = TxtEHoPCMPID
RsMst_CMP!CMPName = TxtCMPName.Text
RsMst_CMP!CMPAlias = TxtCMPDispalyName


'RsMst_CMP!AssayerID = " " 'mAssayerID
RsMst_CMP!LocationID = mLocationID
RsMst_CMP!Address = TxtAddress.Text
RsMst_CMP!TelephoneNo = TxtTelephoneNo.Text
RsMst_CMP!MobileNo = TxtMobileNo.Text
RsMst_CMP!FaxNo = TxtFaxNo.Text
RsMst_CMP!StartDate = StartDate.Value
If Check2.Value = 1 Then
   RsMst_CMP!CloseDate = CloseDate.Value
Else
   RsMst_CMP!CloseDate = Null
End If

RsMst_CMP!Flag = "0" 'TxtFlag.Text

If IsNull(RsMst_CMP!CreatedBy) = True Or RsMst_CMP!CreatedBy = "" Then
   RsMst_CMP!CreatedBy = Gen_UserLoginID
   RsMst_CMP!CreatedDate = Now
Else
   RsMst_CMP!UpdatedBy = Gen_UserLoginID
   RsMst_CMP!UpdatedDate = Now
End If

RsMst_CMP.Update
mGridMsg = ""
Call TableMst_Godown("Where CMPID = " & TxtCMPID)
If MSHFlexGrid2.Rows > 3 Then
   For I = 2 To MSHFlexGrid2.Rows - 1
     If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
     
     If MSHFlexGrid2.TextMatrix(I, 0) = "" Then
     
        If mGridMsg = "" Then
            mGridMsg = mGridMsg & "<tr>"
            mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(1, 3) & " </td>"
            mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(1, 4) & " </td>"
            mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(1, 5) & " </td>"
            mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(1, 6) & " </td>"
            mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(1, 7) & " </td>"
            mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(1, 8) & " </td>"
            mGridMsg = mGridMsg & "<td> " & Format(MSHFlexGrid2.TextMatrix(1, 9), "dd-MMM-yyyy") & " </td>"
            mGridMsg = mGridMsg & "<td> " & Format(MSHFlexGrid2.TextMatrix(1, 15), "dd-MMM-yyyy") & " </td>"
            mGridMsg = mGridMsg & "</tr>"
        End If
        mGridMsg = mGridMsg & "<tr>"
        mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td>"
        mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 4) & " </td>"
        mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 5) & " </td>"
        mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 6) & " </td>"
        mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 7) & " </td>"
        mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 8) & " </td>"
        mGridMsg = mGridMsg & "<td> " & Format(MSHFlexGrid2.TextMatrix(I, 9), "dd-MMM-yyyy") & " </td>"
        mGridMsg = mGridMsg & "<td> " & Format(MSHFlexGrid2.TextMatrix(I, 15), "dd-MMM-yyyy") & " </td>"
        mGridMsg = mGridMsg & "</tr>"

        RsMst_Godown.AddNew
        RsMst_Godown!GodownID = GetMaxGodownID
        RsMst_Godown!Concurrency = Now
        RsMst_Godown!CreatedBy = Gen_UserLoginID
        RsMst_Godown!CreatedDate = Now
        
        RsMst_Godown!G_UID = GetGUID
     Else
        RsMst_Godown.MoveFirst
        RsMst_Godown.Find "GodownID = " & MSHFlexGrid2.TextMatrix(I, 0)
        RsMst_Godown!UpdatedBy = Gen_UserLoginID
        RsMst_Godown!UpdatedDate = Now
     End If
     RsMst_Godown!CMPID = TxtCMPID.Text
     RsMst_Godown!UnitTypeID = GetUnitID(MSHFlexGrid2.TextMatrix(I, 1))
     If LCase(MSHFlexGrid2.TextMatrix(I, 1)) = "leased franchisee" Then
        RsMst_Godown!LeasedUptoDate = MSHFlexGrid2.TextMatrix(I, 23)
        'TxtLeasedUpto.Visible = True
     Else
        RsMst_Godown!LeasedUptoDate = Null
        'TxtLeasedUpto.Visible = False
     End If
     RsMst_Godown!CategoryID = GetCategoryID(MSHFlexGrid2.TextMatrix(I, 2))
     RsMst_Godown!GodownNo = MSHFlexGrid2.TextMatrix(I, 3)
     RsMst_Godown!Address = MSHFlexGrid2.TextMatrix(I, 4)
     RsMst_Godown!TelephoneNo = MSHFlexGrid2.TextMatrix(I, 5)
     RsMst_Godown!Area = MSHFlexGrid2.TextMatrix(I, 6)
     RsMst_Godown!BagsCapacity = MSHFlexGrid2.TextMatrix(I, 7)
     RsMst_Godown!GodownCapacity = Val(MSHFlexGrid2.TextMatrix(I, 8))
     RsMst_Godown!StartDate = MSHFlexGrid2.TextMatrix(I, 9)
     If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
        RsMst_Godown!CloseDate = MSHFlexGrid2.TextMatrix(I, 10)
     Else
        RsMst_Godown!CloseDate = Null
     End If
     RsMst_Godown!LicenseID = Val(MSHFlexGrid2.TextMatrix(I, 12))
     Dim oldLicenseID As Variant
     
     oldLicenseID = IIf(MSHFlexGrid2.TextMatrix(I, 14) = "", 0, Val(MSHFlexGrid2.TextMatrix(I, 14)))
     If MSHFlexGrid2.TextMatrix(I, 15) <> "" Then
        RsMst_Godown!AgreementEndDate = MSHFlexGrid2.TextMatrix(I, 15)
     End If
     RsMst_Godown!Amount = Val(MSHFlexGrid2.TextMatrix(I, 16))
     RsMst_Godown!AdvAmount = Val(MSHFlexGrid2.TextMatrix(I, 17))
     RsMst_Godown!NCDEXAccreditedQty = Val(MSHFlexGrid2.TextMatrix(I, 18))
     If MSHFlexGrid2.TextMatrix(I, 22) <> "" Then
        RsMst_Godown!AccreditedID = Val(MSHFlexGrid2.TextMatrix(I, 22))
     Else
        RsMst_Godown!AccreditedID = Null
     End If
     
     RsMst_Godown!CPEN = MSHFlexGrid2.TextMatrix(I, 24)
     If MSHFlexGrid2.TextMatrix(I, 25) <> "" Then
         RsMst_Godown!CPED = MSHFlexGrid2.TextMatrix(I, 25)
     End If
     RsMst_Godown!COPEN = MSHFlexGrid2.TextMatrix(I, 26)
     If MSHFlexGrid2.TextMatrix(I, 27) <> "" Then
         RsMst_Godown!COPED = MSHFlexGrid2.TextMatrix(I, 27)
     End If
     RsMst_Godown!YearlyBillingFlg = Val(MSHFlexGrid2.TextMatrix(I, 28))
'     If Edit_Flg = "A" Then
'        tmp = " Select isnull(Max(AccreditedID),0) AccreditedID from  Mst_NCDEXAccreditedCapacity Where LocationID = " & mLocationID
'        Call TmpTable
'        If TmpRs!AccreditedID <> 0 Then
'            RsMst_Godown!AccreditedID = TmpRs!AccreditedID
'        End If
 '     End If
        

     RsMst_Godown.Update
     RsMst_Godown.Requery
     
     If mChkCloseGodown = True Then
        tmp = "Select Max(CloseDate) as CD From Mst_Godown Where CMPID = " & TxtCMPID & ""
        Call TmpTable
        
        If IsNull(TmpRs!CD) = False Then
           tmp = "Update Mst_CMP Set CloseDate = '" & CDate(TmpRs!CD) & "' Where CMPID = " & TxtCMPID & ""
           Call TmpTable
        End If
     End If
     
     If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
        Call CheckQQSMapping(MSHFlexGrid2.TextMatrix(I, 0))
     End If
     
'     tmp = "update Txn_WHR_Detail set LicenseID = " & Val(MSHFlexGrid2.TextMatrix(I, 12)) & " Where LicenseID = " & oldLicenseID
'     Call TmpTable
   
   Next I
End If
Call GButtonLock(Me, True)
RsMst_CMP.Requery
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPID = " & TxtCMPID & ""
If RsMst_CMP.EOF Then
   RsMst_CMP.MoveLast
End If

Call ClearGodownContent
Call Indata


If mGridMsg <> "" Then
   Call Gen_Email
End If

End Sub
Private Sub ClearGodownContent()

TxtUnitTypeID = ""
CmbUnitTypeID.Text = ""
TxtCategoryID.Text = ""
CmbCategoryID.Text = ""

TxtGodownNo = ""
TxtBagsCapacity = ""
TxtGodownCapacity = ""
TxtGodownAddress = ""
TxtGodownArea = ""
TxtGodownTelephoneNo = ""
GodownStartDate = Date
TxtAgreementEndDate.Value = GodownStartDate.Value + 330
'GodownCloseDate = Date
CmbNCDEXAcccreditedQty = ""
TxtCEN = ""
TxtCENDate.Value = Date
TxtCOEN = ""
TxtCOENDate.Value = Date

   
End Sub
Private Function GetCMPID() As Double
Dim Retval As Double
tmp = "Select max(CMPID) from Mst_CMP"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCMPID = Retval
End Function

Private Function GetMaxCMPLockedDate() As Double
Dim Retval As Double
tmp = "Select max(CMPLockedDate) from Mst_CMP"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
'RetVal = RetVal + 1
GetMaxCMPLockedDate = Retval
End Function

Private Function GetMaxGodownID() As Double
Dim Retval As Double
tmp = "Select max(GodownID) from Mst_Godown"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
TmpRs.Close
GetMaxGodownID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12
MSHFlexGrid1.TextMatrix(0, 0) = "CMP ID"
MSHFlexGrid1.TextMatrix(0, 1) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "Address"
MSHFlexGrid1.TextMatrix(0, 4) = "Telephone No"
MSHFlexGrid1.TextMatrix(0, 5) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 6) = "Fax No"
MSHFlexGrid1.TextMatrix(0, 7) = "Start Date"
MSHFlexGrid1.TextMatrix(0, 8) = "Close Date"
MSHFlexGrid1.TextMatrix(0, 9) = "Flag"
MSHFlexGrid1.TextMatrix(0, 10) = "Data Entered Up to "
MSHFlexGrid1.TextMatrix(0, 11) = "EHoP ID"
    
MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 3200
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200

MSHFlexGrid1.RowHeight(0) = 600
    
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_CMP.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub CmdAssayerID_Click()
    FrmMst_Assayer.Show
End Sub

Private Sub CmdLocationID_Click()
    FrmMst_Location.Show
End Sub
'----
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationID <> 0 ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
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
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub
'----
Private Sub CmdUnitTypeID_Click()
    FrmMst_UnitType.Show
End Sub
'----
Private Sub CmbUnitTypeID_GotFocus()
tmp = CmbUnitTypeID.Text
Call TableMst_UnitType
CmbUnitTypeID.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_UnitType.RecordCount
    CmbUnitTypeID.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbUnitTypeID.Text = tmp
End Sub
'----
Private Sub CmbUnitTypeID_LostFocus()
If CmbUnitTypeID.Text = "" Then
   Exit Sub
End If
RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbUnitTypeID.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid selection "
   CmbUnitTypeID.SetFocus
   Exit Sub
End If
mUnitTypeID = RsMst_UnitType!UnitTypeID
If RsMst_UnitType!Flag = "Fixed" Then
   Label11.Caption = "Lease Amount"
   TxtAdvAmount.Locked = False
Else
   Label11.Caption = "Income Sharing %"
   TxtAdvAmount.Locked = True
   TxtAdvAmount = ""
End If
If LCase(CmbUnitTypeID.Text) = "leased franchisee" Then
   TxtLeasedUpto.Value = DateAdd("M", 3, GodownStartDate.Value)
   TxtLeasedUpto.Visible = True
   lblleaseduptodate.Visible = True
Else
   TxtLeasedUpto.Visible = False
   lblleaseduptodate.Visible = False
End If

End Sub
Private Sub CmdCategoryID_Click()
    FrmMst_Category.Show
End Sub
'----
Private Sub CmbCategoryID_GotFocus()
tmp = CmbCategoryID.Text
Call TableMst_Category
CmbCategoryID.Clear
If RsMst_Category.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Category.RecordCount
    CmbCategoryID.AddItem RsMst_Category!Category
    RsMst_Category.MoveNext
Next
CmbCategoryID.Text = tmp
End Sub
'----
Private Sub CmbCategoryID_LostFocus()
If CmbCategoryID.Text = "" Then
   Exit Sub
End If
RsMst_Category.MoveFirst
RsMst_Category.Find "Category = '" & CmbCategoryID.Text & "'"
If RsMst_Category.EOF Then
   MsgBox "Invalid selection "
   CmbCategoryID.SetFocus
   Exit Sub
End If
mCategoryID = RsMst_Category!CategoryID
End Sub
'----


Private Sub GridHead2()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.Cols = 29 '28

MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.FixedCols = 1

MSHFlexGrid2.TextMatrix(0, 0) = "Total"
MSHFlexGrid2.TextMatrix(1, 0) = "GodownID"
MSHFlexGrid2.TextMatrix(1, 1) = "Unit Type"
MSHFlexGrid2.TextMatrix(1, 2) = "Category"
MSHFlexGrid2.TextMatrix(1, 3) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 4) = "Godown Address"
MSHFlexGrid2.TextMatrix(1, 5) = "Godown TelephoneNo"
MSHFlexGrid2.TextMatrix(1, 6) = "Area in Sq.Ft."
MSHFlexGrid2.TextMatrix(1, 7) = "Capacity in bags"
MSHFlexGrid2.TextMatrix(1, 8) = "Capacity in MT"
MSHFlexGrid2.TextMatrix(1, 9) = "Start Date"
MSHFlexGrid2.TextMatrix(1, 10) = "Close Date"
MSHFlexGrid2.TextMatrix(1, 11) = "License No"
MSHFlexGrid2.TextMatrix(1, 12) = "License ID"
MSHFlexGrid2.TextMatrix(1, 13) = "Stock"
MSHFlexGrid2.TextMatrix(1, 14) = "Old-License ID"

MSHFlexGrid2.TextMatrix(1, 15) = "Agreement End Date"
MSHFlexGrid2.TextMatrix(1, 16) = "Amount"
MSHFlexGrid2.TextMatrix(1, 17) = "Advance Amount"
MSHFlexGrid2.TextMatrix(1, 18) = "NCDEX Accredited Qty"
MSHFlexGrid2.TextMatrix(1, 19) = "Bags Stock"
MSHFlexGrid2.TextMatrix(1, 20) = "Accredited Description"
MSHFlexGrid2.TextMatrix(1, 21) = "NCDEX Accredited Capacity"
MSHFlexGrid2.TextMatrix(1, 22) = "NCDEX Accredited ID"
MSHFlexGrid2.TextMatrix(1, 23) = "Leased Upto Date"

MSHFlexGrid2.TextMatrix(1, 24) = "Cotton Endorsement no "
MSHFlexGrid2.TextMatrix(1, 25) = "CPE Date"
MSHFlexGrid2.TextMatrix(1, 26) = "Commodity Endorement no"
MSHFlexGrid2.TextMatrix(1, 27) = "COPE Date"
MSHFlexGrid2.TextMatrix(1, 28) = "Yearly Billing Flag"

'MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 1000
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 2800
MSHFlexGrid2.ColWidth(5) = 2000
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1200
MSHFlexGrid2.ColWidth(8) = 1200
MSHFlexGrid2.ColWidth(9) = 1200
MSHFlexGrid2.ColWidth(10) = 1000
MSHFlexGrid2.ColWidth(11) = 1000
MSHFlexGrid2.ColWidth(12) = 0
MSHFlexGrid2.ColWidth(13) = 1000
MSHFlexGrid2.ColWidth(14) = 0
MSHFlexGrid2.ColWidth(19) = 1200
MSHFlexGrid2.ColWidth(20) = 3500
MSHFlexGrid2.ColWidth(21) = 1200
MSHFlexGrid2.ColWidth(22) = 0
MSHFlexGrid2.ColWidth(23) = 1500
MSHFlexGrid2.RowHeight(1) = 600

End Sub

Private Sub TxtGodownArea_LostFocus()
If TxtGodownArea <> "" Then
   If IsNumeric(TxtGodownArea) = False Then
      MsgBox "Please Provide Proper Godown Area!!!"
      TxtGodownArea.SetFocus
      Exit Sub
   End If
End If
End Sub

Private Sub TxtGodownCapacity_LostFocus()
If TxtGodownCapacity <> "" Then
   If IsNumeric(TxtGodownCapacity) = False Then
      MsgBox "Please Provide Proper Godown Capacity!!!"
      TxtGodownCapacity.SetFocus
      Exit Sub
   End If
End If
End Sub

Private Function ChkGodownCloseDate() As Boolean
Dim Retval As Boolean
Retval = True

For I = 2 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(I, 10) <> "" Then
       If CDate(MSHFlexGrid2.TextMatrix(I, 10)) > CloseDate.Value Then
          Retval = False
          Exit For
       End If
    End If
Next

ChkGodownCloseDate = Retval

End Function
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

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
TOADD = ""
TOCC = ""

mMsg = "<tr><td colspan=3>New Godown Creation Details (From Store-Man)</td></tr>"
mMsg = mMsg & "<tr><td>CMP ID</td><td  colspan=2>" & TxtCMPID.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>CMP Start Date</td><td  colspan=2>" & Format(StartDate.Value, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td>CMP Name</td><td  colspan=2>" & TxtCMPName.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>CMP Display Name</td><td  colspan=2>" & TxtCMPDispalyName.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Address</td><td colspan=2>" & TxtAddress.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Location</td><td colspan=2>" & TxtLocationID.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>State</td><td colspan=2>" & GetStateFromLocation(Val(mLocationID)) & "</td></tr>"
mMsg = mMsg & "<tr><td>Created By</td><td colspan=2>" & TxtCreated.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Updated By</td><td colspan=2>" & TxtUpdated.Text & "</td></tr>"
mMsg = mMsg & "<tr><td colspan=3> </td></tr>"

'mGridMsg = ""
'For I = 1 To MSHFlexGrid2.Rows - 2
'    mGridMsg = mGridMsg & "<tr>"
'    mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 3) & " </td>"
'    mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 4) & " </td>"
'    mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 5) & " </td>"
'    mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 6) & " </td>"
'    mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 7) & " </td>"
'    mGridMsg = mGridMsg & "<td> " & MSHFlexGrid2.TextMatrix(I, 8) & " </td>"
'    mGridMsg = mGridMsg & "<td> " & Format(MSHFlexGrid2.TextMatrix(I, 9), "dd-MMM-yyyy") & " </td>"
'    mGridMsg = mGridMsg & "<td> " & Format(MSHFlexGrid2.TextMatrix(I, 15), "dd-MMM-yyyy") & " </td>"
'    mGridMsg = mGridMsg & "</tr>"
'Next

mMsg = mMsg & mGridMsg
TOADD = "survey.ccg@ncml.com"
TOCC = Gen_UserEmailID

mMsgSubject = "New Godown Creation Details"

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

End Sub


Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  .Server = servertxt
'  .MailFrom = "info@ncml.com" 'Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  .SendTo = ToADD_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With


Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = "info@ncml.com" 'Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

Private Function ChkAllGodownCloseStatus() As Boolean

Retval = True
For I = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 10) = "" Then
       Retval = False
       Exit For
    End If
Next I

ChkAllGodownCloseStatus = Retval

End Function

Private Sub CheckQQSMapping(mGodownID As Variant)
tmp = " Select QQSMapDID,QQSMapID from Txn_QQSGodownMappingDetail where GodownID=" & mGodownID & " "
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
   
       tmp = " Insert into Log_Txn_QQSGodownMappingDetail" & _
             " (ID," & _
             "  QQSMAPID," & _
             "  GodownID," & _
             "  UpdatedBy)" & _
             " ( Select Isnull((Select Max(ID) from Log_Txn_QQSGodownMappingDetail)+1,1) , QQSMAPID,GodownID,'" & Gen_UserLoginID & "'" & _
             " From Txn_QQSGodownMappingDetail Where QQSMapID= " & Val(TmpRs1!QQSMapID) & " And QQSMapDID=" & Val(TmpRs1!QQSMapDID) & ")"

       Call Tmp4Table
   
       tmp = " Delete From Txn_QQSGodownMappingDetail where QQSMapDID=" & TmpRs1!QQSMapDID & ""
       Call Tmp2Table
       tmp = " Select QQSMapID From Txn_QQSGodownMappingDetail Where QQSMapID=" & TmpRs1!QQSMapID & ""
       Call Tmp3Table
       If TmpRs3.RecordCount = 0 Then
          tmp = " Update Txn_QQSGodownMapping Set FreeFlag='Y' Where QQSMAPID=" & TmpRs1!QQSMapID & ""
          Call Tmp4Table
       End If
       TmpRs1.MoveNext
   Next
End If
End Sub
