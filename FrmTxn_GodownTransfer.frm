VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_GodownTransfer 
   Caption         =   "Godown Transfer"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   69
      Top             =   9600
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   68
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9615
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   15015
      Begin VB.Frame Frame0 
         Height          =   1845
         Left            =   720
         TabIndex        =   45
         Top             =   4560
         Visible         =   0   'False
         Width           =   14580
         Begin VB.ComboBox CmbSCMPName 
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   53
            Top             =   1050
            Width           =   3405
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
            Height          =   350
            Left            =   13410
            TabIndex        =   52
            Top             =   765
            Width           =   1000
         End
         Begin VB.CommandButton CmdSearch 
            Caption         =   "Search"
            Height          =   350
            Left            =   11400
            TabIndex        =   51
            Top             =   765
            Width           =   1000
         End
         Begin VB.ComboBox CmbSNewCMPName 
            Height          =   315
            Left            =   6075
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   375
            Width           =   3765
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12405
            TabIndex        =   49
            Top             =   765
            Width           =   1000
         End
         Begin VB.ComboBox CmbSLocationID 
            Height          =   315
            Left            =   2790
            Sorted          =   -1  'True
            TabIndex        =   48
            Top             =   375
            Width           =   3240
         End
         Begin VB.ComboBox CmbSGodownNo 
            Height          =   315
            Left            =   3555
            Sorted          =   -1  'True
            TabIndex        =   47
            Top             =   1020
            Width           =   3045
         End
         Begin VB.ComboBox CmbSNewGodownNo 
            Height          =   315
            Left            =   9900
            Sorted          =   -1  'True
            TabIndex        =   46
            Top             =   360
            Width           =   3765
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   1215
            Left            =   120
            TabIndex        =   54
            Top             =   1560
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2143
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   300
            Left            =   1425
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   375
            Width           =   1320
            _ExtentX        =   2328
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
            Format          =   3932161
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   300
            Left            =   75
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   375
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   3932161
            CurrentDate     =   36494
         End
         Begin VB.Label Label6 
            Caption         =   "Old CMP"
            Height          =   225
            Left            =   1440
            TabIndex        =   63
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
            Height          =   255
            Left            =   360
            TabIndex        =   62
            Top             =   165
            Width           =   855
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
            Height          =   255
            Left            =   1935
            TabIndex        =   61
            Top             =   165
            Width           =   675
         End
         Begin VB.Label Label37 
            Caption         =   "New CMP"
            Height          =   225
            Left            =   7440
            TabIndex        =   60
            Top             =   120
            Width           =   825
         End
         Begin VB.Label Label34 
            Caption         =   "Location"
            Height          =   225
            Left            =   4155
            TabIndex        =   59
            Top             =   135
            Width           =   690
         End
         Begin VB.Label Label9 
            Caption         =   "Old Godown"
            Height          =   225
            Left            =   4710
            TabIndex        =   58
            Top             =   765
            Width           =   1065
         End
         Begin VB.Label Label10 
            Caption         =   "New Godown"
            Height          =   225
            Left            =   11265
            TabIndex        =   57
            Top             =   105
            Width           =   1170
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1170
         Left            =   120
         TabIndex        =   40
         Top             =   8235
         Width           =   14775
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   225
            TabIndex        =   67
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   225
            TabIndex        =   66
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   3165
            TabIndex        =   65
            Top             =   195
            Width           =   2940
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   3165
            TabIndex        =   64
            Top             =   600
            Width           =   2940
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   6105
            TabIndex        =   44
            Top             =   195
            Width           =   2820
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   6105
            TabIndex        =   43
            Top             =   600
            Width           =   2820
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   8925
            TabIndex        =   42
            Top             =   195
            Width           =   2820
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   8925
            TabIndex        =   41
            Top             =   600
            Width           =   2820
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   11745
            TabIndex        =   10
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   11745
            TabIndex        =   9
            Top             =   600
            Width           =   2700
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
         Height          =   2610
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   14775
         Begin VB.CheckBox ChkCloseDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            TabIndex        =   3
            Top             =   2025
            Width           =   210
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   1545
            Width           =   3240
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
            Left            =   1800
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   1125
            Width           =   3240
         End
         Begin VB.ComboBox CmbNewCMPName 
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
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   1125
            Width           =   3720
         End
         Begin VB.ComboBox CmbNewGodownNo 
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
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1545
            Width           =   3720
         End
         Begin VB.TextBox TxtTransferID 
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
            Left            =   10575
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.TextBox TxtRemark 
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
            Height          =   540
            Left            =   6720
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   1980
            Width           =   7920
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
            Left            =   11625
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1545
            Width           =   3000
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
            Left            =   11625
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1125
            Width           =   3000
         End
         Begin MSComCtl2.DTPicker TxtCloseDate 
            Height          =   375
            Left            =   3405
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1965
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   661
            _Version        =   393216
            Format          =   3932161
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtTransferDate 
            Height          =   375
            Left            =   1800
            TabIndex        =   26
            TabStop         =   0   'False
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
            Format          =   3932161
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtCMPName 
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
            TabIndex        =   23
            Top             =   1125
            Visible         =   0   'False
            Width           =   3240
         End
         Begin VB.TextBox TxtGodownNo 
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
            TabIndex        =   24
            Top             =   1545
            Visible         =   0   'False
            Width           =   3240
         End
         Begin VB.TextBox TxtNewCMPName 
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1125
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.TextBox txtNewGodownNo 
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1545
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.ComboBox CmbNewLocationID 
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
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   720
            Width           =   3720
         End
         Begin VB.ComboBox CmbLocationID 
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
            TabIndex        =   0
            Top             =   690
            Width           =   3240
         End
         Begin VB.TextBox TxtLocationID 
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
            TabIndex        =   22
            Top             =   690
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.TextBox TxtNewLocationID 
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   720
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label Label7 
            Caption         =   "New Location"
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
            Left            =   5160
            TabIndex        =   71
            Top             =   750
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "Close Godown"
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
            Left            =   120
            TabIndex        =   39
            Top             =   2010
            Width           =   1560
         End
         Begin VB.Label lblGodownNo 
            Caption         =   "Old Godown No"
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
            TabIndex        =   38
            Top             =   1598
            Width           =   1440
         End
         Begin VB.Label LblCMPName 
            Caption         =   "Old CMP"
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
            Top             =   1155
            Width           =   1800
         End
         Begin VB.Label Label26 
            Caption         =   "New CMP"
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
            Left            =   5175
            TabIndex        =   36
            Top             =   1155
            Width           =   1800
         End
         Begin VB.Label Label21 
            Caption         =   "New Godown"
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
            Left            =   5175
            TabIndex        =   35
            Top             =   1620
            Width           =   1860
         End
         Begin VB.Label LblCloseDate 
            Caption         =   "Close Date"
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
            Left            =   2130
            TabIndex        =   34
            Top             =   1995
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label8 
            Caption         =   "Transfer Date"
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
            TabIndex        =   33
            Top             =   270
            Width           =   1800
         End
         Begin VB.Label Label4 
            Caption         =   "Transfer ID"
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
            Left            =   8880
            TabIndex        =   32
            Top             =   270
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Left            =   5175
            TabIndex        =   30
            Top             =   2010
            Width           =   1440
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
            Left            =   10560
            TabIndex        =   29
            Top             =   1545
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
            Left            =   10560
            TabIndex        =   28
            Top             =   1140
            Width           =   1695
         End
         Begin VB.Label LblLocationID 
            Caption         =   "Old Location"
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
            TabIndex        =   31
            Top             =   720
            Width           =   1440
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Old GSL Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5475
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   7455
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   4845
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   8546
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "New GSL Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5475
         Left            =   7575
         TabIndex        =   12
         Top             =   2760
         Width           =   7335
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   4845
            Left            =   105
            TabIndex        =   13
            Top             =   480
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   8546
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
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label2 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   4650
            TabIndex        =   15
            Top             =   225
            Width           =   390
         End
         Begin VB.Label Label3 
            Caption         =   "Stack No. Already Present"
            Height          =   255
            Left            =   5250
            TabIndex        =   14
            Top             =   225
            Width           =   2025
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_GodownTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mLocationID As Double
Dim mSLocationID As Double
Dim mCMPID, mNewCMPID As Double
Dim mGodownID, mNewGodownID As Double
Dim mSCMPID, mSNewCMPID As Double
Dim mSGodownID, mSNewGodownID As Double
Dim mSM_Prefix, msSM_prefix As String
Dim mNewLocationID As Double
Dim mNewSM_Prefix As String
Dim mOldGodownYearlyFlag, mNewGodownYearlyFlag As Variant

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 9
    .Rows = 3
    .FixedRows = 2
    .Font.Size = 9
    .TextMatrix(1, 0) = "GSL ID"
    .TextMatrix(1, 1) = "Stack No"
    .TextMatrix(1, 2) = "Lot No"
    .TextMatrix(1, 3) = "No. of Bags"
    .TextMatrix(1, 4) = "Quantity"
    .TextMatrix(1, 5) = "Exchange Type"
    .TextMatrix(1, 6) = "Flag"
    .TextMatrix(1, 7) = "Spillage"
    .TextMatrix(1, 8) = "Made Up Flag"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 800
    .ColWidth(2) = 800
    .ColWidth(3) = 1100
    .ColWidth(4) = 1000
    .ColWidth(5) = 1500
    .ColWidth(6) = 700
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    
End With
    

End Sub

Private Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Cols = 9
    .Rows = 3
    .FixedRows = 2
    .Font.Size = 9
    .TextMatrix(1, 0) = "GSL ID"
    .TextMatrix(1, 1) = "Stack No"
    .TextMatrix(1, 2) = "Lot No"
    .TextMatrix(1, 3) = "No. of Bags"
    .TextMatrix(1, 4) = "Quantity"
    .TextMatrix(1, 5) = "Exchange Type"
    .TextMatrix(1, 6) = "Flag"
    .TextMatrix(1, 7) = "Spillage"
    .TextMatrix(1, 8) = "Made up Flag"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 800
    .ColWidth(2) = 800
    .ColWidth(3) = 1100
    .ColWidth(4) = 1000
    .ColWidth(5) = 1500
    .ColWidth(6) = 700
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    
End With
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"

Frame1.Enabled = True
Call GButtonLock(Me, False)
TxtTransferDate.Enabled = False
TxtTransferID.Text = ""
CmbLocationID.Visible = True
CmbLocationID.Text = ""
TxtLocationID.Visible = False

CmbNewLocationID.Visible = True
CmbNewLocationID.Text = ""
TxtNewLocationID.Visible = False

CmbCMPName.Visible = True
CmbCMPName.Text = ""
TxtCMPName.Visible = False
CmbGodownNo.Visible = True
CmbGodownNo.Text = ""
TxtGodownNo.Visible = False
CmbNewCMPName.Visible = True
CmbNewCMPName.Text = ""
TxtNewCMPName.Visible = False
CmbNewGodownNo.Visible = True
CmbNewGodownNo.Text = ""
txtNewGodownNo.Visible = False
TxtRemark.Text = ""
ChkCloseDate.Value = vbUnchecked
LblCloseDate.Visible = False
TxtCloseDate.Visible = False
TxtTransferDate.Value = Date
Call Grid_Head
Call Grid_Head1
If RsTxn_GSL_Transfer.RecordCount > 0 Then
   CmbLocationID.SetFocus
End If
End Sub

Private Sub ChkCloseDate_Click()
If ChkCloseDate.Value = vbChecked Then
   LblCloseDate.Visible = True
   TxtCloseDate.Visible = True
   TxtCloseDate.Value = Date
Else
   LblCloseDate.Visible = False
   TxtCloseDate.Visible = False
End If
End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   If RsMst_Location.RecordCount > 0 Then
      CmbLocationID.SetFocus
   End If
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
   MsgBox "No Record found !!!! "
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
   CmbGodownNo.Text = ""
   mCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodownNo.Text = ""
End If
mCMPID = RsMst_CMP!CMPID

End Sub





Private Sub CmbGodownNo_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
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

Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then
   mGodownID = 0
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
mOldGodownYearlyFlag = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), 0, RsMst_Godown!YearlyBillingFlg)
Call FillOrignalGSLDetail
If CmbNewGodownNo.Text <> "" Then
   Call CheckStackAndLot
End If


End Sub
Private Sub FillOrignalGSLDetail()
Call Grid_Head

tmp = "Select Mst_GSL_ID,StackNo,LotNo,BalanceBags,BalanceWeight,ExchangeType,GSL.Status ,GSL.SpillageFlag,GSL.MadeupFlag"
tmp = tmp & " From Mst_GSL GSL"
tmp = tmp & " left Join Mst_ExchangeType ME on GSL.ExchangeTypeID=ME.ExchangeTypeID"
tmp = tmp & " Where GSL.Status not in ('E','Z') And GodownID=" & mGodownID
tmp = tmp & " order By StackNo,LotNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
   For I = 2 To TmpRs.RecordCount + 1
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       If MSHFlexGrid1.TextMatrix(I, 7) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 7) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 7) = "No"
       End If
       If MSHFlexGrid1.TextMatrix(I, 8) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 8) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 8) = "No"
       End If
       
       MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(I, 3))
       MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
       TmpRs.MoveNext
       
   Next
Else
   MsgBox " GSL Details not found."
End If
End Sub

Private Sub InDataOldGSLDetail(mOldGodownID As Double)

Call Grid_Head

If Format(TxtTransferDate.Value, "mm/dd/yyyy") < "09/26/2013" Then
   tmp = "Select Mst_GSL_ID,StackNo,LotNo,BalanceBags,BalanceWeight,ExchangeType,GSL.Flag ,GSL.SpillageFlag,GSL.MadeUpFlag"
   tmp = tmp & " From Mst_GSL GSL"
   tmp = tmp & " left Join Mst_ExchangeType ME on GSL.ExchangeTypeID=ME.ExchangeTypeID"
   tmp = tmp & " Where GSL.Flag in ('Z') And GodownID=" & mOldGodownID
   tmp = tmp & " order By StackNo,LotNo"
Else
   tmp = " Select GSL.Mst_GSL_ID,StackNo,LotNo,BalanceBags,BalanceWeight,ExchangeType,GSL.Flag ,GSL.SpillageFlag,GSL.MadeUpFlag"
   tmp = tmp & " From Mst_GSL GSL"
   tmp = tmp & " left Join Mst_ExchangeType ME on GSL.ExchangeTypeID=ME.ExchangeTypeID"
   tmp = tmp & " Inner Join Txn_GSL_Transfer TGT On GSL.Mst_GSL_ID = TGT.OldMst_GSL_ID And GSL.SM_Prefix = '" & mSM_Prefix & "'"
   tmp = tmp & " Where GSL.Flag in ('Z') And GSL.GodownID=" & mOldGodownID & "  And TGT.Flag = 'G'"
   tmp = tmp & " order By StackNo,LotNo"
End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
   For I = 2 To TmpRs.RecordCount + 1
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       If MSHFlexGrid1.TextMatrix(I, 7) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 7) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 7) = "No"
       End If
       If MSHFlexGrid1.TextMatrix(I, 8) = "1" Then
          MSHFlexGrid1.TextMatrix(I, 8) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 8) = "No"
       End If
       
       MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(I, 3))
       MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
       
       TmpRs.MoveNext
       
   Next
End If
End Sub
Private Sub InDataNewGSLDetail(mNewGodownID_ As Double)

Call Grid_Head1

If Format(TxtTransferDate.Value, "mm/dd/yyyy") < "09/26/2013" Then
   tmp = "Select Mst_GSL_ID,StackNo,LotNo,BalanceBags,BalanceWeight,ExchangeType,GSL.Flag ,GSL.SpillageFlag,GSL.MadeupFlag"
   tmp = tmp & " From Mst_GSL GSL"
   tmp = tmp & " left Join Mst_ExchangeType ME on GSL.ExchangeTypeID=ME.ExchangeTypeID"
   tmp = tmp & " Where GSL.Flag not in ('E','Z') And GodownID=" & mNewGodownID_ & "" ' And BMst_GSL_ID is not Null"
   tmp = tmp & " order By StackNo,LotNo"
Else

   tmp = " Select GSL.Mst_GSL_ID,StackNo,LotNo,BalanceBags,BalanceWeight,ExchangeType,GSL.Flag ,GSL.SpillageFlag,GSL.MadeUpFlag"
   tmp = tmp & " From Mst_GSL GSL"
   tmp = tmp & " left Join Mst_ExchangeType ME on GSL.ExchangeTypeID=ME.ExchangeTypeID"
   tmp = tmp & " Inner Join Txn_GSL_Transfer TGT On GSL.Mst_GSL_ID = TGT.NewMst_GSL_ID And GSL.SM_Prefix = '" & mNewSM_Prefix & "'"
   tmp = tmp & " Where GSL.Flag not in ('E','Z') And GSL.GodownID=" & mNewGodownID_ & " And TGT.Flag = 'G'"
   tmp = tmp & " order By StackNo,LotNo"

End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 3
   For I = 2 To TmpRs.RecordCount + 1
       
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       If MSHFlexGrid2.TextMatrix(I, 7) = "1" Then
          MSHFlexGrid2.TextMatrix(I, 7) = "Yes"
       Else
          MSHFlexGrid2.TextMatrix(I, 7) = "No"
       End If
       
       If MSHFlexGrid2.TextMatrix(I, 8) = "1" Then
          MSHFlexGrid2.TextMatrix(I, 8) = "Yes"
       Else
          MSHFlexGrid2.TextMatrix(I, 8) = "No"
       End If
       
       MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(I, 3))
       MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(I, 4))
       
       TmpRs.MoveNext
   Next
End If
End Sub
Private Sub CheckStackAndLot()

Call Grid_Head1
tmp = "Select Mst_GSL_ID,StackNo,LotNo,BalanceBags,BalanceWeight,ExchangeType,GSL.Flag ,GSL.SpillageFlag,GSL.MadeupFlag"
tmp = tmp & " From Mst_GSL GSL"
tmp = tmp & " left Join Mst_ExchangeType ME on GSL.ExchangeTypeID=ME.ExchangeTypeID"
tmp = tmp & " Where GSL.Status not in ('E','Z') And GodownID=" & mNewGodownID
tmp = tmp & " order By StackNo,LotNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 3
   For I = 2 To TmpRs.RecordCount + 1
       
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       If MSHFlexGrid2.TextMatrix(I, 7) = "1" Then
          MSHFlexGrid2.TextMatrix(I, 7) = "Yes"
       Else
          MSHFlexGrid2.TextMatrix(I, 7) = "No"
       End If
       If MSHFlexGrid2.TextMatrix(I, 8) = "1" Then
          MSHFlexGrid2.TextMatrix(I, 8) = "Yes"
       Else
          MSHFlexGrid2.TextMatrix(I, 8) = "No"
       End If
       
       MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(I, 3))
       MSHFlexGrid2.TextMatrix(0, 4) = Val(MSHFlexGrid2.TextMatrix(0, 4)) + Val(MSHFlexGrid2.TextMatrix(I, 4))
       
       TmpRs.MoveNext
   Next
End If

  
For I = 2 To MSHFlexGrid1.Rows - 1
    For r = 2 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(r, 0) <> "" Then
           If Trim(MSHFlexGrid2.TextMatrix(r, 1)) = Trim(MSHFlexGrid1.TextMatrix(I, 1)) And Trim(MSHFlexGrid2.TextMatrix(r, 2)) = Trim(MSHFlexGrid1.TextMatrix(I, 2)) Then
              For C = 0 To MSHFlexGrid2.Cols - 1
                  MSHFlexGrid2.row = r
                  MSHFlexGrid2.Col = C
                  MSHFlexGrid2.CellBackColor = vbRed
              Next
           End If
        End If
    Next

Next
End Sub
Private Function DuplicateStack() As Boolean
Dim Retval As Boolean
Retval = False
For r = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(r, 0) <> "" Then
       MSHFlexGrid2.row = r
       MSHFlexGrid2.Col = 0
       If MSHFlexGrid2.CellBackColor = vbRed Then
          Retval = True
          Exit For
       End If
    End If
Next
DuplicateStack = Retval
End Function

Private Sub CmbNewLocationID_GotFocus()

tmp = CmbNewLocationID.Text
If LCase(Gen_UserRole) = "power" Then
   If Gen_WcLocation <> "" Then
      Call TableMst_Location(" Where " & Gen_WcLocation)
   Else
      Call TableMst_Location("")
   End If
Else
   If Trim(CmbLocationID.Text) = "" Then
      MsgBox "Select Old Location first!!!"
      CmbLocationID.SetFocus
      Exit Sub
   End If
   Call TableMst_Location(" where LocationID = " & mLocationID)
End If
CmbNewLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbNewLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbNewLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbNewLocationID.Text = tmp
End Sub

Private Sub CmbNewLocationID_LostFocus()

If CmbNewLocationID.Text = "" Then
   mNewLocationID = 0
   TxtNewPreFix = ""
   mNewSM_Prefix = ""
   Exit Sub
End If
' Code for clearing Godown combox in case of change
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbNewLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbNewLocationID.SetFocus
   Exit Sub
End If
If mNewLocationID <> RsMst_Location!LocationID Then
   CmbNewCMPName.Text = ""
   CmbNewGodownNo.Text = ""
   Call Grid_Head1
End If
mNewLocationID = RsMst_Location!LocationID
mNewSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmdClear_Click()

SFrom.Value = Date - 365
STo.Value = Date
CmbSLocationID.Text = ""
CmbSCMPName.Text = ""
CmbSNewCMPName.Text = ""
CmbSNewGodownNo.Text = ""
CmbSGodownNo.Text = ""

Call Grid_Head2

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_GodownTransfer.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_GodownTransfer.xls")
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
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdSearch_Click()
Dim wc As String
Dim tmp1 As String
Call Grid_Head1

If STo.Value < SFrom.Value Then
   MsgBox "'To date' must be always greater than 'From date'!!!"
   STo.SetFocus
   Exit Sub
End If

tmp = "Select Distinct GSLTransferID,GSLTransferDate,ML.LocationName as NewLocationName,MC.CMPName as NewCMPName,MG.GodownNo as NewGodownNo,ML1.LocationName as OldLocationName,MC1.CMPName as OldCMPName,MG1.GodownNO as OldGodownNo,TG.Remarks"
tmp = tmp & " From Txn_GSL_Transfer TG"
'tmp = tmp & " inner Join Mst_Godown  MG on TG.NewMst_GSL_ID=MG.GodownID"
'tmp = tmp & " inner Join Mst_Godown  MG1 on TG.OldMst_GSL_ID=MG1.GodownID"
tmp = tmp & " inner Join Mst_Godown  MG on TG.NewGodownID=MG.GodownID"
tmp = tmp & " inner Join Mst_Godown  MG1 on TG.OldGodownID=MG1.GodownID"
tmp = tmp & " inner Join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner Join Mst_CMP MC1 on MG1.CMPID=MC1.CMPID"
tmp = tmp & " Left Join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner Join Mst_Location ML1 on MC1.LocationID=ML1.LocationID"
If Gen_DBType = "MDB" Then
   wc = " Where TG.GSLTransferDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = " Where TG.GSLTransferDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If
tmp = tmp & " And TG.Flag='G'"
If CmbSLocationID <> "" Then
   tmp = tmp & " And MC.LocationID=" & mSLocationID
End If
If CmbSNewCMPName <> "" Then
   tmp = tmp & " And MC.CMPID=" & mSNewCMPID
End If

If CmbSNewGodownNo <> "" Then
   tmp = tmp & " And MG.GodownID=" & mSNewGodownID
End If

If CmbSCMPName <> "" Then
   tmp = tmp & " And MC1.CMPID=" & mSCMPID
End If

If CmbSGodownNo <> "" Then
   tmp = tmp & " And MG1.GodownID=" & mSGodownID
End If

Call TmpTable

Call Grid_Head2

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid3.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    TmpRs.MoveNext
Next

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call CmdClear_Click

Call TableTxn_GSL_Transfer("Where Flag='G' Order By GSLTransferID ")


If RsTxn_GSL_Transfer.RecordCount = 0 Then
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
Call Indata

End Sub
Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub
Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location("")
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocationID")
End If
End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   TxtPreFix = ""
   mSM_Prefix = ""
   Exit Sub
End If
' Code for clearing Godown combox in case of change
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   CmbCMPName.Text = ""
   CmbNewCMPName.Text = ""
   CmbGodownNo.Text = ""
   CmbNewGodownNo.Text = ""
   CmbNewLocationID.Text = ""
   Call Grid_Head
   Call Grid_Head1
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
'TxtPreFix = RsMst_Location!SM_Prefix
End Sub



Private Sub CmbNewCMPName_GotFocus()
If CmbNewLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   If RsMst_Location.RecordCount > 0 Then
      CmbNewLocationID.SetFocus
   End If
   Exit Sub
End If
tmp = CmbNewCMPName.Text
If Gen_WcCMP <> "" Then
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mNewLocationID & " And " & Gen_WcCMP)
Else
   Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mNewLocationID)
End If
CmbNewCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbNewCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next

CmbNewCMPName.Text = tmp

End Sub

Private Sub CmbNewCMPName_LostFocus()
If CmbNewCMPName.Text = "" Then
   CmbNewGodownNo.Text = ""
   mNewCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbNewCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbNewCMPName.SetFocus
   Exit Sub
End If
If mNewCMPID <> RsMst_CMP!CMPID Then
   CmbNewGodownNo.Text = ""
End If
mNewCMPID = RsMst_CMP!CMPID
End Sub
Private Sub CmbNewGodownNo_GotFocus()
If CmbNewCMPName.Text = "" Then
   MsgBox "Select New CMP !!!! "
   CmbNewCMPName.SetFocus
   Exit Sub
End If
tmp = CmbNewGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mNewCMPID)
CmbNewGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbNewGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbNewGodownNo.Text = tmp
End Sub

Private Sub CmbNewGodownNo_LostFocus()
If CmbNewGodownNo.Text = "" Then
   mNewGodownID = 0
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbNewGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbNewGodownNo.Enabled = True Then
      CmbNewGodownNo.SetFocus
   End If
   Exit Sub
End If
mNewGodownID = RsMst_Godown!GodownID
mNewGodownYearlyFlag = IIf(IsNull(RsMst_Godown!YearlyBillingFlg), 0, RsMst_Godown!YearlyBillingFlg)
If mNewGodownID = mGodownID Then
   MsgBox "New Godown must be Different from Old Godown."
   CmbNewGodownNo.SetFocus
   Exit Sub
End If
Call CheckStackAndLot
End Sub

Public Sub Indata()

Frame1.Enabled = False
TxtTransferID = RsTxn_GSL_Transfer!GSLTransferID
TxtTransferDate = RsTxn_GSL_Transfer!GSLTransferDate
TxtTransferDate.Enabled = False
TxtRemark = IIf(IsNull(RsTxn_GSL_Transfer!Remarks), "", RsTxn_GSL_Transfer!Remarks)

TxtCreated = IIf(IsNull(RsTxn_GSL_Transfer!CreatedBy), "", RsTxn_GSL_Transfer!CreatedBy) & " :- " & IIf(IsNull(RsTxn_GSL_Transfer!CreatedDate), "", RsTxn_GSL_Transfer!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_GSL_Transfer!UpdatedBy), "", RsTxn_GSL_Transfer!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_GSL_Transfer!UpdatedDate), "", RsTxn_GSL_Transfer!UpdatedDate)


CmbLocationID.Visible = False
TxtLocationID.Visible = True

CmbNewLocationID.Visible = False
TxtNewLocationID.Visible = True


CmbCMPName.Visible = False
TxtCMPName.Visible = True

CmbNewCMPName.Visible = False
TxtNewCMPName.Visible = True


CmbGodownNo.Visible = False
TxtGodownNo.Visible = True

CmbNewGodownNo.Visible = False
txtNewGodownNo.Visible = True

TxtLocationID = ""
TxtNewLocationID = ""
TxtCMPName = ""
TxtGodownNo = ""

TxtNewCMPName = ""
txtNewGodownNo = ""

'TxtGodownNo = GetGodownName(RsTxn_GSL_Transfer!OldMst_GSL_ID, "N")
'txtNewGodownNo = GetGodownName(RsTxn_GSL_Transfer!NewMst_GSL_ID, "N")

TxtGodownNo = GetGodownName(RsTxn_GSL_Transfer!OldGodownID, "N")
txtNewGodownNo = GetGodownName(RsTxn_GSL_Transfer!NewGodownID, "N")

TxtCMPName = GetCMPFromGodown(RsTxn_GSL_Transfer!OldGodownID)
TxtNewCMPName = GetCMPFromGodown(RsTxn_GSL_Transfer!NewGodownID)
TxtLocationID = GetLocationFromCMPName(TxtCMPName)
TxtNewLocationID = GetLocationFromCMPName(TxtNewCMPName)

Call TableMst_Godown(" Where GodownID=" & RsTxn_GSL_Transfer!OldGodownID)
If RsMst_Godown.RecordCount > 0 Then
   If IsNull(RsMst_Godown!CloseDate) = False Then
      ChkCloseDate.Value = vbChecked
      TxtCloseDate.Value = RsMst_Godown!CloseDate
   End If
End If

Call InDataOldGSLDetail(RsTxn_GSL_Transfer!OldGodownID)
Call InDataNewGSLDetail(RsTxn_GSL_Transfer!NewGodownID)

'Call InDataOldGSLDetail(RsTxn_GSL_Transfer!OldMst_GSL_ID)
'Call InDataNewGSLDetail(RsTxn_GSL_Transfer!NewMst_GSL_ID)



If RsTxn_GSL_Transfer.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

DeleteCmd.Enabled = False

EditCmd.Enabled = False

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid3_Click()
Dim I As Variant
I = MSHFlexGrid3.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid3.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_GSL_Transfer.MoveFirst
RsTxn_GSL_Transfer.Find "GSLTransferID = '" & MSHFlexGrid3.TextMatrix(I, 0) & "'"
If Not RsTxn_GSL_Transfer.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub SaveCmd_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Old Location Not Allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCMPName.Text = "" Then
   MsgBox "Blank Old CMP Name Not Allowed!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
If CmbGodownNo.Text = "" Then
   MsgBox "Blank Old Godown Not Allowed!!!"
   CmbGodownNo.SetFocus
   Exit Sub
End If

If CmbNewLocationID.Text = "" Then
   MsgBox "Blank New Location Not Allowed!!!"
   CmbNewLocationID.SetFocus
   Exit Sub
End If

If CmbNewCMPName.Text = "" Then
   MsgBox "Blank New CMP Name Not Allowed!!!"
   CmbNewCMPName.SetFocus
   Exit Sub
End If
If CmbNewGodownNo.Text = "" Then
   MsgBox "Blank New Godown Not Allowed!!!"
   CmbNewGodownNo.SetFocus
   Exit Sub
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark Not Allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If MSHFlexGrid1.Rows <= 2 Then
   MsgBox "Blank GSL Details not Allowed."
   TxtGodownNo.SetFocus
   Exit Sub
End If


If mNewGodownYearlyFlag <> mOldGodownYearlyFlag Then
   MsgBox "Yearly billing flag should be same for GSL Transfer!!!"
   Exit Sub
End If

If DuplicateStack = True Then
   MsgBox " Duplicate Stack No, You can't Transfer this Godown."
   CmbNewGodownNo.SetFocus
   Exit Sub
End If

If mGodownID <> mNewGodownID Then
   If CheckMappedReservationContract = False Then
      MsgBox "Transferee godown  is mapped to Reservation Contract(s). Kindly inform CMG team to map  the same to new godown before saving this Transfer entry!!!"
      Exit Sub
   End If
End If




'If Edit_Flg = "A" Then
'   RsTxn_GSL_Transfer.AddNew
'   TxtTransferID = GetGSLTransferID
'   RsTxn_GSL_Transfer!G_UID = GetGUID
'End If
'
'RsTxn_GSL_Transfer!GSLTransferID = TxtTransferID 'GetGSLTransferID
'RsTxn_GSL_Transfer!GSLTransferDate = TxtTransferDate.Value
'RsTxn_GSL_Transfer!OldMst_GSL_ID = mGodownID
'RsTxn_GSL_Transfer!NewMst_GSL_ID = mNewGodownID
'RsTxn_GSL_Transfer!SM_Prefix = mSM_Prefix
'RsTxn_GSL_Transfer!NewSM_Prefix = mNewSM_Prefix
'RsTxn_GSL_Transfer!Flag = "G" 'TxtFlag.Text
'RsTxn_GSL_Transfer!Remarks = TxtRemark.Text
'
'If IsNull(RsTxn_GSL_Transfer!CreatedBy) = True Or RsTxn_GSL_Transfer!CreatedBy = "" Then
'   RsTxn_GSL_Transfer!CreatedBy = Gen_UserLoginID
'   RsTxn_GSL_Transfer!CreatedDate = Now
'Else
'   RsTxn_GSL_Transfer!UpdatedBy = Gen_UserLoginID
'   RsTxn_GSL_Transfer!UpdatedDate = Now
'End If
'
'
'RsTxn_GSL_Transfer.Update

Call SaveGSL

If ChkCloseDate.Value = vbChecked Then
   Call TableMst_Godown(" Where GodownID=" & mGodownID)
   If TxtCloseDate.Visible = True Then
      RsMst_Godown!CloseDate = TxtCloseDate.Value
      RsMst_Godown.Update
   End If
End If
If RsTxn_GSL_Transfer.RecordCount > 0 Then
   RsTxn_GSL_Transfer.MoveLast
End If

Call Indata

End Sub
Private Sub SaveGSL()
Dim mFMst_GSL_ID As Double
Call TableMst_GSL("Where SM_Prefix='" & mSM_Prefix & "'") '("Where SM_Prefix='" & mSM_Prefix & "'")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       'Call TableMst_GSL("Where SM_Prefix='" & mSM_Prefix & "'") '("Where SM_Prefix='" & mSM_Prefix & "'")
       If RsMst_GSL.RecordCount > 0 Then
          RsMst_GSL.MoveFirst
       End If
       RsMst_GSL.Find ("Mst_GSL_ID= " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & "")
       If Not RsMst_GSL.EOF Then
         
          tmp = "Select * from Mst_GSL Where Mst_GSL_ID = " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & " And SM_Prefix='" & mSM_Prefix & "'"
          Call Tmp3Table
          If TmpRs3.RecordCount > 0 Then
             RsMst_GSL.AddNew
             'RsMst_GSL!Mst_GSL_ID = GetGSLID
             RsMst_GSL!G_UID = GetGUID
             mFGSLID = RsMst_GSL!GSLID
             RsMst_GSL!SM_Prefix = mNewSM_Prefix 'mSM_Prefix
             RsMst_GSL!CMPID = mNewCMPID
             RsMst_GSL!GodownID = mNewGodownID
             RsMst_GSL!CommodityID = TmpRs3!CommodityID
             RsMst_GSL!StackNo = TmpRs3!StackNo
             RsMst_GSL!LotNo = TmpRs3!LotNo
             RsMst_GSL!DepositeWeight = TmpRs3!DepositeWeight
             RsMst_GSL!DepositeBags = TmpRs3!DepositeBags
             RsMst_GSL!WHRBags = TmpRs3!WHRBags
             RsMst_GSL!CDTFBags = TmpRs3!CDTFBags
             RsMst_GSL!BalanceBags = TmpRs3!BalanceBags
             mBalanceBags = RsMst_GSL!BalanceBags
             
             RsMst_GSL!CDTFWeight = TmpRs3!CDTFWeight
             RsMst_GSL!BalanceWeight = TmpRs3!BalanceWeight
             mBalanceWeight = RsMst_GSL!BalanceWeight
             
             RsMst_GSL!ExchangeTypeID = TmpRs3!ExchangeTypeID
             RsMst_GSL!ClientIDRow = TmpRs3!ClientIDRow
             RsMst_GSL!PledgeNo = TmpRs3!PledgeNo
             
             RsMst_GSL!Status = TmpRs3!Status
             RsMst_GSL!State = TmpRs3!State
            
             RsMst_GSL!WHRWeight = TmpRs3!WHRWeight
             
             RsMst_GSL!ISINID = TmpRs3!ISINID
             RsMst_GSL!CdtfNo = TmpRs3!CdtfNo
             RsMst_GSL!CDTFDate = TmpRs3!CDTFDate
             
             RsMst_GSL!FirstISINID = TmpRs3!FirstISINID
             RsMst_GSL!PledgeBags = TmpRs3!PledgeBags
             RsMst_GSL!PledgeWeight = TmpRs3!PledgeWeight
             RsMst_GSL!SpillageFlag = TmpRs3!SpillageFlag
             
             RsMst_GSL!DepositAdjWt = TmpRs3!DepositAdjWt
             RsMst_GSL!WithdrawalAdjWt = TmpRs3!WithdrawalAdjWt
             RsMst_GSL!Flag1 = TmpRs3!Flag1
             RsMst_GSL!LCDTF_Date = TmpRs3!LCDTF_Date
             RsMst_GSL!GrossWt = TmpRs3!GrossWt
             RsMst_GSL!KRNo = TmpRs3!KRNo
             RsMst_GSL!KRNoBatchID = TmpRs3!KRNoBatchID
             RsMst_GSL!KRDate = TmpRs3!KRDate
             RsMst_GSL!BGSLID = TmpRs3!Mst_GSL_ID
             mBGSLID = RsMst_GSL!BGSLID
             RsMst_GSL!BSM_Prefix = TmpRs3!SM_Prefix
             RsMst_GSL!ContractID = TmpRs3!ContractID
             RsMst_GSL!BagSize = TmpRs3!BagSize
             RsMst_GSL!BagEquivalent = TmpRs3!BagEquivalent
             RsMst_GSL!WtQty = TmpRs3!WtQty
             RsMst_GSL!CRAmount = TmpRs3!CRAmount
             RsMst_GSL!CRDate = TmpRs3!CRDate
             RsMst_GSL!NContractID = TmpRs3!NContractID
             RsMst_GSL!SCMFlag = TmpRs3!SCMFlag
             RsMst_GSL!MadeupFlag = TmpRs3!MadeupFlag
             RsMst_GSL!CMSPBatchID = TmpRs3!CMSPBatchID
             RsMst_GSL!CMSPBatchGeneratedBy = TmpRs3!CMSPBatchGeneratedBy
             RsMst_GSL!CMSPBatchGeneratedDate = TmpRs3!CMSPBatchGeneratedDate
             RsMst_GSL!CMSPUpdatedBy = TmpRs3!CMSPUpdatedBy
             RsMst_GSL!CMSPUpdatedDate = TmpRs3!CMSPUpdatedDate
             
             RsMst_GSL!RateCardType = TmpRs3!RateCardType
             RsMst_GSL!BillingCycleID = TmpRs3!BillingCycleID
             RsMst_GSL!BillingUnit = TmpRs3!BillingUnit
             RsMst_GSL!Rate = TmpRs3!Rate
             RsMst_GSL!RateForDailyPnL = TmpRs3!RateForDailyPnL
             RsMst_GSL!DepConDID = TmpRs3!DepConDID
             RsMst_GSL!NoOfDays = TmpRs3!NoOfDays
             
             
             If IsNull(RsMst_GSL!CreatedBy) = True Or RsMst_GSL!CreatedBy = "" Then
                RsMst_GSL!CreatedBy = Gen_UserLoginID
                RsMst_GSL!CreatedDate = Now
             Else
                RsMst_GSL!UpdatedBy = Gen_UserLoginID
                RsMst_GSL!UpdatedDate = Now
             End If
            
             TmpRs3!CDTFBags = TmpRs3!DepositeBags
             TmpRs3!CDTFWeight = TmpRs3!DepositeWeight
             TmpRs3!BalanceBags = 0
             TmpRs3!BalanceWeight = 0
             TmpRs3!FGSLID = mFGSLID
             TmpRs3!FSM_Prefix = mNewSM_Prefix
             TmpRs3!Flag = "Z"
             RsMst_GSL.Update
             RsMst_GSL.Requery
             TmpRs3.Update
             TmpRs3.Requery
             
             tmp = "Update Txn_RRNPWRGSL Set NewTransferSM_Prefix = '" & mNewSM_Prefix & "',NewTransferGSLID = " & RsMst_GSL!GSLID & " Where SM_Prefix = '" & mSM_Prefix & "' And (NewGSLID = " & TmpRs3!GSLID & " Or NewTransferGSLID = " & TmpRs3!GSLID & ")"
             Call Tmp3Table
             
             RsTxn_GSL_Transfer.AddNew
             RsTxn_GSL_Transfer!GSLTransferID = GetGSLTransferID
             RsTxn_GSL_Transfer!G_UID = GetGUID
             RsTxn_GSL_Transfer!GSLTransferDate = TxtTransferDate.Value
             RsTxn_GSL_Transfer!OldMst_GSL_ID = mBGSLID
             RsTxn_GSL_Transfer!NewMst_GSL_ID = mFGSLID
             RsTxn_GSL_Transfer!OldGodownID = mGodownID
             RsTxn_GSL_Transfer!NewGodownID = mNewGodownID
             RsTxn_GSL_Transfer!SM_Prefix = mSM_Prefix
             RsTxn_GSL_Transfer!NewSM_Prefix = mNewSM_Prefix
             RsTxn_GSL_Transfer!TransferedBags = mBalanceBags
             RsTxn_GSL_Transfer!TransferedWeight = mBalanceWeight
             RsTxn_GSL_Transfer!Flag = "G" 'TxtFlag.Text
             RsTxn_GSL_Transfer!Remarks = TxtRemark.Text
             RsTxn_GSL_Transfer!CreatedBy = Gen_UserLoginID
             RsTxn_GSL_Transfer!CreatedDate = Now
             RsTxn_GSL_Transfer.Update
             RsTxn_GSL_Transfer.Requery
             
          End If
       End If
    End If
Next

tmp = " select gsl2.Mst_GSL_ID gsl2id , mgsl.Mst_GSL_ID, mgsl.connectedMst_GSL_ID, mgsl.bMst_GSL_ID, mgsl.fMst_GSL_ID, mgsl.madeupflag "
tmp = tmp & " from Mst_GSL mgsl"
tmp = tmp & " inner join mst_gsl gsl1 on gsl1.Mst_GSL_ID = mgsl.bMst_GSL_ID and gsl1.sm_prefix = '" & mSM_Prefix & "' "
tmp = tmp & " inner join mst_gsl gsl2 on gsl1.connectedMst_GSL_ID = gsl2.bMst_GSL_ID and gsl2.sm_prefix = '" & mNewSM_Prefix & "' "
tmp = tmp & " Where mgsl.GodownID = " & mNewGodownID & " And mgsl.madeupflag = 1 And mgsl.connectedMst_GSL_ID Is Null"
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   For k = 1 To TmpRs4.RecordCount
      tmp = "Update Mst_GSL set ConnectedMst_GSL_ID = " & TmpRs4!GSL2ID & " where Mst_GSL_ID = " & TmpRs4!Mst_GSL_ID & " And GodownID = " & mNewGodownID
      Call Tmp5Table
      tmp = "Update Mst_GSL set ConnectedMst_GSL_ID = " & TmpRs4!Mst_GSL_ID & " where Mst_GSL_ID = " & TmpRs4!GSL2ID & " And GodownID = " & mNewGodownID
      Call Tmp5Table
      TmpRs4.MoveNext
   Next
End If

End Sub
Private Function GetGSLID() As Double
Dim Retval As Double
tmp = "Select max(GSLID) from Mst_GSL where SM_Prefix = '" & mNewSM_Prefix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetGSLID = Retval
End Function

Private Function GetGSLTransferID() As Double
Dim Retval As Double
tmp = "Select Max(GSLTransferID) from Txn_GSL_Transfer"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
TmpRs.Close
GetGSLTransferID = Retval
End Function
Private Sub CmbSLocationID_GotFocus()

tmp = CmbSLocationID.Text
Call TableMst_Location("")
CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp

End Sub


Private Sub CmbSLocationID_LostFocus()

If CmbSLocationID.Text = "" Then
   mSLocationID = 0
   msSM_prefix = ""
   Exit Sub
End If
' Code for clearing Godown combox in case of change
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location  selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
If mSLocationID <> RsMst_Location!LocationID Then
   CmbSCMPName.Text = ""
   CmbSNewCMPName.Text = ""
   CmbSGodownNo.Text = ""
   CmbSNewGodownNo.Text = ""
End If
mSLocationID = RsMst_Location!LocationID
msSM_prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmbSNewCMPName_GotFocus()
If CmbSLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   If RsMst_Location.RecordCount > 0 Then
      CmbSLocationID.SetFocus
   End If
   Exit Sub
End If
tmp = CmbSNewCMPName.Text
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mSLocationID)
CmbSNewCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSNewCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSNewCMPName.Text = tmp

End Sub

Private Sub CmbSNewCMPName_LostFocus()
If CmbSNewCMPName.Text = "" Then
   CmbSNewGodownNo.Text = ""
   mSNewCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSNewCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSNewCMPName.SetFocus
   Exit Sub
End If
If mSNewCMPID <> RsMst_CMP!CMPID Then
   CmbSNewGodownNo.Text = ""
End If
mSNewCMPID = RsMst_CMP!CMPID

End Sub
Private Sub CmbSNewGodownNo_GotFocus()
If CmbSNewCMPName.Text = "" Then
   MsgBox "Select New CMP !!!! "
   CmbSNewCMPName.SetFocus
   Exit Sub
End If
tmp = CmbSNewGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mSNewCMPID)
CmbSNewGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSNewGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSNewGodownNo.Text = tmp
End Sub

Private Sub CmbSNewGodownNo_LostFocus()
If CmbSNewGodownNo.Text = "" Then
   mSNewGodownID = 0
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSNewGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbSNewGodownNo.Enabled = True Then
      CmbSNewGodownNo.SetFocus
   End If
   Exit Sub
End If
mSNewGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSCMPName_GotFocus()
If CmbSLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   If RsMst_Location.RecordCount > 0 Then
      CmbSLocationID.SetFocus
   End If
   Exit Sub
End If
tmp = CmbSCMPName.Text
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mSLocationID)
CmbSCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbSGodownNo.Text = ""
   mSCMPID = 0
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCMPName.SetFocus
   Exit Sub
End If
If mSCMPID <> RsMst_CMP!CMPID Then
   CmbSGodownNo.Text = ""
End If
mSCMPID = RsMst_CMP!CMPID

End Sub

Private Sub CmbSGodownNo_GotFocus()
If CmbSCMPName.Text = "" Then
   MsgBox "Select old CMP !!!! "
   CmbSCMPName.SetFocus
   Exit Sub
End If
tmp = CmbSGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mSCMPID)
CmbSGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodownNo.Text = tmp
End Sub

Private Sub CmbSGodownNo_LostFocus()
If CmbSGodownNo.Text = "" Then
   mSGodownID = 0
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodownNo.SetFocus
   Exit Sub
End If
mSGodownID = RsMst_Godown!GodownID

End Sub

Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Cols = 9
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.FixedRows = 1

MSHFlexGrid3.TextMatrix(0, 0) = "Transfer ID"
MSHFlexGrid3.TextMatrix(0, 1) = "Transfer Date"
MSHFlexGrid3.TextMatrix(0, 2) = "New Location Name"
MSHFlexGrid3.TextMatrix(0, 3) = "New CMP Name"
MSHFlexGrid3.TextMatrix(0, 4) = "New Godown No"
MSHFlexGrid3.TextMatrix(0, 5) = "New Location Name"
MSHFlexGrid3.TextMatrix(0, 6) = "Old CMP Name"
MSHFlexGrid3.TextMatrix(0, 7) = "Old Godwn No"
MSHFlexGrid3.TextMatrix(0, 8) = "Remarks"

MSHFlexGrid3.ColWidth(0) = 1000
MSHFlexGrid3.ColWidth(1) = 1400
MSHFlexGrid3.ColWidth(2) = 2200
MSHFlexGrid3.ColWidth(3) = 2000
MSHFlexGrid3.ColWidth(4) = 1400
MSHFlexGrid3.ColWidth(5) = 2000
MSHFlexGrid3.ColWidth(6) = 1400
MSHFlexGrid3.ColWidth(7) = 2200

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_GSL_Transfer.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

'Frame0.Height = Me.Height - 500
'Frame0.Left = Frame1.Left
'Frame0.Top = Frame1.Top
'Frame0.Width = Me.Width - 300
'Frame0.Visible = True
'MSHFlexGrid3.Width = Frame0.Width - 200
'MSHFlexGrid3.Height = Frame0.Height - 1500


'Frame0.Height = Me.Height - 800
'Frame0.Left = Frame4.Left
'Frame0.Top = Frame4.Top
'Frame0.Width = Me.Width - 500
'Frame0.Visible = True
'MSHFlexGrid3.Width = Frame0.Width - 500
'MSHFlexGrid3.Height = Frame0.Height - 2000

Frame0.Height = Frame4.Height - 200 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 200 'Me.Width - 200
Frame0.Visible = True

MSHFlexGrid3.Width = Frame0.Width - 200
MSHFlexGrid3.Height = Frame0.Height - 1400

''===========comment on 7 july 2012 for not clearing search grid=============
'Call CmdClear_Click
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub LastCmd_Click()
RsTxn_GSL_Transfer.MoveLast
Call Indata
'TmpRs.MoveLast
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_GSL_Transfer.MoveNext
If RsTxn_GSL_Transfer.EOF Then
   RsTxn_GSL_Transfer.MoveLast
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
RsTxn_GSL_Transfer.MovePrevious

If RsTxn_GSL_Transfer.BOF Then
   RsTxn_GSL_Transfer.MoveFirst
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
RsTxn_GSL_Transfer.MoveFirst
Call Indata
'TmpRs.MoveLast
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
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

Private Function CheckMappedReservationContract() As Boolean
Retval = True

tmp = "Select Distinct ContractID from Txn_SRGodownMapDetails a "
tmp = tmp & " inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID "
tmp = tmp & " Where a.GodownID = " & mGodownID & " And b.ContractID in (Select Distinct NContractID from Mst_GSL Where Flag in ('O') And GodownID = " & mGodownID & ")"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For x = 1 To TmpRs.RecordCount
       tmp = "Select Distinct a.GodownID from Txn_SRGodownMapDetails a "
       tmp = tmp & " inner Join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID "
       tmp = tmp & " Where b.ContractID = " & TmpRs!ContractID & " And a.GodownID = " & mNewGodownID & " "
       
       Call Tmp1Table
       
       If TmpRs1.RecordCount = 0 Then
          Retval = False
          Exit For
       End If
       TmpRs.MoveNext
   Next
End If

CheckMappedReservationContract = Retval

End Function

