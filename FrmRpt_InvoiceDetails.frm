VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_InvoiceDetails 
   Caption         =   "Invoice Details Report"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox ChkBatchNo 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10785
         TabIndex        =   45
         Top             =   1725
         Value           =   1  'Checked
         Width           =   225
      End
      Begin VB.CheckBox Check2 
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
         TabIndex        =   11
         Top             =   1750
         Value           =   1  'Checked
         Width           =   225
      End
      Begin VB.TextBox TxtNAVCode 
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
         Left            =   12870
         TabIndex        =   10
         Top             =   1245
         Width           =   2010
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show ClientName"
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
         Left            =   6435
         TabIndex        =   15
         Top             =   1750
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.ComboBox CmbStateID 
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
         TabIndex        =   1
         Top             =   420
         Width           =   3765
      End
      Begin VB.TextBox TxtSearchBy 
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
         Left            =   2160
         TabIndex        =   14
         Top             =   2100
         Width           =   3825
      End
      Begin VB.ComboBox CmbSearchBy 
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
         ItemData        =   "FrmRpt_InvoiceDetails.frx":0000
         Left            =   9600
         List            =   "FrmRpt_InvoiceDetails.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1245
         Width           =   3240
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
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
         Left            =   13020
         TabIndex        =   32
         Top             =   420
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
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
         Left            =   14220
         TabIndex        =   31
         Top             =   420
         Width           =   615
      End
      Begin VB.ComboBox CmbDPName 
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
         TabIndex        =   6
         Top             =   885
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CmbCNonNcdexClient 
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
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   885
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CmbClientDP 
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
         Left            =   5160
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1245
         Width           =   4335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show DP "
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
         Left            =   11640
         TabIndex        =   4
         Top             =   480
         Width           =   1140
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
         Height          =   375
         Left            =   13590
         TabIndex        =   18
         Top             =   2055
         Width           =   1335
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
         Height          =   375
         Left            =   12255
         TabIndex        =   17
         Top             =   2055
         Width           =   1335
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   10920
         TabIndex        =   16
         Top             =   2055
         Width           =   1335
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         Left            =   8250
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   3240
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
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   4245
      End
      Begin VB.TextBox TxtCDPID 
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1215
      End
      Begin VB.TextBox TxtCDPName 
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1245
         Width           =   3135
      End
      Begin VB.TextBox TxtCClientName 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1245
         Width           =   3495
      End
      Begin VB.TextBox TxtClientID 
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
         Left            =   120
         TabIndex        =   5
         Top             =   1245
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   345
         Left            =   1800
         TabIndex        =   12
         Top             =   1705
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   111869953
         CurrentDate     =   40012
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   345
         Left            =   4440
         TabIndex        =   13
         Top             =   1705
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   111869953
         CurrentDate     =   40012
      End
      Begin VB.Label Label30 
         Caption         =   "Approved By FAG"
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
         Left            =   8880
         TabIndex        =   46
         Top             =   1725
         Width           =   1710
      End
      Begin VB.Label lblToDate 
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
         Left            =   3480
         TabIndex        =   44
         Top             =   1757
         Width           =   855
      End
      Begin VB.Label lblFromDate 
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
         Height          =   225
         Left            =   510
         TabIndex        =   43
         Top             =   1765
         Width           =   1020
      End
      Begin VB.Label Label4 
         Caption         =   "NAV Code"
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
         Left            =   13365
         TabIndex        =   42
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   1470
         TabIndex        =   40
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label lblSearchBy 
         Caption         =   "TI No (Voucher No)"
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
         TabIndex        =   39
         Top             =   2160
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Search By"
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
         Left            =   10718
         TabIndex        =   38
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label28 
         Caption         =   "Client ID"
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
         Left            =   420
         TabIndex        =   30
         Top             =   915
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Client Name "
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
         Left            =   3120
         TabIndex        =   29
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "DP ID"
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
         Left            =   5535
         TabIndex        =   28
         Top             =   915
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "DP Name "
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
         TabIndex        =   27
         Top             =   915
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "ClientDPID"
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
         Left            =   6360
         TabIndex        =   26
         Top             =   915
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   8835
         TabIndex        =   22
         Top             =   2115
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record Found :-"
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
         Left            =   6420
         TabIndex        =   21
         Top             =   2115
         Width           =   2280
      End
      Begin VB.Label LblExchangeType 
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
         Height          =   255
         Left            =   9120
         TabIndex        =   20
         Top             =   135
         Width           =   1920
      End
      Begin VB.Label Label8 
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
         Left            =   5662
         TabIndex        =   19
         Top             =   135
         Width           =   840
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   33
      Top             =   2640
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Invoice Details"
      TabPicture(0)   =   "FrmRpt_InvoiceDetails.frx":0033
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Clients having same NAV code"
      TabPicture(1)   =   "FrmRpt_InvoiceDetails.frx":004F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   6255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   11033
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
      End
      Begin VB.Frame Frame2 
         Height          =   6615
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   14790
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   6360
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6135
            Left            =   120
            TabIndex        =   35
            Top             =   150
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   10821
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
      End
   End
End
Attribute VB_Name = "FrmRpt_InvoiceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityID, mLocationID As Variant
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mClientNAVCode As String
Dim mClientRNAVCode As String
Dim mDPNAVCode As String
Dim mStateID As Variant
Dim mDPID As Variant

''-- Ullhas Patil 27 Sep 2017
Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        TxtFromDate.Enabled = True: TxtFromDate.Value = Date
        TxtToDate.Enabled = True: TxtToDate.Value = Date
    Else
        TxtFromDate.Enabled = False: TxtFromDate.Value = Date
        TxtToDate.Enabled = False: TxtToDate.Value = Date
    End If
End Sub
''-----

Private Sub CmbLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Select State!!!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID=" & mStateID & " ")
Else
   Call TableMst_Location("Where StateID =" & mStateID & " And " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!Locationname
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

End Sub

Private Sub CmbSearchBy_LostFocus()

If CmbSearchBy.Text = "" Then
   lblSearchBy.Caption = "TI No (Voucher No)"
   TxtSearchBy = ""
   TxtSearchBy.Locked = True
   Exit Sub
End If

If LCase(CmbSearchBy.Text) <> "ti no" And LCase(CmbSearchBy.Text) <> "post invoice no" And LCase(CmbSearchBy.Text) <> "entry no" Then
   MsgBox "Invalid Selection!!!"
   CmbSearchBy.SetFocus
   Exit Sub
End If

lblSearchBy.Caption = CmbSearchBy.Text
TxtSearchBy.Locked = False

End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where " & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
Label2.Caption = ""
CmbLocation.Text = ""
Call Grid_Head
Call Grid_Head1
Label2.Caption = ""
TxtClientID = ""
TxtCClientName = ""
TxtCDPID = ""
TxtCDPName = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
CmbExchangeTypeID.Text = ""
mExchangeTypeID = 0
CmbCNonNcdexClient.Text = ""
Check3.Value = 0
CmbDPName.Text = ""
CmdExcel.Enabled = False
Option1.Value = True
SSTab1.Tab = 0
CmbSearchBy.Text = ""
TxtSearchBy = ""
TxtSearchBy.Locked = True
CmbCNonNcdexClient.Width = 9400
CmbCNonNcdexClient.Top = TxtClientID.Top
CmbCNonNcdexClient.Left = TxtClientID.Left
CmbCNonNcdexClient.Visible = True
Call Check3_Click
Check1.Visible = False
Check1.Value = 1
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
   Check1.Visible = True
   Check1.Value = 1
End If
ProgressBar1.Value = 0

End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
Call Grid_Head1
ProgressBar1.Value = 0
MsgBox "Wait till next message!!!"

'If CmbExchangeTypeID.Text = "" And CmbLocation.Text = "" Then
'   MsgBox "Select either Location or Client!!"
'   Exit Sub
'End If

If CmbExchangeTypeID.Text <> "" Then
   If Check3.Value = 0 Then
'      If CmbClientDP.Visible = True Then
'         MsgBox "Invalid selection !!! "
'         Exit Sub
'      End If
      If CmbCNonNcdexClient.Visible = True Then
         If CmbCNonNcdexClient.Text = "" And TxtNAVCode = "" Then
            MsgBox "Blank Client Name not allowed!!!" '"Invalid selection !!! "
            CmbCNonNcdexClient.SetFocus
            Exit Sub
         End If
      End If
'      If TxtClientID = "" Then
'         MsgBox "Invalid Client ID !!! "
'         Exit Sub
'      End If
   Else
      If CmbDPName.Visible = True Then
         If CmbDPName.Text = "" Then
            MsgBox "Blank DP Name Not allowed !!! "
            CmbDPName.SetFocus
            Exit Sub
         End If
      End If
   End If
End If

''-- Ullhas Patil 27 Sep 2017
If Check2.Value = vbChecked Then
    If CDate(TxtFromDate.Value) > CDate(TxtToDate.Value) Then
        MsgBox "'From Date' can not bigger then 'To Date'."
        Exit Sub
    End If
End If
''----------------------------

''''tmp = " Select TINo,TIDate,TIAmount,(select Sum(RecAmount) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecAmount,"
''''tmp = tmp & " (select Sum(RecTDS) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecTDSAmount"
''''tmp = tmp & " ,(select Sum(RecOtherDed) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecOtherDed"
''''tmp = tmp & " ,TIID,TI.NAVCode,TI.PostInvoiceNo,TI.EntryNo,ML.LocationName,MS.StateName From  Txn_TIDetails TI "

If ChkBatchNo.Value = vbChecked Then
    tmp = " Select TINo,TIDate,TIAmount,"
    tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0 )+ isnull(TI.AdjAmountFAG,0) as RecAmount,"
    tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.BatchID > 0 )RecTDSAmount,"
    tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.BatchID > 0 )RecOtherDed"
    tmp = tmp & " ,TIID,TI.NAVCode,TI.PostInvoiceNo,TI.EntryNo,ML.LocationName,MS.StateName,TI.CollectionRemark,TI.CRemarkUpdatedBy,TI.CRemarkUpdatedDate From  Txn_TIDetails TI "
    tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID "
    tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID "
Else
    tmp = " Select TINo,TIDate,TIAmount,"
    tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID )+ isnull(TI.AdjAmountFAG,0) as RecAmount,"
    tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
    tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecOtherDed"
    tmp = tmp & " ,TIID,TI.NAVCode,TI.PostInvoiceNo,TI.EntryNo,ML.LocationName,MS.StateName,TI.CollectionRemark,TI.CRemarkUpdatedBy,TI.CRemarkUpdatedDate From  Txn_TIDetails TI "
    tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID "
    tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID "
End If

tmp = tmp & " Where  TINo<>'' "
If Option1.Value = True Then
   
   ''tmp = tmp & " And TIAmount-(select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID)>0 "
   If ChkBatchNo.Value = vbChecked Then
      tmp = tmp & " And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y' And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+"
      tmp = tmp & " (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R' And TCR.BatchID > 0 )+AdjAmountFAG)>0"
   Else
      tmp = tmp & " And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y' )+"
      tmp = tmp & " (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R' )+AdjAmountFAG)>0"
   End If
End If

If CmbStateID.Text <> "" Then
   If CmbExchangeTypeID.Text = "" Then
      tmp = tmp & " And (MS.StateID=" & mStateID & " ) "
   Else
      tmp = tmp & " And (MS.StateID=" & mStateID & " Or MS.StateID is Null) "
   End If
Else
   If Gen_WcState <> "" Then
      If CmbStateID.Text = "" Then
         tmp = tmp & " And (MS." & Gen_WcState & " )"
      Else
         tmp = tmp & " And (MS." & Gen_WcState & " Or MS.StateID is Null) "
      End If
   End If
End If
If CmbLocation.Text <> "" Then
   If CmbExchangeTypeID.Text = "" Then
      tmp = tmp & " And (ML.LocationID=" & mLocationID & " ) "
   Else
      tmp = tmp & " And (ML.LocationID=" & mLocationID & " Or ML.LocationID is Null) "
   End If
Else
   If Gen_WcLocation <> "" Then
      If CmbExchangeTypeID.Text = "" Then
         tmp = tmp & " And (ML." & Gen_WcLocation & " )"
      Else
         tmp = tmp & " And (ML." & Gen_WcLocation & " Or ML.LocationID is Null) "
      End If
   End If
End If



If CmbExchangeTypeID.Text <> "" Then
    If Check3.Value = 0 Then
       If mClientRNAVCode <> "" Then
          'tmp = tmp & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
          tmp = tmp & " And (TI.NAVCode in (Select distinct GSTNavCode From Mst_ClientStateWiseGST  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct RGSTNavCode From Mst_ClientStateWiseGST  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct NAVCode From Mst_Client  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct RNAVCode From Mst_Client  Where ClientID = " & mClientIDRow & " ))"
       Else
          'tmp = tmp & " And TI.NAVCode = '" & mClientNAVCode & "'"
          tmp = tmp & " And (TI.NAVCode in (Select distinct GSTNavCode From Mst_ClientStateWiseGST  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct NAVCode From Mst_Client  Where ClientID = " & mClientIDRow & " ))"
       End If
    Else
       'tmp = tmp & " And TI.NAVCode='" & mDPNAVCode & "'"
       tmp = tmp & " And (TI.NAVCode in (Select distinct GSTNavCode From Mst_DepositoryStateWiseGST  Where DepositoryID = '" & mDPID & "' ) Or TI.NAVCode in (Select distinct NAVCode From Mst_Depository  Where DepositoryID = '" & mDPID & "' ))"
    End If
End If

If CmbSearchBy.Text <> "" Then
   If Left(LCase(CmbSearchBy.Text), 1) = "t" Then
      tmp = tmp & " And TINo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "p" Then
      tmp = tmp & " And PostInvoiceNo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "e" Then
      tmp = tmp & " And EntryNo like '%" & TxtSearchBy & "%'"
   End If
Else
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
      If CmbExchangeTypeID.Text = "" And CmbStateID.Text = "" Then
         MsgBox "Select either State or Client!!"
         Exit Sub
      End If
   End If
End If

If TxtNAVCode <> "" Then
   tmp = tmp & " And TI.NAVCode like '%" & Trim(TxtNAVCode) & "%'"
End If

''-- Ullhas Patil 27 Sep 2017
If Check2.Value = vbChecked Then
    tmp = tmp & " And TI.TIDate Between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' AND '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
End If
''--------

''If CmbExchangeTypeID.Text <> "" Then
''   If Check3.Value = 0 Then
''      tmp = tmp & " And TI.ClientIDRow = " & mClientIDRow & ""
''   Else
''      tmp = tmp & " And TI.DepositoryID ='" & TxtCDPID & "'"
''   End If
''End If

Call Tmp3Table

Call Grid_Head
Call Grid_Head1
If TmpRs3.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs3.RecordCount + 3
   For gr = 1 To TmpRs3.RecordCount
       rn = MSHFlexGrid1.Rows - 1

       MSHFlexGrid1.TextMatrix(rn, 0) = TmpRs3!TINo
       MSHFlexGrid1.TextMatrix(rn, 1) = IIf(IsNull(TmpRs3!TIDate), "", Format(TmpRs3!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(rn, 2) = IIf(IsNull(TmpRs3!TIAmount), 0, TmpRs3!TIAmount)
       MSHFlexGrid1.TextMatrix(rn, 3) = IIf(IsNull(TmpRs3!RecAmount), 0, TmpRs3!RecAmount)
       MSHFlexGrid1.TextMatrix(rn, 4) = IIf(IsNull(TmpRs3!RecTDSAmount), 0, TmpRs3!RecTDSAmount)
       MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs3!RecOtherDed), 0, TmpRs3!RecOtherDed)
       MSHFlexGrid1.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 2)) - (Val(MSHFlexGrid1.TextMatrix(rn, 3)) + Val(MSHFlexGrid1.TextMatrix(rn, 4)) + Val(MSHFlexGrid1.TextMatrix(rn, 5))), "#####0.00")
       MSHFlexGrid1.TextMatrix(rn, 7) = IIf(IsNull(TmpRs3!PostInvoiceNo), 0, TmpRs3!PostInvoiceNo)
       MSHFlexGrid1.TextMatrix(rn, 8) = IIf(IsNull(TmpRs3!EntryNo), 0, TmpRs3!EntryNo)
       MSHFlexGrid1.TextMatrix(rn, 9) = IIf(IsNull(TmpRs3!Locationname), "", TmpRs3!Locationname)
       MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs3!StateName), "", TmpRs3!StateName)
       MSHFlexGrid1.TextMatrix(rn, 13) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
       MSHFlexGrid1.TextMatrix(rn, 14) = TmpRs3!TIID
       'CollectionRemark,CRemarkUpdatedBy,CRemarkUpdatedDate
       MSHFlexGrid1.TextMatrix(rn, 15) = IIf(IsNull(TmpRs3!CollectionRemark), "", TmpRs3!CollectionRemark)
       MSHFlexGrid1.TextMatrix(rn, 16) = IIf(IsNull(TmpRs3!CRemarkUpdatedBy), "", TmpRs3!CRemarkUpdatedBy)
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs3!CRemarkUpdatedDate), "", TmpRs3!CRemarkUpdatedDate)
       
       If MSHFlexGrid1.TextMatrix(rn, 17) <> "" Then
          MSHFlexGrid1.TextMatrix(rn, 17) = Format(MSHFlexGrid1.TextMatrix(rn, 17), "dd-mmm-yyyy")
       End If
       
       If Check1.Value = 1 Then
          If InStr(1, TmpRs3!NavCode, "CSWD") > 0 Then
    
             tmp = "Select top(1) MC.DepositoryID,MC.DepositoryName,ME.ExchangeType from Mst_Depository MC"
             tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
             tmp = tmp & " Where MC.NAVCode='" & TmpRs3!NavCode & "'"
             
             Call Tmp2Table
             If TmpRs2.RecordCount > 0 Then
                MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs2!DepositoryID), "", TmpRs2!DepositoryID)
                MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
    '             MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
             Else
                tmp = "Select top(1) MD.DepositoryID,MD.DepositoryName,ME.ExchangeType from Mst_DepositoryStateWiseGST a"
                tmp = tmp & " Inner join Mst_ExchangeType ME on a.ExchangeTypeID=ME.ExchangeTypeID"
                tmp = tmp & " Inner Join Mst_Depository MD On a.DepositoryID = MD.DepositoryID"
                tmp = tmp & " Where a.GSTNavCode = '" & TmpRs3!NavCode & "'"
                
                Call Tmp2Table
                If TmpRs2.RecordCount > 0 Then
                   MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs2!DepositoryID), "", TmpRs2!DepositoryID)
                   MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
                Else
                   tmp = " Select top(1) MC.ClientID,MC.ClientName,ME.ExchangeType from Mst_Client MC"
                   tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
                   tmp = tmp & " Where MC.NAVCode='" & TmpRs3!NavCode & "' or MC.RNAVCode='" & TmpRs3!NavCode & "'"
                   Call Tmp2Table
                   If TmpRs2.RecordCount > 0 Then
                      MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs2!ClientID), "", TmpRs2!ClientID)
                      MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
                   Else
                      tmp = "Select top(1) MC.ClientID,MC.ClientName,ME.ExchangeType from Mst_Client MC"
                      tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
                      tmp = tmp & " Inner Join Mst_ClientStateWiseGST b On MC.ClientID = b.ClientID"
                      tmp = tmp & " Where b.GSTNavCode='" & TmpRs3!NavCode & "' or b.RGSTNavCode='" & TmpRs3!NavCode & "'"
                      Call Tmp2Table
                      If TmpRs2.RecordCount > 0 Then
                         MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs2!ClientID), "", TmpRs2!ClientID)
                         MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
                      End If
                   End If
                End If
             End If
          Else
             tmp = " Select top(1) MC.ClientID,MC.ClientName,ME.ExchangeType from Mst_Client MC"
             tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
             tmp = tmp & " Where MC.NAVCode='" & TmpRs3!NavCode & "' or MC.RNAVCode='" & TmpRs3!NavCode & "'"
             Call Tmp2Table
             If TmpRs2.RecordCount > 0 Then
                MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs2!ClientID), "", TmpRs2!ClientID)
                MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
    '             MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
             Else
                tmp = "Select top(1) MC.ClientID,MC.ClientName,ME.ExchangeType from Mst_Client MC"
                tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
                tmp = tmp & " Inner Join Mst_ClientStateWiseGST b On MC.ClientID = b.ClientID"
                tmp = tmp & " Where b.GSTNavCode='" & TmpRs3!NavCode & "' or b.RGSTNavCode='" & TmpRs3!NavCode & "'"
                Call Tmp2Table
                If TmpRs2.RecordCount > 0 Then
                   MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs2!ClientID), "", TmpRs2!ClientID)
                   MSHFlexGrid1.TextMatrix(rn, 12) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
                End If
             End If
           End If
       End If
       MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(rn, 3))
       MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(rn, 4))
       MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(rn, 5))
       MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(rn, 6))
       TmpRs3.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   CmdExcel.Enabled = True
   If Check1.Value = 1 Then
      Call FillClientDetails
   End If
Else
   CmdExcel.Enabled = False
End If

Label2.Refresh
Label2.Caption = TmpRs3.RecordCount

MsgBox ("Done!!!")

End Sub

Private Sub CmdExcel_Click()
'Gen_mTimeStamp = GetTimeStamp
Dim Path, strFilName, strFilLocation As Variant
Gen_mTimeStamp = GetTimeStamp
tmp = ""
ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If ans = vbNo Then
'   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseProfitability.xls")
'   Call Xls_Detail_Rpt
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.xls")
   Call Xls_Detail_Rpt
Else
   Path = App.Path
   strFilName = "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.csv"
   strFilLocation = Path + strFilName
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
   Open strFilLocation For Output As #1
   For I = 0 To MSHFlexGrid1.Rows - 2
       tmp = ""
       For C = 0 To MSHFlexGrid1.Cols - 1
          If C = 0 Then
             tmp = tmp & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          ElseIf C = 12 Or C = 15 Then
             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          Else
             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          End If
       Next
       Print #1, tmp
   Next
   Close #1
   MsgBox "Done!!!"
End If

End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Invoice Details"
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 3
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
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
ProgressBar1.Value = 0
Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Client Details"
oWS.ClearArrows
mRow = 1
ProgressBar1.Max = MSHFlexGrid2.Rows + 3
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 2
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click

Check2.Value = vbUnchecked

CmdExcel.Enabled = False

End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 18 '15
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(1, 0) = "TI No (Voucher No)"
MSHFlexGrid1.TextMatrix(1, 1) = "TI Date"
MSHFlexGrid1.TextMatrix(1, 2) = "TI Amount"
MSHFlexGrid1.TextMatrix(1, 3) = "Received Amt"
MSHFlexGrid1.TextMatrix(1, 4) = "TDS Deducted"
MSHFlexGrid1.TextMatrix(1, 5) = "Other Deduction"
MSHFlexGrid1.TextMatrix(1, 6) = "Balance Amt"
MSHFlexGrid1.TextMatrix(1, 7) = "PostInvoice No"
MSHFlexGrid1.TextMatrix(1, 8) = "Entry No"
MSHFlexGrid1.TextMatrix(1, 9) = "Location"
MSHFlexGrid1.TextMatrix(1, 10) = "State Name"
MSHFlexGrid1.TextMatrix(1, 11) = "Client ID"
MSHFlexGrid1.TextMatrix(1, 12) = "Client Name"
'MSHFlexGrid1.TextMatrix(1, 12) = "Exchange Type"
MSHFlexGrid1.TextMatrix(1, 13) = "NAV Code"
MSHFlexGrid1.TextMatrix(1, 14) = "TIID"
MSHFlexGrid1.TextMatrix(1, 15) = "Collection Remarks"
MSHFlexGrid1.TextMatrix(1, 16) = "Collection Remarks updated by"
MSHFlexGrid1.TextMatrix(1, 17) = "Collection Remarks updated on"

MSHFlexGrid1.ColWidth(0) = 1700
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 2200
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 2000
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 2000
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 2000
'MSHFlexGrid1.ColWidth(11) = 2000

MSHFlexGrid1.RowHeight(1) = 800

End Sub

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   CmbDPName.Clear
   Call Grid_Head
   Call Grid_Head1
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   Check3.Value = 0
   Check3.Enabled = False
Else
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   
''   Label25.Visible = True
''   Label26.Visible = True
''   Label28.Visible = True
''
''   CmbCNonNcdexClient.Visible = False
   Check3.Enabled = True
   Call Check3_Click
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbDPName.Visible = True
   CmbDPName.Width = 9400 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbCNonNcdexClient.Visible = False
   CmbDPName.Text = ""
Else
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbCNonNcdexClient.Visible = True
   CmbDPName.Visible = False
   CmbCNonNcdexClient.Text = ""
End If

End Sub

Private Sub MSHFlexGrid1_Click()
mSelRow = MSHFlexGrid2.RowSel
mSelCol = 8
Call ViewInvoice(mSelRow, mSelCol)

End Sub

Private Sub ViewInvoice(irow_ As Variant, icol_ As Variant)

Dim fso As New FileSystemObject
On Error GoTo Error:
'Shell (App.Path & "\Invoices\" & TxtEntryNo & " ")
If MSHFlexGrid1.TextMatrix(irow_, icol_) <> "" Then
   wsPathAdobe = Gen_ReadPathConfigINI("PDFPATH") '"\\10.0.0.8\Reader\AcroRD32.exe" '"C:\Program Files\Adobe\Acrobat 6.0\Reader\AcroRd32.exe" '"\\10.0.0.8\Reader\AcroRD32.exe"
   mFileName = App.Path & "\NavisionInvoice\" & MSHFlexGrid1.TextMatrix(irow_, icol_) & ".pdf"
   'mFileName = App.Path & "\Invoices\" & Replace(TxtPostInvoiceNo, "/", "") & ".pdf"
   If fso.FileExists(mFileName) = True Then
      If wsPathAdobe <> "" Then
         iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
         'Call Log_Print_History(mSM_Prefix, "COMM", TxtOriginalInvoiceNo, TxtEntryNo)
      End If
   Else
      MsgBox "Invoice not found. Please contact CMG Department for same!!!"
   End If
End If
Exit Sub

Error:
      MsgBox "Please contact Technology department!!!"


End Sub


Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then
   TxtCClientName = ""
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
   
End If
If CmbExchangeTypeID.Text = "" Then Exit Sub

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
      "Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.NAVCode,Mst_Client.RNAVCode " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

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


Private Sub CmbCNonNcdexClient_GotFocus()

tmp1 = CmbCNonNcdexClient.Text
'Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)

tmp = "Select MC.ClientName,MC.ClientID,MD.DepositoryName,MC.ClientIDRow,MC.NAVCode,MC.RNAVCode From Mst_Client MC"
tmp = tmp & " inner join Mst_Depository MD on MD.DepositoryID=MC.DepositoryID"
tmp = tmp & " Where MC.ExchangeTypeId=" & mExchangeTypeID

Call TmpTable

CmbCNonNcdexClient.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbCNonNcdexClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientIDRow & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientID
    TmpRs.MoveNext
Next
CmbCNonNcdexClient.Text = tmp1

End Sub

Private Sub CmbCNonNcdexClient_LostFocus()
If CmbCNonNcdexClient.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbCNonNcdexClient, "~")

If UBound(strSplitString) <> 3 Then
   MsgBox "Invalid Selection!!!"
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If

ClientName_ = Trim(strSplitString(0))
ClientNo_ = Trim(strSplitString(1))
DPName_ = Trim(strSplitString(2))
ClientIDRow_ = Trim(strSplitString(3))

TmpRs.MoveFirst
TmpRs.Find "ClientID = " & ClientIDRow_ & ""

If TmpRs.EOF Then
   MsgBox "Invalid Selection "
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If
  
mClientIDRow = TmpRs!ClientID
mClientNAVCode = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
End Sub
Private Sub CmbClientDP_LostFocus()
Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      CmbClientDP.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtCClientName = TmpRs1!ClientName
TxtCDPID = TmpRs1!DepositoryID
TxtCDPName = TmpRs1!DepositoryName
mClientNAVCode = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs1!RNAVCode), "", TmpRs1!RNAVCode)
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientID

End Sub

Private Sub CmbDPName_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select Exchange Type!!!"
   CmbDPName.Clear
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
tmp = CmbDPName.Text

Call TableMst_Depository("Where ExchangeTypeId = " & mExchangeTypeID & " ")

CmbDPName.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDPName.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDPName.Text = tmp

End Sub

Private Sub CmbDPName_LostFocus()

Dim strSplitString() As String
Dim DPID_, DPName_ As String

If CmbDPName.Text = "" Then
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
End If
strSplitString = Split(CmbDPName, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbDPName.SetFocus
   Exit Sub
End If


DPName_ = Trim(strSplitString(0))
DPID_ = Trim(strSplitString(1))

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDPName.SetFocus
   Exit Sub
End If

mDPID = RsMst_Depository!DepositoryID
TxtCDPID = mDPID
TxtCDPName = DPName_ 'CmbDPName.Text
mDPNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
End Sub

Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 5
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.TextMatrix(0, 0) = "Client ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Client Name"
MSHFlexGrid2.TextMatrix(0, 2) = "Exchange Type"
MSHFlexGrid2.TextMatrix(0, 3) = "NAV Code"
MSHFlexGrid2.TextMatrix(0, 4) = "Reservation NAV Code"

MSHFlexGrid2.ColWidth(0) = 1800
MSHFlexGrid2.ColWidth(1) = 3000
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 1600
MSHFlexGrid2.ColWidth(4) = 1600

End Sub

Private Sub FillClientDetails()
Call Grid_Head1
If CmbExchangeTypeID.Text <> "" Then
    If Check3.Value = 1 Then
          tmp = "Select MC.DepositoryID,MC.DepositoryName,ME.ExchangeType,MC.NAVCode from Mst_Depository MC"
          tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
          tmp = tmp & " Where MC.NAVCode='" & mDPNAVCode & "'"
          Call Tmp2Table
          For I = 1 To TmpRs2.RecordCount
             rn = MSHFlexGrid2.Rows - 1
             MSHFlexGrid2.TextMatrix(rn, 0) = IIf(IsNull(TmpRs2!DepositoryID), "", TmpRs2!DepositoryID)
             MSHFlexGrid2.TextMatrix(rn, 1) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
             MSHFlexGrid2.TextMatrix(rn, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
             MSHFlexGrid2.TextMatrix(rn, 3) = IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode)
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             TmpRs2.MoveNext
          Next
          
       
    Else
          tmp = " Select MC.ClientID,MC.ClientName,ME.ExchangeType,MC.NAVCode,MC.RNAVCode from Mst_Client MC"
          tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
          If mClientRNAVCode <> "" Then
             tmp = tmp & " Where MC.NAVCode='" & mClientNAVCode & "' or MC.RNAVCode='" & mClientRNAVCode & "'"
          Else
             tmp = tmp & " Where MC.NAVCode='" & mClientNAVCode & "'"
          End If

          Call Tmp2Table
          For I = 1 To TmpRs2.RecordCount
              rn = MSHFlexGrid2.Rows - 1
              MSHFlexGrid2.TextMatrix(rn, 0) = IIf(IsNull(TmpRs2!ClientID), "", TmpRs2!ClientID)
              MSHFlexGrid2.TextMatrix(rn, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
              MSHFlexGrid2.TextMatrix(rn, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
              MSHFlexGrid2.TextMatrix(rn, 3) = IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode)
              MSHFlexGrid2.TextMatrix(rn, 4) = IIf(IsNull(TmpRs2!RNAVCode), "", TmpRs2!RNAVCode)
              MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
              TmpRs2.MoveNext
          Next
    End If
End If
        
          
End Sub
