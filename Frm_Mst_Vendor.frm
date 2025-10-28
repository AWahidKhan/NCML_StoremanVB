VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Vendor 
   Caption         =   "Vendor Master"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11745
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2415
      Left            =   120
      TabIndex        =   35
      Top             =   8400
      Visible         =   0   'False
      Width           =   15000
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
         Left            =   6630
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   405
         Width           =   3360
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   10260
         TabIndex        =   50
         Top             =   390
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchVendor 
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
         Left            =   75
         MaxLength       =   200
         TabIndex        =   48
         Top             =   405
         Width           =   6480
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   11325
         TabIndex        =   36
         Top             =   390
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1380
         Left            =   75
         TabIndex        =   37
         Top             =   915
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2434
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
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
         Left            =   7920
         TabIndex        =   52
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Vendor Name"
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
         Left            =   2730
         TabIndex        =   49
         Top             =   135
         Width           =   1290
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   240
      TabIndex        =   27
      Top             =   7200
      Width           =   9600
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7620
         TabIndex        =   17
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   5805
         TabIndex        =   34
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3870
         TabIndex        =   33
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2055
         TabIndex        =   32
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7620
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5805
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3870
         TabIndex        =   28
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2055
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8385
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   9825
      Begin VB.TextBox TxtTANNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   10
         Top             =   4850
         Width           =   7110
      End
      Begin VB.TextBox TxtRTGSNo 
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
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   8
         Top             =   3920
         Width           =   3255
      End
      Begin VB.TextBox TxtAccNo 
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
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   7
         Top             =   3920
         Width           =   2775
      End
      Begin VB.CommandButton CmbBranchName 
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
         Height          =   330
         Left            =   9120
         TabIndex        =   42
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton CmdBank 
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
         Height          =   330
         Left            =   9120
         TabIndex        =   41
         Top             =   3022
         Width           =   375
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
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   6255
         Width           =   7095
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
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6720
         Width           =   7095
      End
      Begin VB.ComboBox CmbTDSID 
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
         Left            =   2415
         TabIndex        =   4
         Top             =   2535
         Width           =   6630
      End
      Begin VB.CommandButton CmdTDSID 
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
         Left            =   9120
         TabIndex        =   39
         Top             =   2558
         Width           =   375
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
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   2115
         Width           =   6645
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
         Height          =   315
         Left            =   9105
         TabIndex        =   38
         Top             =   2130
         Width           =   375
      End
      Begin VB.TextBox TxtContactNo 
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
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   9
         Top             =   4385
         Width           =   2775
      End
      Begin VB.TextBox TxtFax 
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
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   12
         Top             =   5355
         Width           =   7085
      End
      Begin VB.TextBox TxtEmailId 
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
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   13
         Top             =   5805
         Width           =   7085
      End
      Begin VB.TextBox TxtMobileNo 
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
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   11
         Top             =   4385
         Width           =   3255
      End
      Begin VB.TextBox TxtAddress 
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
         Height          =   975
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1080
         Width           =   7080
      End
      Begin VB.TextBox TxtVendorName 
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
         Left            =   2415
         MaxLength       =   200
         TabIndex        =   1
         Top             =   660
         Width           =   7080
      End
      Begin VB.TextBox TxtVendorID 
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
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtLocationID 
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
         Left            =   2415
         TabIndex        =   43
         Top             =   2115
         Width           =   6615
      End
      Begin VB.TextBox TxtTDSID 
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
         Left            =   2415
         MaxLength       =   25
         TabIndex        =   44
         Top             =   2535
         Width           =   6615
      End
      Begin VB.ComboBox CmbBranch 
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
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   3465
         Width           =   6645
      End
      Begin VB.TextBox TxtBranchName 
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
         Left            =   2415
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3480
         Width           =   6630
      End
      Begin VB.ComboBox CmbBankName 
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
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   3007
         Width           =   6645
      End
      Begin VB.TextBox TxtBankName 
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
         Left            =   2415
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3000
         Width           =   6630
      End
      Begin VB.Label Label33 
         Caption         =   "TAN Number"
         Height          =   225
         Left            =   240
         TabIndex        =   59
         Top             =   4918
         Width           =   1155
      End
      Begin VB.Label Label16 
         Caption         =   "RTGS No"
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
         Left            =   5280
         TabIndex        =   58
         Top             =   3975
         Width           =   1455
      End
      Begin VB.Label LblAccNo 
         Caption         =   "Account No"
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
         Left            =   240
         TabIndex        =   55
         Top             =   3915
         Width           =   1455
      End
      Begin VB.Label LblBranchName 
         Caption         =   "Branch Name"
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
         Left            =   240
         TabIndex        =   54
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Bank Name"
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
         Left            =   240
         TabIndex        =   53
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   47
         Top             =   6300
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   46
         Top             =   6780
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "TDS Description"
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
         Left            =   240
         TabIndex        =   45
         Top             =   2595
         Width           =   2115
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   40
         Top             =   2175
         Width           =   975
      End
      Begin VB.Label LblContactNo 
         Caption         =   "Contact No"
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
         Left            =   240
         TabIndex        =   26
         Top             =   4445
         Width           =   1695
      End
      Begin VB.Label LblFax 
         Caption         =   "Fax"
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
         Left            =   240
         TabIndex        =   25
         Top             =   5415
         Width           =   1695
      End
      Begin VB.Label LblEmailId 
         Caption         =   "Email Id"
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
         Left            =   240
         TabIndex        =   24
         Top             =   5865
         Width           =   1695
      End
      Begin VB.Label LblMobileNo 
         Caption         =   "Mobile No"
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
         Left            =   5280
         TabIndex        =   23
         Top             =   4452
         Width           =   1695
      End
      Begin VB.Label LblAdess 
         Caption         =   "Address"
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
         Left            =   255
         TabIndex        =   22
         Top             =   1567
         Width           =   1695
      End
      Begin VB.Label LblVendorID 
         Caption         =   "Vendor ID"
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
         Left            =   240
         TabIndex        =   21
         Top             =   292
         Width           =   1695
      End
      Begin VB.Label LblVendorName 
         Caption         =   "Vendor Name"
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
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_Vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mTDSID, mLocationID, mBankID, mBranchID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtVendorName.Locked = False
TxtAddress.Locked = False
TxtLocationID.Locked = True
TxtTDSID.Locked = True
TxtContactNo.Locked = False
TxtMobileNo.Locked = False
TxtFax.Locked = False
TxtEmailId.Locked = False
TxtLocationID.Visible = False
CmbLocationID.Visible = True
TxtTDSID.Visible = False
CmbTDSID.Visible = True
TxtBankName.Visible = False
TxtBranchName.Visible = False
CmbBankName.Visible = True
CmbBranch.Visible = True
''--- 10 Aug 2017 Ullhas Patil
TxtTANNO.Text = ""
TxtTANNO.Locked = False
''----------------------------
TxtCreated = ""
TxtUpdated = ""

TxtVendorID.Text = ""
TxtVendorName.Text = ""
TxtAddress.Text = ""
TxtContactNo.Text = ""
TxtLocationID.Text = ""
TxtTDSID.Text = ""
TxtMobileNo.Text = ""
TxtFax.Text = ""
TxtEmailId.Text = ""
CmbLocationID.Text = ""
CmbTDSID.Text = ""
TxtAccNo = ""
TxtRTGSNo = ""
CmbBankName = ""
CmbBranch = ""
If RsMst_Vendor.RecordCount > 0 Then
   TxtVendorName.SetFocus
End If
End Sub

Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_Bank("")
CmbBranch.Clear
CmbBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbBankName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbBankName.AddItem RsMst_Bank!BankName
    RsMst_Bank.MoveNext
Next
CmbBankName.Text = tmp
End Sub

Private Sub CmbBankName_LostFocus()
If CmbBankName.Text = "" Then
   mBankID = ""
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbBankName.Text & "'"

If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If
mBankID = RsMst_Bank!BankId
End Sub



Private Sub CmbBranch_GotFocus()
tmp = CmbBranch.Text
If CmbBankName = "" Then
   Exit Sub
End If
Call TableMst_Branch(" Where BankID=" & mBankID & "")

CmbBranch.Clear
If RsMst_Branch.RecordCount = 0 Then
   MsgBox "No Branch found"
   CmbBranch.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Branch.RecordCount
    CmbBranch.AddItem RsMst_Branch!BranchName
    RsMst_Branch.MoveNext
Next
CmbBranch.Text = tmp
End Sub

Private Sub CmbBranch_LostFocus()
If CmbBranch.Text = "" Then
   mBranchID = ""
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchName = '" & CmbBranch.Text & "'"

If RsMst_Branch.EOF Then
   MsgBox "Invalid Branch selection "
   CmbBranch.SetFocus
   Exit Sub
End If
mBranchID = RsMst_Branch!BranchId
End Sub

Private Sub CmbBranchName_Click()
FrmMst_Branch.Show
End Sub

Private Sub CmbTDSID_GotFocus()
tmp = CmbTDSID.Text
Call TableMst_TDS
CmbTDSID.Clear
If RsMst_TDS.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_TDS.RecordCount
    CmbTDSID.AddItem RsMst_TDS!Description
    RsMst_TDS.MoveNext
Next
CmbTDSID.Text = tmp
End Sub

Private Sub CmbTDSID_LostFocus()
If CmbTDSID.Text = "" Then
   Exit Sub
End If
RsMst_TDS.MoveFirst
RsMst_TDS.Find "Description = '" & CmbTDSID.Text & "'"
If RsMst_TDS.EOF Then
   MsgBox "Invalid selection "
   CmbTDSID.SetFocus
   Exit Sub
End If
mTDSID = RsMst_TDS!TDSID
End Sub
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
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
End Sub

Private Sub CmdBank_Click()
FrmMst_Bank.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Vendor.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Vendor.xls")
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
Dim wc As Variant
Call Grid_Head
wc = ""

tmp = "Select MV.VendorID,MV.VendorName,ML.LocationName,MTDS.Description,MV.ContactNo,MV.MobileNo, "
tmp = tmp & " MV.Fax,MV.EmailID,MB.Bankname,MBr.Branchname,MV.BankAcNo,MV.IFSCCode,MV.TANNO "
tmp = tmp & " From Mst_Vendor MV "
tmp = tmp & " Left Join Mst_Location ML On MV.LocationID = ML.LocationID "
tmp = tmp & " Left Join Mst_TDS MTDS On MV.TDSID = MTDS.TDSID "
tmp = tmp & " Left Join Mst_Bank MB On MV.BankId = MB.BankId"
tmp = tmp & " Left Join Mst_Branch MBr On MV.BranchID = MBr.BranchID"

If CmbLocationSearch.Text <> "" Then
   wc = " Where ML.LocationName = '" & CmbLocationSearch.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = " Where MV." & Gen_WcLocation
   End If
End If
If TxtSearchVendor <> "" Then
   If wc = "" Then
      wc = " Where MV.VendorName like '%" & TxtSearchVendor & "%'"
   Else
      wc = wc & " And MV.VendorName like '%" & TxtSearchVendor & "%' "
   End If
End If
tmp = tmp & wc
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 1

   For I = 1 To TmpRs.RecordCount
       For c = 0 To MSHFlexGrid1.Cols - 1
   '        MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!VendorID
   '        MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!VendorName
   '        MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!LocationName 'GetLocationName(RsMst_Vendor!LocationID)
   '        MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!Description 'GetTDSDescription(RsMst_Vendor!TDSID)
   '        MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!ContactNo
   '        MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!MobileNo
    '        MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!Fax
   '        MSHFlexGrid1.TextMatrix(I, 7) = TmpRs!EmailID
           If c = 10 Then
              MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", "'" & TmpRs.Fields(c))
           Else
              MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
           End If
       Next
       TmpRs.MoveNext
   Next
Else
  Call Grid_Head
End If

End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdTDSID_Click()
FrmMst_TDS.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsMst_Vendor.Delete
   RsMst_Vendor.Update
   If RsMst_Vendor.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Vendor.MoveNext
   If RsMst_Vendor.EOF() Then
      RsMst_Vendor.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Agent.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtVendorName.Locked = False
TxtAddress.Locked = False
TxtLocationID.Visible = False
CmbLocationID.Visible = True
CmbLocationID.Text = TxtLocationID.Text
CmdLocationID.Visible = True
TxtTDSID.Visible = False
CmbTDSID.Visible = True
CmdTDSID.Visible = True
CmbTDSID.Text = TxtTDSID.Text
TxtBankName.Visible = False
TxtBranchName.Visible = False
CmbBankName.Visible = True
CmbBranch.Visible = True
CmbBankName = TxtBankName
CmbBranch = TxtBranchName
TxtEmailId.Locked = False
TxtMobileNo.Locked = False
TxtContactNo.Locked = False
''--- 10 Aug 2017 Ullhas Patil
TxtTANNO.Locked = False
''----------------------------
TxtFax.Locked = False
TxtVendorName.SetFocus
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_Vendor.MoveFirst
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
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
'Call Grid_Head
TxtSearchVendor = ""
CmbLocationSearch.Text = ""
Call TableMst_Vendor
Frame1.Enabled = True
If RsMst_Vendor.RecordCount = 0 Then
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

Private Sub LastCmd_Click()
RsMst_Vendor.MoveLast
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
RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_Vendor.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Vendor.MoveNext
If RsMst_Vendor.EOF Then
   RsMst_Vendor.MoveLast
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
RsMst_Vendor.MovePrevious
If RsMst_Vendor.BOF Then
   RsMst_Vendor.MoveFirst
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

Private Sub SaveCmd_Click()
If TxtVendorName.Text = "" Then
    MsgBox "Blank Vendor Name Not Allowed!!!"
    TxtVendorName.SetFocus
    Exit Sub
End If
If TxtAddress.Text = "" Then
    MsgBox "Blank Address Not Allowed!!!"
    TxtAddress.SetFocus
    Exit Sub
End If
If CmbLocationID.Text = "" Then
    MsgBox "Blank Location Name Not Allowed!!!"
    CmbLocationID.SetFocus
    Exit Sub
End If
If CmbTDSID.Text = "" Then
   MsgBox "Blank TDS Description Not Allowed!!!"
   CmbTDSID.SetFocus
   Exit Sub
End If
''--- 10 Aug 2017 Ullhas Patil
''If TxtTANNo.Text = "" Then
''   MsgBox "Blank TAN No Allowed !!! "
''   TxtTANNo.SetFocus
''   Exit Sub
''Else
If Trim(TxtTANNO.Text) <> "" Then
   If ValidateTANNo(UCase(TxtTANNO.Text)) = False Then
      MsgBox "Invalid TAN No !!!"
      TxtTANNO.SetFocus
      Exit Sub
   End If
End If
''----------------------------
If Edit_Flg = "A" Then
    If RsMst_Vendor.RecordCount > 0 Then
        RsMst_Vendor.MoveFirst
        RsMst_Vendor.Find "VendorName= '" & TxtVendorName.Text & "'"
        If Not RsMst_Vendor.EOF Then
           MsgBox "Duplicate Vendor Name Not Allowed !!! "
           TxtVendorName.SetFocus
           Exit Sub
        End If
        RsMst_Vendor.MoveFirst
    End If
    ''-- Alter for duplicate TAN No 10 Aug 2017 Ullhas Patil
    RsMst_Vendor.Filter = ""
    RsMst_Vendor.MoveFirst
    If Trim(TxtTANNO.Text) <> "" Then
    RsMst_Vendor.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
    If Not RsMst_Vendor.EOF Then
    Select Case MsgBox("Duplicate TAN No Found with another Vendor, Do you want to save the Vendor with same TAN No?", vbYesNo)
        Case vbNo
            TxtTANNO.SetFocus
            Exit Sub
    End Select
    End If
    End If
    ''----------------------------
    RsMst_Vendor.AddNew
    RsMst_Vendor!VendorID = GetVendorID
    RsMst_Vendor!G_UID = GetGUID
    TxtVendorID.Text = RsMst_Vendor!VendorID
Else
    RsMst_Vendor.MoveFirst
    RsMst_Vendor.Find "VendorName='" & TxtVendorName.Text & "'  "
    If Not RsMst_Vendor.EOF Then
         If RsMst_Vendor!VendorID <> TxtVendorID.Text Then
          MsgBox "Duplicate Vendor Name Not Allowed !!! "
          TxtVendorName.SetFocus
          Exit Sub
       End If
    End If
    ''-- Alter for duplicate TAN No 09 Aug 2017 Ullhas Patil
    RsMst_Vendor.Filter = ""
    RsMst_Vendor.MoveFirst
    If Trim(TxtTANNO.Text) <> "" Then
    RsMst_Vendor.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
    If Not RsMst_Vendor.EOF Then
     If RsMst_Vendor!VendorID <> TxtVendorID.Text Then
       Select Case MsgBox("Duplicate TAN No Found with another Vendor, Do you want to save the Vendor with same TAN No?", vbYesNo)
           Case vbNo
               TxtTANNO.SetFocus
               Exit Sub
       End Select
     End If
    End If
    End If
    ''----------------------------
    RsMst_Vendor.MoveFirst
    RsMst_Vendor.Find " VendorID= " & TxtVendorID.Text
End If

RsMst_Vendor!VendorID = TxtVendorID.Text
RsMst_Vendor!VendorName = TxtVendorName.Text
RsMst_Vendor!VendorAddress = TxtAddress.Text
RsMst_Vendor!LocationID = mLocationID
RsMst_Vendor!ContactNo = TxtContactNo.Text
RsMst_Vendor!MobileNo = TxtMobileNo.Text
RsMst_Vendor!Fax = TxtFax.Text
RsMst_Vendor!EmailID = TxtEmailId.Text
RsMst_Vendor!TDSID = mTDSID
RsMst_Vendor!Flag = "0" 'TxtFlag.Text
''--- 10 Aug 2017 Ullhas Patil
RsMst_Vendor!TANNO = TxtTANNO
''----------------------------
If mBankID <> "" And mBranchID <> "" Then
    RsMst_Vendor!BankId = mBankID
    RsMst_Vendor!BranchId = mBranchID
    RsMst_Vendor!BankAcNo = TxtAccNo
    RsMst_Vendor!IFSCCode = TxtRTGSNo
Else
    RsMst_Vendor!BankId = Null
    RsMst_Vendor!BranchId = Null
    RsMst_Vendor!BankAcNo = Null
    RsMst_Vendor!IFSCCode = Null
End If

If IsNull(RsMst_Vendor!CreatedBy) = True Or RsMst_Vendor!CreatedBy = "" Then
   RsMst_Vendor!CreatedBy = Gen_UserLoginID
   RsMst_Vendor!CreatedDate = Now
Else
   RsMst_Vendor!UpdatedBy = Gen_UserLoginID
   RsMst_Vendor!UpdatedDate = Now
End If
RsMst_Vendor.Update
Call Indata
End Sub
Private Function GetVendorID() As Double
Dim Retval As Double
tmp = "Select max(VendorID) from Mst_Vendor"
Call TmpTable

Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

If Retval = 0 Then
   
End If
Retval = Retval + 1
GetVendorID = Retval
End Function

Public Sub Indata()
TxtVendorID.Locked = True
TxtVendorName.Locked = True
TxtLocationID.Locked = True
CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True
TxtTDSID.Locked = True
CmbTDSID.Visible = False
CmdTDSID.Visible = False
TxtTDSID.Visible = True
TxtBankName.Visible = True
TxtBranchName.Visible = True
CmbBankName.Visible = False
CmbBranch.Visible = False
TxtAddress.Locked = True
TxtContactNo.Locked = True
TxtMobileNo.Locked = True
TxtFax.Locked = True
TxtEmailId.Locked = True
''--- 10 Aug 2017 Ullhas Patil
TxtTANNO.Locked = True
''----------------------------

TxtVendorID = IIf(IsNull(RsMst_Vendor!VendorID), "", RsMst_Vendor!VendorID)

tmp = ""
tmp = "Select MV.VendorID,MV.VendorName,MV.VendorAddress,ML.LocationID,ML.LocationName,MTDS.TDSID,MTDS.Description,MV.ContactNo,MV.MobileNo, "
tmp = tmp & " MV.Fax,MV.EmailID,MB.BankId,MB.Bankname,MBr.BranchId,MBr.Branchname,MV.BankAcNo,MV.IFSCCode,MV.CreatedBy,MV.CreatedDate,MV.UpdatedBy,MV.UpdatedDate,MV.TANNO "
tmp = tmp & " From Mst_Vendor MV "
tmp = tmp & " Left Join Mst_Location ML On MV.LocationID = ML.LocationID "
tmp = tmp & " Left Join Mst_TDS MTDS On MV.TDSID = MTDS.TDSID "
tmp = tmp & " Left Join Mst_Bank MB On MV.BankId = MB.BankId"
tmp = tmp & " Left Join Mst_Branch MBr On MV.BranchID = MBr.BranchID"
tmp = tmp & " where MV.VendorID=" & TxtVendorID & ""

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
    mLocationID = TmpRs1!LocationID
    mBankID = TmpRs1!BankId
    mBranchID = TmpRs1!BranchId
    mTDSID = TmpRs1!TDSID
    TxtVendorName = IIf(IsNull(TmpRs1!VendorName), "", TmpRs1!VendorName)
    TxtLocationID = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
    TxtAddress = IIf(IsNull(TmpRs1!VendorAddress), "", TmpRs1!VendorAddress)
    TxtBankName = IIf(IsNull(TmpRs1!BankName), "", TmpRs1!BankName)
    TxtBranchName = IIf(IsNull(TmpRs1!BranchName), "", TmpRs1!BranchName)
    TxtAccNo = IIf(IsNull(TmpRs1!BankAcNo), "", TmpRs1!BankAcNo)
    TxtRTGSNo = IIf(IsNull(TmpRs1!IFSCCode), "", TmpRs1!IFSCCode)
    TxtContactNo = IIf(IsNull(TmpRs1!ContactNo), "", TmpRs1!ContactNo)
    TxtMobileNo = IIf(IsNull(TmpRs1!MobileNo), "", TmpRs1!MobileNo)
    TxtFax = IIf(IsNull(TmpRs1!Fax), "", TmpRs1!Fax)
    TxtEmailId = IIf(IsNull(TmpRs1!EmailID), "", TmpRs1!EmailID)
    TxtTDSID.Text = IIf(IsNull(TmpRs1!Description), "", TmpRs1!Description)
    ''--- 10 Aug 2017 Ullhas Patil
    TxtTANNO = IIf(IsNull(TmpRs1!TANNO), "", TmpRs1!TANNO)
    ''----------------------------
    TxtCreated = IIf(IsNull(TmpRs1!CreatedBy), "", TmpRs1!CreatedBy) & " :- " & IIf(IsNull(TmpRs1!CreatedDate), "", TmpRs1!CreatedDate)
    TxtUpdated = IIf(IsNull(TmpRs1!UpdatedBy), "", TmpRs1!UpdatedBy) & " :- " & IIf(IsNull(TmpRs1!UpdatedDate), "", TmpRs1!UpdatedDate)
End If

If RsMst_Vendor.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13
MSHFlexGrid1.TextMatrix(0, 0) = "Vendor ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Vendor Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "TDS"
MSHFlexGrid1.TextMatrix(0, 4) = "Contact No"
MSHFlexGrid1.TextMatrix(0, 5) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 6) = "Fax No"
MSHFlexGrid1.TextMatrix(0, 7) = "Email ID"
MSHFlexGrid1.TextMatrix(0, 8) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 9) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 10) = "Account No"
MSHFlexGrid1.TextMatrix(0, 11) = "RTGS No"
MSHFlexGrid1.TextMatrix(0, 12) = "TAN No"
MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Vendor.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 400
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200
RsMst_Vendor.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchVendor = ""
'CmbLocationSearch.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============


'MSHFlexGrid1.Rows = RsMst_Vendor.RecordCount + 1
'
'For I = 1 To RsMst_Vendor.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'        MSHFlexGrid1.TextMatrix(I, 0) = RsMst_Vendor!VendorID
'        MSHFlexGrid1.TextMatrix(I, 1) = RsMst_Vendor!VendorName
'        MSHFlexGrid1.TextMatrix(I, 2) = GetLocationName(RsMst_Vendor!LocationID)
'        MSHFlexGrid1.TextMatrix(I, 3) = GetTDSDescription(RsMst_Vendor!TDSID)
'        MSHFlexGrid1.TextMatrix(I, 4) = RsMst_Vendor!ContactNo
'        MSHFlexGrid1.TextMatrix(I, 5) = RsMst_Vendor!MobileNo
'        MSHFlexGrid1.TextMatrix(I, 6) = RsMst_Vendor!Fax
'        MSHFlexGrid1.TextMatrix(I, 7) = RsMst_Vendor!EmailID
'    Next
'    RsMst_Vendor.MoveNext
'Next

End Sub

Private Sub CmbLocationSearch_GotFocus()
tmp = CmbLocationSearch.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location(" Where LocationID <> 0 ")
End If

CmbLocationSearch.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationSearch.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationSearch.Text = tmp
End Sub
'----
Private Sub CmbLocationSearch_LostFocus()
If CmbLocationSearch.Text = "" Then
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
End Sub

Private Sub TxtTANNo_LostFocus()
If TxtTANNO.Locked = False Then
    If Trim(TxtTANNO.Text) <> "" Then
       TxtTANNO = UCase(TxtTANNO)
       If ValidateTANNo(TxtTANNO) = False Then
        MsgBox "Invalid TAN No"
        TxtTANNO.SetFocus
        Exit Sub
       End If
    End If
End If
End Sub

