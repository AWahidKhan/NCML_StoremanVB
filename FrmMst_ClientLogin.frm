VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ClientLogin 
   Caption         =   "Client Login Master"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   13275
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2685
      Left            =   -2760
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   13800
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   12450
         TabIndex        =   44
         Top             =   660
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchClient 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1545
         MaxLength       =   200
         TabIndex        =   43
         Top             =   675
         Width           =   5430
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   11445
         TabIndex        =   42
         Top             =   660
         Width           =   1005
      End
      Begin VB.OptionButton Option5 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2790
         TabIndex        =   40
         Top             =   240
         Width           =   1470
      End
      Begin VB.OptionButton Option1 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   240
         Width           =   1350
      End
      Begin VB.OptionButton Option4 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox CmbSearchExchange 
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
         ItemData        =   "FrmMst_ClientLogin.frx":0000
         Left            =   8400
         List            =   "FrmMst_ClientLogin.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   675
         Width           =   2910
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   45
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label6 
         Caption         =   "Client ID "
         Height          =   300
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label Label21 
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
         Left            =   9390
         TabIndex        =   48
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label20 
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
         Left            =   6870
         TabIndex        =   47
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Exchange Type"
         Height          =   300
         Left            =   7080
         TabIndex        =   46
         Top             =   720
         Width           =   1320
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   90
      TabIndex        =   26
      Top             =   7440
      Width           =   10335
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   7845
         TabIndex        =   31
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5970
         TabIndex        =   30
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4080
         TabIndex        =   29
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2220
         TabIndex        =   28
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   345
         TabIndex        =   27
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   7845
         TabIndex        =   41
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5970
         TabIndex        =   39
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4095
         TabIndex        =   37
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2220
         TabIndex        =   35
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   345
         TabIndex        =   33
         Top             =   645
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4590
      Left            =   75
      TabIndex        =   0
      Top             =   -15
      Width           =   10335
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2175
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   2445
      End
      Begin VB.TextBox TxtMobileNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6615
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3210
         Width           =   3000
      End
      Begin VB.TextBox TxtTelephoneNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2175
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3210
         Width           =   3000
      End
      Begin VB.TextBox TxtEmailID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2175
         MaxLength       =   200
         TabIndex        =   8
         Top             =   2775
         Width           =   7440
      End
      Begin VB.TextBox TxtOrganization 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2175
         MaxLength       =   255
         TabIndex        =   7
         Top             =   2340
         Width           =   7440
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   2175
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1500
         Width           =   7440
      End
      Begin VB.TextBox TxtClientLoginName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2175
         MaxLength       =   200
         TabIndex        =   4
         Top             =   645
         Width           =   7440
      End
      Begin VB.TextBox TxtDepositorID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   5520
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2175
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   2445
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3645
         Width           =   7410
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
         Left            =   2175
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4080
         Width           =   7410
      End
      Begin VB.TextBox TxtClientLoginID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7155
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   2445
      End
      Begin VB.Label Label13 
         Caption         =   "Password"
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
         Left            =   240
         TabIndex        =   61
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Email ID"
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
         Left            =   270
         TabIndex        =   60
         Top             =   2835
         Width           =   1455
      End
      Begin VB.Label LblMobileNo 
         Caption         =   "MobileNo"
         Height          =   300
         Left            =   5670
         TabIndex        =   25
         Top             =   3210
         Width           =   840
      End
      Begin VB.Label LblTelephoneNo 
         Caption         =   "Telephone No"
         Height          =   300
         Left            =   270
         TabIndex        =   24
         Top             =   3195
         Width           =   1800
      End
      Begin VB.Label LblOrganization 
         Caption         =   "Organization"
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   2385
         Width           =   1800
      End
      Begin VB.Label LblAddress 
         Caption         =   "Address"
         Height          =   300
         Left            =   255
         TabIndex        =   22
         Top             =   1800
         Width           =   1800
      End
      Begin VB.Label LblClientName 
         Caption         =   "Client Login Name"
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
         Left            =   240
         TabIndex        =   21
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label LblClientID 
         Caption         =   "DepositorID"
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   5520
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label5 
         Caption         =   "Client Code"
         Height          =   225
         Left            =   240
         TabIndex        =   19
         Top             =   308
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Client Login ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   18
         Top             =   278
         Width           =   1800
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
         Height          =   375
         Left            =   270
         TabIndex        =   17
         Top             =   3630
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   375
         Left            =   270
         TabIndex        =   16
         Top             =   4110
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   90
      TabIndex        =   50
      Top             =   4560
      Width           =   10320
      Begin VB.ComboBox CmbNonNcdexClient 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   63
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtDPID 
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
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtDPName 
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
         Left            =   7950
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton CmdClientAdd2List 
         Caption         =   "Add to List"
         Height          =   375
         Left            =   8880
         TabIndex        =   13
         Top             =   960
         Width           =   1335
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
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox TxtClientRowID 
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
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1335
         Left            =   120
         TabIndex        =   52
         Top             =   1440
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2355
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox CmbClientDP 
         Height          =   315
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmMst_ClientLogin.frx":0004
         Left            =   120
         List            =   "FrmMst_ClientLogin.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TxtExchangeTypeID 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         Left            =   6975
         TabIndex        =   59
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label11 
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
         Left            =   8775
         TabIndex        =   58
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Client ID"
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
         Left            =   2160
         TabIndex        =   55
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label8 
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
         Left            =   4200
         TabIndex        =   54
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Exchange Type"
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
         Left            =   225
         TabIndex        =   51
         Top             =   150
         Width           =   1755
      End
   End
End
Attribute VB_Name = "FrmMst_ClientLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, I, InmClientRowID, mClientRowID As Variant
Dim mExchangeID, mDepositoryID, mDepositorID, tmpExchangeID As Double
Dim tmpDepositorID As String
Dim mExchangeTypeID As Double

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
'   Call Grid_Head_Client
'   Call Grid_Head_GCTD
   Call ClearClientTabControl
'   Call ClearGCTDTabControl
  
   CmdClientAdd2List.Enabled = True
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID


If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 8340
   CmbNonNcdexClient.Top = TxtClientRowID.Top
   CmbNonNcdexClient.Left = TxtClientRowID.Left
   Label7.Visible = False
   Label3.Visible = False
   Label11.Visible = False
   
   CmbNonNcdexClient.Visible = True
Else
   Label3.Visible = True
   Label7.Visible = True
   Label11.Visible = True
   CmbNonNcdexClient.Visible = False
End If
'Call Check1_Click
End Sub
Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub
Private Sub CmbNonNcdexClient_LostFocus()
On Error Resume Next
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientRowID = RsMst_Client!ClientIDRow
Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
TxtClientRowID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

End Sub

Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      TxtClientRowID.SetFocus
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

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow
CmdClientAdd2List.SetFocus
End Sub


Private Sub CmdClientAdd2List_Click()
Dim r As Variant

If CmbClientDP.Visible = True Then
   MsgBox "Invalid selection !!! "
   Exit Sub
End If
If CmbNonNcdexClient.Visible = True Then 'And Check1.Value = 0 Then
    If CmbNonNcdexClient.Text = "" Then
       MsgBox "Invalid selection !!! "
       CmbNonNcdexClient.SetFocus
       Exit Sub
    End If
End If
If TxtClientRowID = "" Then
   MsgBox "Invalid Client ID !!! "
   Exit Sub
End If

If CheckDuplicateClient = True Then
  MsgBox "Duplicate Client ID  not allowd !!!"
  TxtClientRowID.SetFocus
  Exit Sub
End If

If LCase(CmdClientAdd2List.Caption) = "update list" Then
   r = MSHFlexGrid2.RowSel
Else
  r = MSHFlexGrid2.Rows - 1
  MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
End If
MSHFlexGrid2.TextMatrix(r, 0) = r
MSHFlexGrid2.TextMatrix(r, 1) = CmbExchangeTypeID.Text
MSHFlexGrid2.TextMatrix(r, 2) = TxtClientRowID
MSHFlexGrid2.TextMatrix(r, 3) = TxtClientName
MSHFlexGrid2.TextMatrix(r, 4) = TxtDPID
MSHFlexGrid2.TextMatrix(r, 5) = TxtDPName
MSHFlexGrid2.TextMatrix(r, 6) = mClientRowID

TxtClientRowID.Locked = False
If mExchangeTypeID <> 1 Then
   TxtClientRowID.SetFocus
Else
   CmbNonNcdexClient.SetFocus
End If
TxtClientRowID.Text = ""
TxtClientName.Text = ""
TxtDPID.Text = ""
TxtDPName.Text = ""

End Sub
Function CheckDuplicateClient() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows >= 3 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 2) = TxtClientRowID.Text And MSHFlexGrid2.TextMatrix(u, 4) = TxtDPID.Text Then
          If LCase(CmdClientAdd2List.Caption) = "add to list" Then
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

CheckDuplicateClient = Retval
End Function
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_ClientLogin.MoveFirst
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
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Option4.Value = True
Call Grid_Head1
Frame1.Enabled = True
Frame2.Enabled = False
Call Grid_Head

Call TableMst_ClientLogin
If RsMst_ClientLogin.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
RsMst_ClientLogin.MoveFirst
Call GButtonLock(Me, True)

Call Indata

End Sub
Public Sub Indata()

TxtDepositorID.Locked = True
TxtClientLoginID.Locked = True

TxtClientLoginName.Locked = True
TxtPassword.Locked = True
CmbExchangeTypeID.Visible = False
CmbExchangeTypeID.Locked = True
CmbNonNcdexClient.Visible = False
CmbNonNcdexClient.Locked = True

TxtExchangeTypeID.Visible = True
TxtClientRowID.Text = ""
TxtClientName.Text = ""
TxtDPID.Text = ""
TxtDPName.Text = ""

TxtAddress.Locked = True
TxtOrganization.Locked = True
TxtEmailID.Locked = True
TxtTelephoneNo.Locked = True
TxtMobileNo.Locked = True


TxtClientID = IIf(IsNull(RsMst_ClientLogin!ClientID), "", RsMst_ClientLogin!ClientID)
InmClientRowID = TxtClientID

TxtClientLoginID = IIf(IsNull(RsMst_ClientLogin!ClientLoginID), "", RsMst_ClientLogin!ClientLoginID)

TxtClientLoginName = IIf(IsNull(RsMst_ClientLogin!ClientLoginName), "", RsMst_ClientLogin!ClientLoginName)
TxtPassword.Text = IIf(IsNull(RsMst_ClientLogin!Password), "", RsMst_ClientLogin!Password)
TxtAddress = IIf(IsNull(RsMst_ClientLogin!Address), "", RsMst_ClientLogin!Address)
TxtOrganization = IIf(IsNull(RsMst_ClientLogin!Organization), "", RsMst_ClientLogin!Organization)
TxtEmailID = IIf(IsNull(RsMst_ClientLogin!EmailID), "", RsMst_ClientLogin!EmailID)
TxtTelephoneNo = IIf(IsNull(RsMst_ClientLogin!TelephoneNo), "", RsMst_ClientLogin!TelephoneNo)
TxtMobileNo = IIf(IsNull(RsMst_ClientLogin!MobileNo), "", RsMst_ClientLogin!MobileNo)

TxtCreated = IIf(IsNull(RsMst_ClientLogin!CreatedBy), "", RsMst_ClientLogin!CreatedBy) & " :- " & IIf(IsNull(RsMst_ClientLogin!CreatedDate), "", RsMst_ClientLogin!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ClientLogin!UpdatedBy), "", RsMst_ClientLogin!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ClientLogin!UpdatedDate), "", RsMst_ClientLogin!UpdatedDate)


Call InDataClientDetail

If RsMst_ClientLogin.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub
Private Sub InDataClientDetail()
Dim ans, EID As Variant
Call Grid_Head
Call TableMst_ClientLoginDetail(" Where ClientLoginID= '" & TxtClientLoginID.Text & "' ")
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + RsMst_ClientLoginDetail.RecordCount
For I = 1 To MSHFlexGrid2.Rows - 2
    
    ans = GetClientDetail(RsMst_ClientLoginDetail!ClientIDRow)
    MSHFlexGrid2.TextMatrix(I, 0) = I
    'EID = Ans(0)
    MSHFlexGrid2.TextMatrix(I, 1) = ans(0) 'GetExchangeName(EID)  'Exchange Type
    MSHFlexGrid2.TextMatrix(I, 2) = ans(1) 'ClientID
    MSHFlexGrid2.TextMatrix(I, 3) = ans(2) 'ClientName
    MSHFlexGrid2.TextMatrix(I, 4) = ans(3) 'DepositoryID
    MSHFlexGrid2.TextMatrix(I, 5) = ans(4) 'DepositoryName
    MSHFlexGrid2.TextMatrix(I, 6) = ans(5)
    RsMst_ClientLoginDetail.MoveNext
Next
'MsgBox TmpRs.State

End Sub

Public Function GetClientDetail(CID_ As Double) As String()

Dim Retval(6) As String
Retval(0) = "" 'Client ID
Retval(1) = "" 'Client Name
Retval(2) = "" 'DP ID
Retval(3) = "" 'DP Name
Retval(4) = ""
Retval(5) = ""

tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName,"
tmp = tmp & " Mst_Client.ClientID , Mst_Client.ClientName, Mst_Client.ClientIDRow, Mst_ExchangeType.ExchangeType"
tmp = tmp & " From Mst_Depository"
tmp = tmp & " INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID"
tmp = tmp & " INNER JOIN Mst_ExchangeType ON Mst_ExchangeType.ExchangeTypeID=Mst_Client.ExchangeTypeID"
tmp = tmp & " Where Mst_Client.ClientIDRow = " & CID_ & ""

Call Tmp3Table
If Not TmpRs3.EOF Then
  Retval(0) = TmpRs3!ExchangeType
  Retval(1) = TmpRs3!ClientID
  Retval(2) = TmpRs3!ClientName
  Retval(3) = TmpRs3!DepositoryID
  Retval(4) = TmpRs3!DepositoryName
  Retval(5) = TmpRs3!ClientIDRow
End If
TmpRs3.Close
GetClientDetail = Retval
End Function

Private Sub AddCmd_Click()
Edit_Flg = "A"

Call GButtonLock(Me, False)
TxtClientLoginName.Locked = False
TxtClientLoginID.Locked = False
TxtPassword.Locked = False

CmbNonNcdexClient.Locked = False


CmbExchangeTypeID.Locked = False
CmbExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False

TxtAddress.Locked = False
TxtOrganization.Locked = False
TxtEmailID.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False
TxtPassword.Text = ""
TxtCreated = ""
TxtUpdated = ""
TxtClientID = ""
TxtClientLoginID = ""
TxtFlag = ""
TxtAddress = ""
TxtOrganization = ""
TxtEmailID = ""
TxtTelephoneNo = ""
TxtMobileNo = ""
TxtClientLoginName.Text = ""

Frame2.Enabled = True
CmbExchangeTypeID.Text = ""
CmbNonNcdexClient.Text = ""

Call Grid_Head

If RsMst_ClientLogin.RecordCount > 0 Then
   TxtClientLoginID.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtClientLoginID.Locked = False
TxtExchangeTypeID.Locked = False
CmbExchangeTypeID.Visible = True
CmbExchangeTypeID.Locked = False
CmbNonNcdexClient.Locked = False

TxtExchangeTypeID.Visible = False
TxtClientLoginName.Locked = False
TxtPassword.Locked = False

TxtAddress.Locked = False
TxtOrganization.Locked = False
TxtEmailID.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False
Frame2.Enabled = True
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError



ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)


If ans = vbYes Then
   RsMst_ClientLogin.Find "ClientID = '" & TxtClientID & "'"
   RsMst_ClientLogin.Delete
   RsMst_ClientLogin.Update
   If RsMst_ClientLogin.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ClientLogin.MoveNext
   If RsMst_ClientLogin.EOF() Then
      RsMst_ClientLogin.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_ClientLogin.CancelUpdate
End Sub
'Function CheckDuplicateClient() As Boolean
'Dim RetVal As Boolean
'RetVal = False
'
'For j = 1 To MSHFlexGrid2.Rows - 1
'    If MSHFlexGrid2.TextMatrix(j, 2) = TxtClientLoginID And _
'       MSHFlexGrid2.TextMatrix(j, 1) = CmbExchange.Text And _
'       MSHFlexGrid2.TextMatrix(j, 4) = CmbDepositoryID.Text Then
'       If CmdAdd2List.Caption = "Add to List" Then
'          RetVal = True
'          Exit For
'       Else
'          If j <> MSHFlexGrid2.RowSel Then
'             RetVal = True
'             Exit For
'          End If
'       End If
'    End If
'Next
'CheckDuplicateClient = RetVal
'End Function

Private Sub LastCmd_Click()
RsMst_ClientLogin.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub



Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ClientLogin.MoveNext
If RsMst_ClientLogin.EOF Then
   RsMst_ClientLogin.MoveLast
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
RsMst_ClientLogin.MovePrevious
If RsMst_ClientLogin.BOF Then
   RsMst_ClientLogin.MoveFirst
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

If TxtClientLoginID.Text = "" Then
   MsgBox "Blank Client Login ID Not Allowed !!! "
   TxtClientLoginID.SetFocus
   Exit Sub
End If

If TxtClientLoginName.Text = "" Then
   MsgBox "Blank Client Login Name Not Allowed !!! "
   TxtClientLoginName.SetFocus
   Exit Sub
End If

If TxtPassword.Text = "" Then
   MsgBox "Blank Password Not Allowed !!! "
   TxtPassword.SetFocus
   Exit Sub
End If

If TxtEmailID.Text = "" Then
   MsgBox "Blank Email ID Not Allowed !!! "
   TxtEmailID.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsMst_ClientLogin.RecordCount > 0 Then
      RsMst_ClientLogin.MoveFirst
      tmp = "ClientLoginID = '" & TxtClientLoginID.Text & "' " 'And DepositoryID = '" & mDepositoryID & "' And ExchangeTypeID = " & mExchangeTypeID & ""
      RsMst_ClientLogin.Filter = tmp
      If Not RsMst_ClientLogin.EOF Then
         MsgBox "Duplicate Client Login ID Not Allowed !!!"
         TxtClientLoginID.SetFocus
         Exit Sub
         
      End If
   End If
   RsMst_ClientLogin.Filter = ""
   RsMst_ClientLogin.AddNew
   RsMst_ClientLogin!ClientID = GetClientID
   RsMst_ClientLogin!G_UID = GetGUID
Else
   RsMst_ClientLogin.MoveFirst
   tmp = "ClientLoginID = '" & TxtClientLoginID.Text & "' " 'And DepositoryID = '" & mDepositoryID & "' And ExchangeTypeID = " & mExchangeTypeID & ""
   RsMst_ClientLogin.Filter = tmp
   If Not RsMst_ClientLogin.EOF Then
      If RsMst_ClientLogin!ClientLoginID <> TxtClientLoginID.Text Then
         MsgBox "Duplicate Client Login ID  Not Allowed !!! "
         TxtClientLoginID.SetFocus
         Exit Sub
      End If
   End If
   RsMst_ClientLogin.Filter = ""
   Call DeleteChildTableDetail
   RsMst_ClientLogin.MoveFirst
   RsMst_ClientLogin.Find "ClientID = " & TxtClientID.Text & ""
   If RsMst_ClientLogin.EOF Then
      MsgBox "No Such Record Found!!! "
      Exit Sub
   End If
End If

RsMst_ClientLogin!Password = TxtPassword.Text
RsMst_ClientLogin!ClientLoginID = TxtClientLoginID.Text
RsMst_ClientLogin!ClientLoginName = TxtClientLoginName.Text
RsMst_ClientLogin!Address = TxtAddress.Text
RsMst_ClientLogin!Organization = TxtOrganization.Text
RsMst_ClientLogin!EmailID = TxtEmailID.Text
RsMst_ClientLogin!TelephoneNo = TxtTelephoneNo.Text
RsMst_ClientLogin!MobileNo = TxtMobileNo.Text


If IsNull(RsMst_ClientLogin!CreatedBy) = True Or RsMst_ClientLogin!CreatedBy = "" Then
   RsMst_ClientLogin!CreatedBy = Gen_UserLoginID
   RsMst_ClientLogin!CreatedDate = Now
Else
   RsMst_ClientLogin!UpdatedBy = Gen_UserLoginID
   RsMst_ClientLogin!UpdatedDate = Now
End If

RsMst_ClientLogin.Update

Call SaveClientDetail
Call Indata
Call GButtonLock(Me, True)
End Sub
Private Sub SaveClientDetail()
Call TableMst_ClientLoginDetail("where ClientLoginID= '" & TxtClientLoginID.Text & "'") 'Txn_CxTF_ClientDetail(" Where TxnType = 'D'  And  CxTFNo = " & TxtCxTFNo & " And SM_Prefix = '" & TxtPreFix & "'")

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       RsMst_ClientLoginDetail.AddNew
       RsMst_ClientLoginDetail!ClientDetailID = GetMaxClientDetailID
       RsMst_ClientLoginDetail!ClientIDRow = MSHFlexGrid2.TextMatrix(I, 6)
       RsMst_ClientLoginDetail!ClientLoginID = Trim(TxtClientLoginID.Text)
       RsMst_ClientLoginDetail.Update
       RsMst_ClientLoginDetail.Requery
    End If
Next
End Sub
Sub DeleteChildTableDetail()
Dim adjWt As Variant
Call TableMst_ClientLoginDetail("where ClientLoginID = '" & Trim(TxtClientLoginID.Text) & "'")

For I = 1 To RsMst_ClientLoginDetail.RecordCount
    RsMst_ClientLoginDetail.Delete
    RsMst_ClientLoginDetail.Update
    RsMst_ClientLoginDetail.MoveNext
Next
End Sub

Private Function GetClientID() As Double
Dim Retval As Double
tmp = "Select max(ClientID) from Mst_ClientLogin "
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
TmpRs1.Close
GetClientID = Retval
End Function

Private Function GetMaxClientDetailID() As Double
Dim Retval As Double
tmp = "Select max(ClientDetailID) from Mst_ClientLoginDetail "
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
TmpRs1.Close
GetMaxClientDetailID = Retval
End Function
Private Function Grid_Head()
With MSHFlexGrid2
     .Clear
     .Cols = 7
     .Rows = 2
     .TextMatrix(0, 0) = "Sr. No"
     .TextMatrix(0, 1) = "Exchange Type"
     .TextMatrix(0, 2) = "Client ID"
     .TextMatrix(0, 3) = "Client Name"
     .TextMatrix(0, 4) = "DP ID"
     .TextMatrix(0, 5) = "DP Name"
     .TextMatrix(0, 6) = "ClientRowID"
     .ColWidth(0) = 600
     .ColWidth(1) = 1500
     .ColWidth(2) = 800
     .ColWidth(3) = 3000
     .ColWidth(4) = 1000
     .ColWidth(5) = 3000
     .ColWidth(6) = 0
End With
End Function
Private Function Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12
'MSHFlexGrid1.TextMatrix(0, 0) = "ClientID"
'MSHFlexGrid1.TextMatrix(0, 1) = "Client Login ID"
'MSHFlexGrid1.TextMatrix(0, 2) = "Client Login Name"
'MSHFlexGrid1.TextMatrix(0, 3) = "ExchangeType"
'MSHFlexGrid1.TextMatrix(0, 4) = "DepositoryID"
'MSHFlexGrid1.TextMatrix(0, 5) = "Flag"
'MSHFlexGrid1.TextMatrix(0, 6) = "Address"
'MSHFlexGrid1.TextMatrix(0, 7) = "Organization"
'MSHFlexGrid1.TextMatrix(0, 8) = "Email"
'MSHFlexGrid1.TextMatrix(0, 9) = "Tel No"
'MSHFlexGrid1.TextMatrix(0, 10) = "Mobile No"
'MSHFlexGrid1.TextMatrix(0, 11) = "Row ID"

MSHFlexGrid1.TextMatrix(0, 0) = "ClientID"
MSHFlexGrid1.TextMatrix(0, 1) = "Client Login ID"
MSHFlexGrid1.TextMatrix(0, 2) = "Client Login Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Address"
MSHFlexGrid1.TextMatrix(0, 4) = "Organization"
MSHFlexGrid1.TextMatrix(0, 5) = "Email "
MSHFlexGrid1.TextMatrix(0, 6) = "Tel No"
MSHFlexGrid1.TextMatrix(0, 7) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 8) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 9) = "ExchangeType"
MSHFlexGrid1.TextMatrix(0, 10) = "DP Name"
MSHFlexGrid1.TextMatrix(0, 11) = "Client Row ID"


MSHFlexGrid1.ColWidth(0) = 650
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1800
MSHFlexGrid1.ColWidth(11) = 0
End Function
Private Sub CmdGo_Click()


tmp = "SELECT  Mst_ClientLogin.ClientID,Mst_ClientLogin.ClientLoginID, Mst_ClientLogin.ClientLoginName, Mst_ClientLogin.Address,"
tmp = tmp & " Mst_ClientLogin.Organization, Mst_ClientLogin.EmailID, Mst_ClientLogin.TelephoneNo, Mst_ClientLogin.MobileNo,"
tmp = tmp & " Mst_Client.ClientName, Mst_ExchangeType.ExchangeType,Mst_Depository.DepositoryID+ '~' + Mst_Depository.DepositoryName 'DP Name',Mst_ClientLoginDetail.ClientIDRow"
tmp = tmp & " FROM Mst_ClientLogin INNER JOIN"
tmp = tmp & " Mst_ClientLoginDetail ON Mst_ClientLogin.ClientLoginID = Mst_ClientLoginDetail.ClientLoginID INNER JOIN"
tmp = tmp & " Mst_Client ON Mst_ClientLoginDetail.ClientIDRow = Mst_Client.ClientIDRow INNER JOIN"
tmp = tmp & " Mst_ExchangeType ON Mst_Client.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID INNER JOIN"
tmp = tmp & " Mst_Depository ON Mst_Client.DepositoryID = Mst_Depository.DepositoryID"



'tmp = "Select MC.ClientID, MC.ClientLoginID,MC.ClientLoginName,ME.ExchangeType,MC.DepositoryID,  MC.Flag ,MC.Address ,"
'tmp = tmp & " MD.DepositoryID + '~' + MD.DepositoryName As [DP Name] ,"
'tmp = tmp & " MC.Organization , MC.EmailID, MC.TelephoneNo, MC.MobileNo"
'tmp = tmp & " from Mst_ClientLogin MC Inner Join  Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID"
'tmp = tmp & " Inner Join  Mst_Depository MD on MC.DepositoryID = MD.DepositoryID"


Dim wc As Variant
wc = ""
If Option5.Value = True Then
   wc = " Where Mst_ClientLogin.ClientLoginName Like '%" & TxtSearchClient & "%' "
End If
If Option1.Value = True Then
   wc = " Where Mst_ClientLogin.ClientLoginID  Like '%" & TxtSearchClient & "%'"
End If
If CmbSearchExchange.Text <> "" Then
   If wc = "" Then
      wc = " Where Mst_ExchangeType.ExchangeTypeID = " & mExchangeID & " "
   Else
      wc = wc & " And Mst_ExchangeType.ExchangeTypeID = " & mExchangeID & " "
   End If
End If
tmp = tmp & " " & wc & " ORDER BY Mst_ClientLogin.ClientID"
Call TmpTable

Call Grid_Head1
If TmpRs.RecordCount = 0 Then
   MsgBox "Record Not Found!!!"
   Exit Sub
End If
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    TmpRs.MoveNext
Next


End Sub

Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType
CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    'If RsMst_ExchangeType.Fields(0) <> 1 Then
    CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    'End If
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp

End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   Exit Sub
End If

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchange.SetFocus
   Exit Sub
End If
If CmbExchange.Text = "SPOT" Then
  Label3.Caption = "TM Name"
Else
  Label3.Caption = "Depository Name"
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
CmbDepositoryID.Text = ""
End Sub
Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsMst_ClientLogin.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 80
Frame0.Top = Frame1.Top - 50
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 10
'Option4.Value = True
'Call Grid_Head1
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ClientLogin.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ClientLogin.xls")
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
Private Sub TxtClientRowID_LostFocus()
CmbClientDP.Visible = False
If TxtClientRowID = "" Then Exit Sub
If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
ans = GetClientNameDtl(TxtClientRowID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientRowID.SetFocus
   Exit Sub
End If
If ans(4, 0) > 1 Then
   CmbClientDP.Visible = True
   CmbClientDP.SetFocus
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientRowID & "'));"

'      "WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    mClientRowID = TmpRs1!ClientIDRow
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
Private Sub MSHFlexGrid2_Click()
If MSHFlexGrid2.Rows = 2 Then Exit Sub
CmdClientAdd2List.Caption = "Update List"
TxtClientRowID.Locked = False
r = MSHFlexGrid2.RowSel
If r <= 0 Then Exit Sub

TxtClientRowID = MSHFlexGrid2.TextMatrix(r, 2)
TxtClientName = MSHFlexGrid2.TextMatrix(r, 3)
TxtDPID = MSHFlexGrid2.TextMatrix(r, 4)
TxtDPName = MSHFlexGrid2.TextMatrix(r, 5)
CmbExchangeTypeID.Text = MSHFlexGrid2.TextMatrix(r, 1)
If CmbNonNcdexClient.Visible = True Then
   CmbNonNcdexClient.Text = MSHFlexGrid2.TextMatrix(r, 3)
End If
    
End Sub

Private Sub CmbSearchExchange_GotFocus()
mExchangeID = Null
tmp = CmbSearchExchange.Text
Call TableMst_ExchangeType
CmbSearchExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSearchExchange.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSearchExchange.Text = tmp

End Sub

Private Sub CmbSearchExchange_LostFocus()
Dim x As Integer
If CmbSearchExchange.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSearchExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSearchExchange.SetFocus
   Exit Sub
End If
mExchangeID = RsMst_ExchangeType!ExchangeTypeID

End Sub

Sub ClearClientTabControl()

TxtClientRowID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
CmdClientAdd2List.Caption = "Add to List"
End Sub
