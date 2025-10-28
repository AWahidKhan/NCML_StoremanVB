VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ISINMapping 
   Caption         =   "ISIN Mapping"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   14235
   WindowState     =   2  'Maximized
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   15045
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   21
         Top             =   195
         Width           =   3060
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   20
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3165
         TabIndex        =   8
         Top             =   195
         Width           =   3060
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3165
         TabIndex        =   19
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6225
         TabIndex        =   18
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6225
         TabIndex        =   17
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   9165
         TabIndex        =   16
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   9165
         TabIndex        =   15
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   12105
         TabIndex        =   14
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   12105
         TabIndex        =   9
         Top             =   600
         Width           =   2820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7065
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   15045
      Begin VB.Frame frame_MainData 
         Height          =   3075
         Left            =   105
         TabIndex        =   22
         Top             =   150
         Width           =   14775
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
            Left            =   3720
            TabIndex        =   55
            Top             =   1020
            Visible         =   0   'False
            Width           =   255
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
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2580
            Visible         =   0   'False
            Width           =   4815
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
            Left            =   8505
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   2580
            Visible         =   0   'False
            Width           =   4980
         End
         Begin VB.TextBox TxtPreFix 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   10800
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1710
            Width           =   2685
         End
         Begin VB.ComboBox CmbCommodityID 
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
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   2115
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ComboBox CmbCMP 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6480
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   990
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   615
            Visible         =   0   'False
            Width           =   11295
         End
         Begin VB.TextBox TxtClientID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            TabIndex        =   1
            Top             =   240
            Width           =   2640
         End
         Begin VB.TextBox TxtClientName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   6480
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   7000
         End
         Begin VB.TextBox TxtCommodity 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   2115
            Width           =   2640
         End
         Begin VB.TextBox TxtISINID 
            Appearance      =   0  'Flat
            Height          =   325
            Left            =   2160
            MaxLength       =   25
            TabIndex        =   5
            Top             =   1350
            Width           =   2640
         End
         Begin VB.TextBox TxtDPName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   630
            Width           =   7000
         End
         Begin VB.TextBox TxtDPID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   630
            Width           =   2640
         End
         Begin VB.TextBox TxtGrade 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   10800
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2115
            Width           =   2685
         End
         Begin VB.CheckBox Check1 
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
            Left            =   2160
            TabIndex        =   3
            Top             =   1020
            Width           =   255
         End
         Begin VB.TextBox TxtCMPID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1005
            Width           =   7000
         End
         Begin VB.TextBox TxtLocationID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1350
            Width           =   7000
         End
         Begin VB.TextBox TxtExchangeTypeID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   325
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   2115
            Width           =   3045
         End
         Begin VB.CommandButton CmdISIN 
            Caption         =   "..."
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
            Left            =   1680
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1350
            Width           =   375
         End
         Begin VB.CommandButton CmdClient 
            Caption         =   "..."
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
            Left            =   1680
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   240
            Width           =   375
         End
         Begin MSComCtl2.DTPicker NE_Date 
            Height          =   360
            Left            =   2160
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1710
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   47579137
            CurrentDate     =   39864
         End
         Begin MSComCtl2.DTPicker FED 
            Height          =   360
            Left            =   6480
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1710
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   47579137
            CurrentDate     =   39864
         End
         Begin VB.Label Label4 
            Caption         =   "Partial "
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
            Left            =   2880
            TabIndex        =   56
            Top             =   990
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label10 
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
            TabIndex        =   54
            Top             =   2647
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label9 
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
            Left            =   7230
            TabIndex        =   53
            Top             =   2640
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Location ID"
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
            Left            =   9720
            TabIndex        =   50
            Top             =   1815
            Width           =   1020
         End
         Begin VB.Label LblDPName 
            Caption         =   "DP Name"
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
            Left            =   5040
            TabIndex        =   48
            Top             =   645
            Width           =   900
         End
         Begin VB.Label LblDPID 
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
            Height          =   300
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label LblClientID 
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
            Height          =   300
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label LblClientName 
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
            Height          =   300
            Left            =   5040
            TabIndex        =   45
            Top             =   255
            Width           =   1260
         End
         Begin VB.Label LblCommodity 
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
            Height          =   300
            Left            =   120
            TabIndex        =   44
            Top             =   2160
            Width           =   1500
         End
         Begin VB.Label LblISINID 
            Caption         =   "CMSE Lot No"
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
            TabIndex        =   43
            Top             =   1365
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Grade"
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
            Left            =   9720
            TabIndex        =   42
            Top             =   2130
            Width           =   660
         End
         Begin VB.Label LblCMPID 
            Caption         =   "CMP"
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
            Left            =   5040
            TabIndex        =   41
            Top             =   1020
            Width           =   1080
         End
         Begin VB.Label LblNE_Date 
            Caption         =   "NED"
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
            TabIndex        =   40
            Top             =   1740
            Width           =   810
         End
         Begin VB.Label LblFED 
            Caption         =   "FED"
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
            Left            =   5025
            TabIndex        =   39
            Top             =   1740
            Width           =   600
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
            Height          =   300
            Left            =   5040
            TabIndex        =   38
            Top             =   1365
            Width           =   1200
         End
         Begin VB.Label LblExchangeTypeID 
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
            Height          =   300
            Left            =   5025
            TabIndex        =   37
            Top             =   2130
            Width           =   1440
         End
         Begin VB.Label Label2 
            Caption         =   "Rejection"
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
            TabIndex        =   36
            Top             =   990
            Width           =   900
         End
         Begin VB.Label lblNameDid 
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
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3660
         Left            =   120
         TabIndex        =   7
         Top             =   3300
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   6456
         _Version        =   393216
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame0 
      Height          =   1125
      Left            =   16920
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2085
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   10
         Top             =   150
         Width           =   1500
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   11
         Top             =   540
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "Frm_ISINMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mCMPID As Double
Dim mCommodityID As Double
Dim Edit_Flg As Variant
Dim mGSLIDs As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

Call Grid_Head
TxtClientID.Locked = False
TxtISINID.Locked = False
TxtClientID = ""
TxtISINID = ""
mClientIDRow = 0
mExchangeTypeID = GetExchangeID("NCDEX")
mCMPID = 0
mCommodityID = 0
CmdClient.Visible = True
CmdISIN.Visible = True

frame_MainData.Enabled = True
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
CmbClientDP.Visible = False
TxtGrade = ""
TxtCMPID = ""
TxtExchangeTypeID = "NCDEX"

TxtCommodity = ""
TxtLocationID = ""
CmbCMP.Text = ""
CmbCommodityID.Text = ""

TxtCreated = ""
TxtUpdated = ""

MSHFlexGrid2.Enabled = True
Check1.Enabled = True
Check1.Value = 0
If RsGSL_ISIN.RecordCount > 0 Then
   TxtClientID.SetFocus
End If
End Sub



Private Sub Check1_Click()
Call Grid_Head
TxtISINID = ""

If Check1.Value = 1 Then
   LblISINID.Caption = "AR Report No"
   If SaveCmd.Enabled = True Then
      TxtCommodity.Visible = False
      CmbCommodityID.Visible = True
      TxtCMPID.Visible = False
      CmbCMP.Visible = True
      CmbCMP.Text = ""
      CmbCommodityID.Text = ""
      mCommodityID = 0
      mCMPID = 0
      LblFED.Visible = False
      LblNE_Date.Visible = False
      FED.Visible = False
      NE_Date.Visible = False
      CmdISIN.Visible = False
      If RsGSL_ISIN.RecordCount > 0 And SaveCmd.Enabled = True Then
         TxtISINID.SetFocus
      End If
   End If
Else
   LblISINID.Caption = "ISIN ID"
   CmbCommodityID.Visible = False
   TxtCommodity.Visible = True
   CmbCMP.Visible = False
   TxtCMPID.Visible = True
   
   'TxtISINID.SetFocus
   
   CmdISIN.Visible = True
   LblFED.Visible = True
   LblNE_Date.Visible = True
   FED.Visible = True
   NE_Date.Visible = True

End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Check1.Value = 1
   Check1.Enabled = False
Else
   Check1.Enabled = True
End If

End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub


Private Sub CmdISIN_Click()
FrmMst_ISIN.Show
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
TxtClientID.Locked = False
TxtISINID.Locked = False
MSHFlexGrid2.Enabled = True
Check1.Enabled = True
CmdClient.Visible = True
If Check1.Value = 0 Then
   CmdISIN.Visible = True
End If
If Check1.Value = 1 Then
   CmbCMP.Locked = False
   CmbCommodityID.Locked = False
   
   CmbCMP.Visible = True
   CmbCommodityID.Visible = True
   TxtCMPID.Visible = False
   TxtCommodity.Visible = False
   CmbCMP.Text = TxtCMPID
   CmbCommodityID.Text = TxtCommodity
End If
frame_MainData.Enabled = False
Call GButtonLock(Me, False)
Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsGSL_ISIN.MoveFirst
Call Indata
'TmpRs.MoveFirst
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
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
Call TableGSL_ISIN("   ISINID <> '' AND ISINID <> '0' ", "Order By ClientIDRow,ISINID")
If RsGSL_ISIN.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
End Sub
Public Sub Indata()
If RsGSL_ISIN.RecordCount = 0 Then Exit Sub
Call GButtonLock(Me, True)
DeleteCmd.Enabled = False
Dim ans As Variant
ans = GetClientNameByRow(IIf(IsNull(RsGSL_ISIN.Fields(0)), 0, RsGSL_ISIN.Fields(0)))
mClientIDRow = IIf(IsNull(RsGSL_ISIN.Fields(0)), 0, RsGSL_ISIN.Fields(0)) 'RsGSL_ISIN.Fields(0)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)


CmdClient.Visible = False
CmdISIN.Visible = False


If RsGSL_ISIN.Fields(2) = "J" Then
   Check1.Value = 1
   LblISINID.Caption = "AR Report No"
   LblFED.Visible = False
   LblNE_Date.Visible = False
   FED.Visible = False
   NE_Date.Visible = False
   
Else
   Check1.Value = 0
   LblISINID.Caption = "ISIN ID"
   LblFED.Visible = True
   LblNE_Date.Visible = True
   FED.Visible = True
   NE_Date.Visible = True
End If

TxtISINID = RsGSL_ISIN.Fields(1)
mCMPID = RsGSL_ISIN!CMPID
TxtCMPID = GetCMPName(mCMPID)
CmbCMP.Visible = False
TxtLocationID = GetLocationFromCMP(mCMPID)
TxtPreFix = GetPreFix(mCMPID, "P")
mCommodityID = RsGSL_ISIN!CommodityID
TxtCommodity = GetCommodityName(mCommodityID)
CmbCommodityID.Visible = False
mExchangeTypeID = GetExchangeID("NCDEX")
TxtCommodity.Visible = True
If Check1.Value <> 1 Then
    Call TxtISINID_LostFocus
End If

'If Check1.Value = 1 Then
'   CmbCMP.Text = TxtCMPID.Text
'   CmbCMP.Locked = True
'   CmbCommodityID = TxtCommodity.Text
'   CmbCommodityID.Locked = True
'End If
Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

TxtClientID.Locked = True
TxtISINID.Locked = True
TxtExchangeTypeID = "NCDEX"
MSHFlexGrid2.Enabled = False
Check1.Enabled = False
SearchCmd.Enabled = False



End Sub
Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 12
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "   "
MSHFlexGrid2.TextMatrix(1, 1) = "CMP ID"
MSHFlexGrid2.TextMatrix(1, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 3) = "Godown  "
MSHFlexGrid2.TextMatrix(1, 4) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 5) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 6) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 7) = "Electronic Weight"
MSHFlexGrid2.TextMatrix(1, 8) = "GSLID"
MSHFlexGrid2.TextMatrix(1, 9) = "CDTF No"
MSHFlexGrid2.TextMatrix(1, 10) = "Pledged or Withdrawn Bags"
MSHFlexGrid2.TextMatrix(1, 11) = "Pledged or Withdrawn Weight"

MSHFlexGrid2.ColWidth(1) = 800
MSHFlexGrid2.ColWidth(1) = 1000
MSHFlexGrid2.ColWidth(2) = 2200
MSHFlexGrid2.ColWidth(3) = 1600
MSHFlexGrid2.ColWidth(4) = 1800
MSHFlexGrid2.ColWidth(5) = 1200
MSHFlexGrid2.ColWidth(6) = 1200
MSHFlexGrid2.ColWidth(7) = 1500
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(10) = 1400
MSHFlexGrid2.ColWidth(11) = 1400

MSHFlexGrid2.RowHeight(1) = 600
End Sub

Private Sub LastCmd_Click()
RsGSL_ISIN.MoveLast
    Call Indata
    'TmpRs.MoveLast
    PreviousCmd.Enabled = True
    FirstCmd.Enabled = True
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid2_Click()
  If MSHFlexGrid2.Rows > 3 Then
   If MSHFlexGrid2.RowSel <> "0" Then
      If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then Exit Sub
      tmp = mID(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 9), 1, 1)
      If tmp = "*" Then
         MsgBox "CDTF Printout is not Generated You can't select this Record !!!!! "
         Exit Sub
      End If
      If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "" Then
          MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
          MSHFlexGrid2.Col = 0
          MSHFlexGrid2.row = MSHFlexGrid2.RowSel
          MSHFlexGrid2.CellBackColor = &HFFC0C0

      Else
          If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 10)) = 0 And Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 11)) = 0 Then
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
             MSHFlexGrid2.Col = 0
             MSHFlexGrid2.row = MSHFlexGrid2.RowSel
             MSHFlexGrid2.CellBackColor = &H80000005
          Else
             MsgBox "You can't de-select this lot !!!!!"
             Exit Sub
          End If
      End If
   End If
  End If
  
End Sub

Private Sub NextCmd_Click()

Dim LF_Flag As Variant
LF_Flag = "N"
RsGSL_ISIN.MoveNext
If RsGSL_ISIN.EOF Then
   RsGSL_ISIN.MoveLast
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
RsGSL_ISIN.MovePrevious

If RsGSL_ISIN.BOF Then
   RsGSL_ISIN.MoveFirst
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


'Function CheckSelection() As Boolean
'Dim RetVal As Boolean
'RetVal = True
'For u = 2 To MSHFlexGrid2.Rows - 1
'    If MSHFlexGrid2.TextMatrix(u, 0) = "*" Then
'       RetVal = False
'       Exit For
'    End If
'Next
'CheckSelection = RetVal
'End Function

Private Sub SaveCmd_Click()
  
  If TxtClientID = "" Then
    MsgBox "Blank ClientID Not Allowed!!!"
    Exit Sub
  End If
'  If CheckSelection() = True Then
'    MsgBox "Please Select at least one G.S.L. from Grid !!!"
'    Exit Sub
'  End If
  
  If TxtISINID = "" Then
     If Check1.Value = 0 Then
        MsgBox "Blank CMSE Lot No Not Allowed!!!"
     Else
        MsgBox "Blank AR No Not Allowed!!!"
     End If
     Exit Sub
  End If
  
  If Check1.Value = 1 Then
    If CmbCMP.Text = "" Then
      MsgBox "Blank CMP Not Allowed!!!"
      CmbCMP.SetFocus
      Exit Sub
    End If
    If CmbCommodityID.Text = "" Then
      MsgBox "Blank Commodity Not Allowed!!!"
      CmbCommodityID.SetFocus
      Exit Sub
    End If
  End If
  
  If MSHFlexGrid2.Rows > 3 Then
      Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "' and GSLID in " & mGSLIDs & "")
      For I = 2 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
        If RsMst_GSL.RecordCount > 0 Then
           RsMst_GSL.MoveFirst
        End If
        RsMst_GSL.Find ("GSLID = " & Val(MSHFlexGrid2.TextMatrix(I, 8)))
        If Not RsMst_GSL.EOF Then
           If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
              RsMst_GSL!ISINID = Trim(TxtISINID.Text)
              RsMst_GSL!FirstISINID = Trim(TxtISINID.Text)
              If Check1.Value = 1 Then
                 RsMst_GSL!Flag = "J" ' Reject
              End If
'           Else
'              RsMst_GSL!ISINID = ""
'              RsMst_GSL!FirstISINID = ""
'              RsMst_GSL!Flag = "O"
              RsMst_GSL.Update
              RsMst_GSL.Requery
           End If

        End If
      Next
      RsMst_GSL.Close
      RsGSL_ISIN.Close
  End If

cn.Close
'Call Form_Load

Call TableGSL_ISIN("   ISINID <> '' AND ISINID <> '0' ", "Order By ClientIDRow,ISINID")

RsGSL_ISIN.Find "ISINID = '" & TxtISINID.Text & "'"
If RsGSL_ISIN.RecordCount > 0 Then
   If RsGSL_ISIN.EOF Then
      RsGSL_ISIN.MoveFirst
   End If
   Call GButtonLock(Me, True)
   Call Indata
End If

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsGSL_ISIN.RecordCount = 0 Then
      Exit Sub
   End If
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
End Sub

Private Sub TxtClientID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
'If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
mExchangeTypeID = GetExchangeID("NCDEX")
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

      
'      "WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
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
Else
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   End If
End If
'Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

 
End Sub
Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
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
Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

'CmdClientAdd2List.SetFocus
End Sub

Private Sub TxtISINID_LostFocus()
If LblISINID.Caption = "ISIN ID" Then
    If TxtISINID.Text = "" Then
        mCommodityID = 0
        mCMPID = 0
        'mExchangeTypeID = 0
        TxtCMPID = ""
        TxtLocationID = ""
        TxtCommodity = ""
        TxtExchangeTypeID = ""
        TxtGrade = ""
        Call Grid_Head
        Exit Sub
    End If
    Call TableMst_ISIN("where ISINID = '" & TxtISINID.Text & "'")
    
    If RsMst_ISIN.RecordCount = 0 Then
       MsgBox "No Such ISIN Exist !!!"
       'TxtISINID.SetFocus
       Exit Sub
    End If
    mCMPID = RsMst_ISIN!CMPID
    TxtCMPID = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
    NE_Date = IIf(IsNull(RsMst_ISIN!NE_Date), "", RsMst_ISIN!NE_Date)
    FED = IIf(IsNull(RsMst_ISIN!FED), "", RsMst_ISIN!FED)
    mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
    TxtLocationID = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
    mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
    TxtCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
    mExchangeTypeID = IIf(IsNull(RsMst_ISIN!ExchangeTypeID), "", RsMst_ISIN!ExchangeTypeID)
    TxtExchangeTypeID = GetExchangeName(RsMst_ISIN!ExchangeTypeID) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
    TxtGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
    TxtPreFix = GetPreFixByLocation(RsMst_ISIN!LocationID, "P")
    
    'If mClientIDRow <> "" And mExchangeTypeID <> "" And mCMPID <> "" And mCommodityID <> "" Then
     Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
    'End If

End If
 
  
End Sub
Private Sub GetGridData(mClientRowId_, mExchangeID_, mCMPID_, mCommodityID_)
Call Grid_Head
If IsNull(mClientRowId_) Or mClientRowId_ = 0 Then
   Exit Sub
End If
If IsNull(mExchangeID_) Then
   Exit Sub
End If
If IsNull(mCMPID_) Or mCMPID_ = 0 Then
   Exit Sub
End If
If IsNull(mCommodityID_) Or mCommodityID_ = 0 Then
   Exit Sub
End If

'tmp = "SELECT Mst_Godown.GodownNo,Mst_Godown.CMPID,Mst_CMP.CMPName, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, Mst_GSL.DepositeWeight,Mst_GSL.ISINID, "
'tmp = tmp & " Mst_GSL.GSLID FROM Mst_CMP INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) "
'tmp = tmp & " ON (Mst_CMP.CMPID = Mst_GSL.CMPID) AND (Mst_CMP.CMPID = Mst_Godown.CMPID) "

'COmmeneted on 30 May 2013 647 PM Mail from Komal
'tmp = "SELECT Mst_Godown.GodownNo, Mst_Godown.CMPID, Mst_CMP.CMPName, Mst_GSL.StackNo, Mst_GSL.LotNo,"
'tmp = tmp & " Mst_GSL.DepositeBags , Mst_GSL.DepositeWeight, Mst_GSL.ISINID, Mst_GSL.GSLID, Mst_GSL.CDTFNo, "
'tmp = tmp & " Mst_GSL.PledgeBags,Mst_GSL.PledgeWeight,Mst_GSL.CDTFBags,Mst_GSL.CDTFWeight , Mst_GSL.flag,Mst_GSL.MadeupFlag "
'tmp = tmp & " FROM (Mst_CMP INNER JOIN Mst_Godown ON Mst_CMP.CMPID = Mst_Godown.CMPID) INNER JOIN Mst_GSL ON"
'tmp = tmp & " Mst_Godown.GodownID = Mst_GSL.GodownID"
'tmp = tmp & " WHERE (((Mst_Godown.CMPID)= " & mCMPID_ & ") AND ((Mst_GSL.ClientIDRow)=" & mClientRowId_ & ") AND ((Mst_GSL.CommodityID)= " & mCommodityID_ & ") "
'tmp = tmp & " AND ((Mst_GSL.ExchangeTypeID)= " & mExchangeID_ & ") AND "
'tmp = tmp & " ((Mst_GSL.SM_Prefix)= '" & TxtPreFix & "') AND " & Gen_ExpireGSLFlag & " AND "
'
'If SaveCmd.Enabled = False Then
'   tmp = tmp & " (Mst_GSL.ISINID= '" & TxtISINID & "'))"
'ElseIf Edit_Flg = "A" Then
'   tmp = tmp & " (Mst_GSL.ISINID Is Null  OR Mst_GSL.ISINID = '0' OR Mst_GSL.ISINID= ''))"
'Else
'   tmp = tmp & " (Mst_GSL.ISINID Is Null  OR Mst_GSL.ISINID = '0' OR Mst_GSL.ISINID= '' OR Mst_GSL.ISINID = '" & TxtISINID & "'))"
'End If

  
tmp = " SELECT Mst_Godown.GodownNo, Mst_Godown.CMPID, Mst_CMP.CMPName, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags ,"
tmp = tmp & " Mst_GSL.DepositeWeight, Mst_GSL.ISINID, Mst_GSL.GSLID, Mst_GSL.CDTFNo,  Mst_GSL.PledgeBags,Mst_GSL.PledgeWeight,"
tmp = tmp & " MST_GSL.CDTFBags , MST_GSL.CDTFWeight, MST_GSL.Flag, MST_GSL.MadeupFlag"
tmp = tmp & " From MST_GSL"
tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
tmp = tmp & " Where MST_GSL.ClientIDRow = " & mClientRowId_ & " And MST_GSL.CMPID = " & mCMPID_ & " And MST_GSL.CommodityID = " & mCommodityID_ & ""
tmp = tmp & " AND Mst_GSL.ExchangeTypeID= " & mExchangeID_ & " AND  Mst_GSL.SM_Prefix= '" & TxtPreFix & "'"
tmp = tmp & " AND  MST_GSL.Flag <> 'E' And MST_GSL.Flag <> 'Z'  AND  " 'Mst_GSL.ISINID= 'INC000390731'"
  
If SaveCmd.Enabled = False Then
   tmp = tmp & " Mst_GSL.ISINID= '" & TxtISINID & "'"
ElseIf Edit_Flg = "A" Then
   tmp = tmp & " (Mst_GSL.ISINID Is Null  OR Mst_GSL.ISINID = '0' OR Mst_GSL.ISINID= '')"
Else
   tmp = tmp & " (Mst_GSL.ISINID Is Null  OR Mst_GSL.ISINID = '0' OR Mst_GSL.ISINID= '' OR Mst_GSL.ISINID = '" & TxtISINID & "')"
End If
  
Call TmpTable

mGSLIDs = ""

If TmpRs.RecordCount <= 0 Then
    MsgBox "No Record Found For Grid!!!"
    'TxtISINID.SetFocus
    Exit Sub
End If



k = 1
For I = 0 To TmpRs.RecordCount - 1
    If TmpRs!Flag <> "R" Then
       If TmpRs!Madeupflag = "0" Then
          tmp = " Where  TxnType = 'D' And  CxTFNo = " & TmpRs!CdtfNo & " And Sm_preFix = '" & TxtPreFix & "'"
          Call TableTxn_CxTF_Details(tmp)
       End If
       k = k + 1
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       If IsNull(TmpRs!ISINID) Or TmpRs!ISINID = "" Or TmpRs!ISINID = "0" Then
          MSHFlexGrid2.TextMatrix(k, 0) = ""
       Else
          MSHFlexGrid2.TextMatrix(k, 0) = "*"
          MSHFlexGrid2.Col = 0
          MSHFlexGrid2.row = k
          MSHFlexGrid2.CellBackColor = &HFFC0C0
       End If
       MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!CMPID
       MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!CMPName
       MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!GodownNo
       MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!StackNo
       MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!LotNo
       MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!DepositeBags
       MSHFlexGrid2.TextMatrix(k, 7) = TmpRs!DepositeWeight
       MSHFlexGrid2.TextMatrix(k, 8) = TmpRs!GSLID
       If mGSLIDs = "" Then
          mGSLIDs = "(" & Val(MSHFlexGrid2.TextMatrix(k, 8))
       Else
          mGSLIDs = mGSLIDs & "," & Val(MSHFlexGrid2.TextMatrix(k, 8))
       End If
       
       If TmpRs!Madeupflag = "0" Then
          If IsNull(RsTxn_CxTF_Details!Flag) Or RsTxn_CxTF_Details!Flag = "" Then
             MSHFlexGrid2.TextMatrix(k, 9) = "* - " & TmpRs!CdtfNo
          Else
             MSHFlexGrid2.TextMatrix(k, 9) = TmpRs!CdtfNo
          End If
       Else
          MSHFlexGrid2.TextMatrix(k, 9) = ""
       End If
       MSHFlexGrid2.TextMatrix(k, 10) = TmpRs!PledgeBags + TmpRs!CDTFBags
       MSHFlexGrid2.TextMatrix(k, 11) = TmpRs!PledgeWeight + TmpRs!CDTFWeight
    End If
    TmpRs.MoveNext
Next

If mGSLIDs <> "" Then
   mGSLIDs = mGSLIDs & ")"
End If

End Sub
Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID


Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

End Sub

Private Sub CmbCMP_GotFocus()
tmp = CmbCMP.Text
Call TableMst_CMP
CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMP.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMP.Text = tmp
End Sub
'----
Private Sub CmbCMP_LostFocus()
If CmbCMP.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMP.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
TxtLocationID = GetLocationName(RsMst_CMP!LocationID)
Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

End Sub
