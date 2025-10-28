VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Depository 
   Caption         =   "DP Master"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2925
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
      Width           =   18870
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "Frm_Mst_Depository.frx":0000
         Left            =   12120
         List            =   "Frm_Mst_Depository.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   1695
      End
      Begin VB.ComboBox CmbSRiskType 
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
         ItemData        =   "Frm_Mst_Depository.frx":0021
         Left            =   9960
         List            =   "Frm_Mst_Depository.frx":002E
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   270
         Width           =   1245
      End
      Begin VB.ComboBox CmbSReturn 
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
         ItemData        =   "Frm_Mst_Depository.frx":0045
         Left            =   7560
         List            =   "Frm_Mst_Depository.frx":004F
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox TxtSearchDP 
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
         Left            =   1185
         TabIndex        =   30
         Top             =   270
         Width           =   5430
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
         Left            =   13890
         TabIndex        =   35
         Top             =   278
         Width           =   1005
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
         Left            =   14910
         TabIndex        =   38
         Top             =   275
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1695
         Left            =   75
         TabIndex        =   37
         Top             =   1080
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   2990
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   11280
         TabIndex        =   49
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Risk Type"
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
         Left            =   8880
         TabIndex        =   46
         Top             =   285
         Width           =   960
      End
      Begin VB.Label Label31 
         Caption         =   "Return"
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
         Left            =   6690
         TabIndex        =   42
         Top             =   345
         Width           =   840
      End
      Begin VB.Label Label19 
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
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   300
         Width           =   1095
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
         Left            =   11280
         TabIndex        =   36
         Top             =   765
         Width           =   2205
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
         Left            =   13560
         TabIndex        =   34
         Top             =   765
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4410
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   8085
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
         ItemData        =   "Frm_Mst_Depository.frx":005C
         Left            =   1800
         List            =   "Frm_Mst_Depository.frx":0066
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   2520
         Width           =   6135
      End
      Begin VB.TextBox TxtStatus 
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
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2520
         Width           =   6135
      End
      Begin VB.TextBox TxtTANNo 
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
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2040
         Width           =   6120
      End
      Begin VB.ComboBox CmbRiskType 
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
         ItemData        =   "Frm_Mst_Depository.frx":007D
         Left            =   1800
         List            =   "Frm_Mst_Depository.frx":008A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1580
         Visible         =   0   'False
         Width           =   5670
      End
      Begin VB.TextBox TxtRiskType 
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
         Left            =   1815
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   45
         Top             =   1580
         Width           =   6120
      End
      Begin VB.CheckBox ChkReturn 
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   2
         Top             =   293
         Width           =   255
      End
      Begin VB.TextBox TxtNAVDPName 
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2955
         Width           =   6120
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3840
         Width           =   6120
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3400
         Width           =   6120
      End
      Begin VB.CommandButton CmdExchangeType 
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
         Height          =   360
         Left            =   7545
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CmbExchange 
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
         TabIndex        =   4
         Top             =   1150
         Visible         =   0   'False
         Width           =   5685
      End
      Begin VB.TextBox TxtDepositoryID 
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   4320
      End
      Begin VB.TextBox TxtDepositoryName 
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
         MaxLength       =   200
         TabIndex        =   3
         Top             =   700
         Width           =   6120
      End
      Begin VB.TextBox TxtExchangeType 
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
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1150
         Width           =   6120
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   120
         TabIndex        =   48
         Top             =   2580
         Width           =   660
      End
      Begin VB.Label Label33 
         Caption         =   "TAN Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   47
         Top             =   2115
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Risk Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   44
         Top             =   1655
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cheque Return"
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
         Left            =   6600
         TabIndex        =   43
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "NAVISION DP Name"
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
         TabIndex        =   41
         Top             =   3015
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
         Left            =   120
         TabIndex        =   29
         Top             =   3900
         Width           =   975
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
         TabIndex        =   28
         Top             =   3450
         Width           =   1695
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
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   1252
         Width           =   1440
      End
      Begin VB.Label LblDepositoryID 
         Caption         =   "DP ID"
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
         TabIndex        =   21
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblDepositoryName 
         Caption         =   "DP Name"
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
         TabIndex        =   22
         Top             =   730
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   23
      Top             =   4500
      Width           =   8085
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
         Left            =   225
         TabIndex        =   16
         Top             =   640
         Width           =   1515
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
         Left            =   1740
         TabIndex        =   17
         Top             =   640
         Width           =   1515
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
         Left            =   3255
         TabIndex        =   18
         Top             =   640
         Width           =   1515
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
         Left            =   4770
         TabIndex        =   19
         Top             =   640
         Width           =   1515
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
         Left            =   6285
         TabIndex        =   12
         Top             =   640
         Width           =   1515
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
         Left            =   6285
         TabIndex        =   11
         Top             =   240
         Width           =   1515
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
         Left            =   4770
         TabIndex        =   15
         Top             =   240
         Width           =   1515
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
         Left            =   3255
         TabIndex        =   14
         Top             =   240
         Width           =   1515
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
         Left            =   1740
         TabIndex        =   10
         Top             =   240
         Width           =   1515
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
         Left            =   225
         TabIndex        =   13
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "FrmMst_Depository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mExchangeTypeID As Double

Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType
CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    If RsMst_ExchangeType.Fields(0) <> 1 Then
       CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    End If
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
  LblDepositoryName.Caption = "TM Name"
Else
  LblDepositoryName.Caption = "Depository Name"
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub

Private Sub CmbSExchangeType_GotFocus()

tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    If RsMst_ExchangeType.Fields(0) <> 1 Then
       CmbSExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    End If
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeType.Text = tmp

End Sub
Private Sub CmbSExchangeType_LostFocus()

If CmbSExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeType.SetFocus
   Exit Sub
End If

End Sub






Private Sub CmbRiskType_LostFocus()
If CmbRiskType.Text = "" Then Exit Sub

If LCase(CmbRiskType.Text) <> "low" And LCase(CmbRiskType.Text) <> "high" And LCase(CmbRiskType.Text) <> "medium" Then
   MsgBox "Invalid selection!!! "
   CmbRiskType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSReturn_LostFocus()
If CmbSReturn.Text = "" Then Exit Sub
If LCase(CmbSReturn.Text) <> "yes" And LCase(CmbSReturn.Text) <> "no" Then
   MsgBox "Invalid Selection"
   CmbSReturn.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSRiskType_LostFocus()
If CmbSRiskType.Text = "" Then Exit Sub

If LCase(CmbSRiskType.Text) <> "low" And LCase(CmbSRiskType.Text) <> "high" And LCase(CmbSRiskType.Text) <> "medium" Then
   MsgBox "Invalid selection!!! "
   CmbSRiskType.SetFocus
   Exit Sub
End If
End Sub



''-- Ullhas Patil 21 Oct 2017
Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub
Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub
''-----------------------------

Private Sub CmdExchangeType_Click()
FrmMst_ExchangeType.Show
End Sub

Private Sub CmdGo_Click()
tmp = "Select MD.DepositoryID,MD.DepositoryName,ME.ExchangeType,MD.G_UID,MD.NAVDepositoryName, MD.ChequeReturn, MD.RiskType,MD.TANNO,MD.Status " & _
      " From Mst_Depository MD " & _
      " Inner Join Mst_ExchangeType ME On MD.ExchangeTypeID = ME.ExchangeTypeID "

Dim wc As Variant
wc = ""
If TxtSearchDP <> "" Then
   wc = " Where DepositoryName Like '%" & TxtSearchDP & "%' "
End If
If LCase(CmbSReturn.Text) = "yes" Then
   If wc <> "" Then
      wc = wc & " AND MD.ChequeReturn='Y'"
   Else
      wc = " Where MD.ChequeReturn='Y'"
   End If
ElseIf LCase(CmbSReturn.Text) = "no" Then
   If wc <> "" Then
      wc = wc & " AND MD.ChequeReturn='N'"
   Else
      wc = " Where MD.ChequeReturn='N'"
   End If
End If
If CmbSRiskType.Text <> "" Then
   If wc <> "" Then
      wc = wc & " AND MD.RiskType='" & CmbSRiskType & "'"
   Else
      wc = " Where MD.RiskType='" & CmbSRiskType & "'"
   End If
End If

If CmbSStatus.Text <> "" Then
   If wc <> "" Then
      wc = wc & " AND MD.Status = '" & Left(CmbSStatus.Text, 1) & "'"
   Else
      wc = " Where MD.Status = '" & Left(CmbSStatus.Text, 1) & "'"
   End If
End If

'If CmbSExchangeType.Text <> "" Then
'   If TxtSearchDP <> "" Then
'      wc = " Where ExchangeType = '%" & CmbSExchangeType & "%' "
'   Else
'      wc = wc & " And ExchangeType Like '%" & CmbSExchangeType & "%' "
'   End If
'End If
tmp1 = " Order By ME.ExchangeType,MD.DepositoryName"
tmp = tmp & wc & tmp1

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If LCase(MSHFlexGrid1.TextMatrix(iu, 5)) = "y" Then
        MSHFlexGrid1.TextMatrix(iu, 5) = "Yes"
    Else
        MSHFlexGrid1.TextMatrix(iu, 5) = "No"
    End If
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 8)) = "a" Then
        MSHFlexGrid1.TextMatrix(iu, 8) = "Active"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 8)) = "d" Then
        MSHFlexGrid1.TextMatrix(iu, 8) = "De-Active"
    Else
        MSHFlexGrid1.TextMatrix(iu, 8) = ""
    End If
    
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

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
TxtSearchDP = ""
Label22.Caption = ""

Call TableMst_Depository
If RsMst_Depository.RecordCount = 0 Then
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
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Depository.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Depository.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 2
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Depository.MoveNext
If RsMst_Depository.EOF Then
   RsMst_Depository.MoveLast
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
RsMst_Depository.MovePrevious
If RsMst_Depository.BOF Then
   RsMst_Depository.MoveFirst
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
RsMst_Depository.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Depository.MoveLast
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
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Depository.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If

Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Height = Me.Height - 900
Frame0.Width = Me.Width - 400
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200
RsMst_Depository.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 2
'Call Grid_Head
'TxtSearchDP = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

'MSHFlexGrid1.Rows = RsMst_Depository.RecordCount + 1
'MSHFlexGrid1.ColWidth(0) = 1000
'For i = 1 To RsMst_Depository.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_Depository.Fields(c)), "", RsMst_Depository.Fields(c))
'    Next
'    RsMst_Depository.MoveNext
'Next
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 9
    
    .TextMatrix(0, 0) = "DepositoryID":             .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "Depositor Name":           .ColWidth(1) = 3200
    .TextMatrix(0, 2) = "Exchange Type":            .ColWidth(2) = 3200
    .TextMatrix(0, 3) = "GUID":                     .ColWidth(3) = 0
    .TextMatrix(0, 4) = "NAVISION DP Name":         .ColWidth(4) = 3200
    .TextMatrix(0, 5) = "Return Flag":              .ColWidth(5) = 1200
    .TextMatrix(0, 6) = "Risk Type":                .ColWidth(6) = 1200
    
    .TextMatrix(0, 7) = "TAN No":                   .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "Status":                   .ColWidth(8) = 1500

End With
End Sub
Public Sub Indata()

TxtDepositoryID.Locked = True
TxtDepositoryName.Locked = True

TxtExchangeType.Visible = True
CmdExchangeType.Visible = False
CmbExchange.Visible = False
CmbExchange.Text = ""

TxtRiskType.Visible = True
CmbRiskType.Visible = False

TxtDepositoryID = IIf(IsNull(RsMst_Depository!DepositoryID), "", RsMst_Depository!DepositoryID)
TxtDepositoryName = IIf(IsNull(RsMst_Depository!DepositoryName), "", RsMst_Depository!DepositoryName)
TxtRiskType = IIf(IsNull(RsMst_Depository!RiskType), "", RsMst_Depository!RiskType)
mExchangeTypeID = IIf(IsNull(RsMst_Depository!ExchangeTypeID), 0, RsMst_Depository!ExchangeTypeID)
TxtExchangeType = GetExchangeName(mExchangeTypeID)
TxtNAVDPName = IIf(IsNull(RsMst_Depository!NAVDepositoryName), "", RsMst_Depository!NAVDepositoryName)
''--- 10 Aug 2017 Ullhas Patil
TxtTANNo.Locked = True
TxtTANNo = IIf(IsNull(RsMst_Depository!TANNO), "", RsMst_Depository!TANNO)
''-- Ullhas Patil 20 Oct 2017
CmbStatus.Visible = False
TxtStatus.Text = IIf(IsNull(RsMst_Depository!Status), "Active", IIf(LCase(RsMst_Depository!Status) = "a", "Active", "De-Active"))
''----------------------------
TxtCreated = IIf(IsNull(RsMst_Depository!CreatedBy), "", RsMst_Depository!CreatedBy) & " :- " & IIf(IsNull(RsMst_Depository!CreatedDate), "", RsMst_Depository!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Depository!UpdatedBy), "", RsMst_Depository!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Depository!UpdatedDate), "", RsMst_Depository!UpdatedDate)
If RsMst_Depository.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtDepositoryID.Locked = False
TxtDepositoryName.Locked = False

TxtDepositoryID = ""
TxtDepositoryName = ""

TxtExchangeType.Visible = False
CmdExchangeType.Visible = True

CmbExchange.Visible = True
CmbExchange.Text = ""

TxtRiskType.Visible = False
CmbRiskType.Visible = True
CmbRiskType.Text = ""
''--- 10 Aug 2017 Ullhas Patil
TxtTANNo.Text = ""
TxtTANNo.Locked = False
''-- Ullhas Patil 20 Oct 2017
TxtStatus.Text = ""
CmbStatus.Text = ""
CmbStatus.Visible = True
''------------------
TxtCreated = ""
TxtUpdated = ""
TxtNAVDPName = ""
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
If TxtDepositoryID = "1" Then
  MsgBox "This DP Is Not Editable"
  Call GButtonLock(Me, True)
Else
  Call GButtonLock(Me, False)
End If
TxtDepositoryName.Locked = False
'TxtExchangeType.Visible = False
'CmdExchangeType.Visible = True
'CmbExchange.Visible = True
CmbExchange.Text = TxtExchangeType
TxtRiskType.Visible = False
CmbRiskType.Visible = True
CmbRiskType = TxtRiskType

TxtTANNo.Locked = False             ''-- 10 Aug 2017 Ullhas Patil
CmbStatus.Text = TxtStatus.Text     ''-- Ullhas Patil 20 Oct 2017
CmbStatus.Visible = True            ''-- Ullhas Patil 20 Oct 2017

End Sub
Private Function CheckForUseage() As Boolean
Dim Retval As Boolean

Retval = True
tmp = "select 'Mst_Client' As  'Type', Mst_Client.DepositoryID  from Mst_Client where Mst_Client.DepositoryID = '" & TxtDepositoryID & "' and Mst_Client.ExchangeTypeID = " & mExchangeTypeID & " "
If mExchangeTypeID = 2 Then
   tmp = tmp & " Union"
   tmp = tmp & " select 'Mst_RRN' As  'Type', Mst_RRN.DepositoryID  from Mst_RRN where Mst_RRN.DepositoryID = '" & TxtDepositoryID & "'"
End If
If Gen_DBType = "MDB" Then
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      Retval = False
   End If
Else
   Call Tmp4Table
   If TmpRs4.RecordCount > 0 Then
      Retval = False
   End If
End If
CheckForUseage = Retval
End Function

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
If CheckForUseage() = False Then
   MsgBox "You can't delete this record since it is used in Transaction !!!  "
   Exit Sub
End If
If TxtDepositoryID.Text = "1" And mExchangeTypeID = 1 Then
   MsgBox "You Can't Delete This Record !!!! "
   Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_Depository.Delete
   RsMst_Depository.Update
   If RsMst_Depository.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Depository.MoveNext
   If RsMst_Depository.EOF() Then
      RsMst_Depository.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Depository.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtDepositoryID.Text = "" Then
   MsgBox "Blank DepositoryID Not Allowed !!! "
   TxtDepositoryID.SetFocus
   Exit Sub
End If
If TxtDepositoryName.Text = "" Then
   MsgBox "Blank DepositoryName Not Allowed !!! "
   TxtDepositoryName.SetFocus
   Exit Sub
End If
If CmbRiskType.Text = "" Then
   MsgBox "Blank Risk Type Not Allowed !!! "
   CmbRiskType.SetFocus
   Exit Sub
End If
''--- 10 Aug 2017 Ullhas Patil
If TxtTANNo.Text = "" Then
   MsgBox "Blank TAN No Allowed !!! "
   TxtTANNo.SetFocus
   Exit Sub
Else
   If ValidateTANNo(UCase(TxtTANNo.Text)) = False Then
      MsgBox "Invalid TAN No !!!"
      TxtTANNo.SetFocus
      Exit Sub
   End If
End If
''----------------------------

''---------- 20 Oct 2017 Ullhas Patil
If CmbStatus.Text = "" Then
      MsgBox "Blank Status Not Allow."
      CmbStatus.SetFocus
      Exit Sub
End If

tmp = "ExchangeTypeID = " & mExchangeTypeID
RsMst_Depository.Filter = tmp
If Edit_Flg = "A" Then
    If RsMst_Depository.RecordCount > 0 Then
       RsMst_Depository.MoveFirst
       RsMst_Depository.Find "DepositoryID= '" & TxtDepositoryID.Text & "'"
       If Not RsMst_Depository.EOF Then
         MsgBox "Duplicate DepositoryID Not Allowed !!! "
          TxtDepositoryID.SetFocus
          Exit Sub
       End If
    End If
    ''-- Alter for duplicate TAN No 10 Aug 2017 Ullhas Patil
    RsMst_Depository.Filter = ""
    RsMst_Depository.MoveFirst
    RsMst_Depository.Find ("TANNO ='" & Trim(TxtTANNo.Text) & "'")
    If Not RsMst_Depository.EOF Then
      Select Case MsgBox("Duplicate TAN No Found with another Depository, Do you want to save the Depository with same TAN No?", vbYesNo)
          Case vbNo
              TxtTANNo.SetFocus
              Exit Sub
      End Select
    End If
    ''----------------------------
   RsMst_Depository.AddNew
   RsMst_Depository!G_UID = GetGUID
   'RsMst_Depository!DepositoryID = GetDepositoryID
   'TxtDepositoryID.Text = RsMst_Depository!DepositoryID
Else
    RsMst_Depository.MoveFirst
    RsMst_Depository.Find "DepositoryID= '" & TxtDepositoryID.Text & "'"
    If Not RsMst_Depository.EOF Then
       If RsMst_Depository!DepositoryID <> TxtDepositoryID.Text Then
          MsgBox "Duplicate DepositoryID Not Allowed !!! "
          TxtDepositoryID.SetFocus
          Exit Sub
       End If
    End If
   ''-- Alter for duplicate TAN No 09 Aug 2017 Ullhas Patil
   RsMst_Depository.Filter = ""
   RsMst_Depository.MoveFirst
   RsMst_Depository.Find ("TANNO ='" & Trim(TxtTANNo.Text) & "'")
   If Not RsMst_Depository.EOF Then
    If RsMst_Depository!DepositoryID <> TxtDepositoryID.Text Then
      Select Case MsgBox("Duplicate TAN No Found with another Depository, Do you want to save the Depository with same TAN No?", vbYesNo)
          Case vbNo
              TxtTANNo.SetFocus
              Exit Sub
      End Select
    End If
   End If
   ''----------------------------
    RsMst_Depository.MoveFirst
    RsMst_Depository.Find " DepositoryID= '" & TxtDepositoryID.Text & "'"
End If

RsMst_Depository!DepositoryID = TxtDepositoryID.Text
RsMst_Depository!DepositoryName = TxtDepositoryName.Text
RsMst_Depository!ExchangeTypeID = mExchangeTypeID
RsMst_Depository!RiskType = CmbRiskType

RsMst_Depository!TANNO = Trim(TxtTANNo)             ''-- Ullhas Patil 09 Sep 2017
RsMst_Depository!Status = Left(CmbStatus.Text, 1)   ''-- Ullhas Patil 20 Oct 2017

If IsNull(RsMst_Depository!CreatedBy) = True Or RsMst_Depository!CreatedBy = "" Then
   RsMst_Depository!CreatedBy = Gen_UserLoginID
   RsMst_Depository!CreatedDate = Now
Else
   RsMst_Depository!UpdatedBy = Gen_UserLoginID
   RsMst_Depository!UpdatedDate = Now
End If
RsMst_Depository.Update
RsMst_Depository.Filter = ""
RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID= '" & TxtDepositoryID.Text & "'"
Call Indata
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_Depository.MoveFirst
RsMst_Depository.Find "g_uid = '" & MSHFlexGrid1.TextMatrix(i, 3) & "'"
If Not RsMst_Depository.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtDepositoryID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
TxtDepositoryID = Trim(TxtDepositoryID)
If TxtDepositoryID = "1" Then
   MsgBox "Reserved ID,you can't use this ID !!!"
   TxtDepositoryID.SetFocus
End If
End Sub

Private Sub TxtDepositoryName_LostFocus()
TxtDepositoryName = Trim(TxtDepositoryName)
End Sub

Private Sub TxtTANNo_LostFocus()
If TxtTANNo.Locked = False Then
    If Trim(TxtTANNo.Text) <> "" Then
       TxtTANNo = UCase(TxtTANNo)
       If ValidateTANNo(TxtTANNo) = False Then
        MsgBox "Invalid TAN No"
        TxtTANNo.SetFocus
        Exit Sub
       End If
    End If
End If
End Sub
