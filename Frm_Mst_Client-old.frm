VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Client 
   Caption         =   "Client Master"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1605
      Left            =   11040
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2040
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
         TabIndex        =   59
         Top             =   240
         Value           =   -1  'True
         Width           =   735
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
         TabIndex        =   58
         Top             =   240
         Width           =   1350
      End
      Begin VB.OptionButton Option2 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4680
         TabIndex        =   57
         Top             =   240
         Width           =   1815
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
         TabIndex        =   56
         Top             =   240
         Width           =   1470
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   8040
         TabIndex        =   55
         Top             =   675
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchClient 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2865
         MaxLength       =   200
         TabIndex        =   54
         Top             =   675
         Visible         =   0   'False
         Width           =   5070
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   12405
         TabIndex        =   13
         Top             =   675
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   12
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
         TabIndex        =   61
         Top             =   240
         Width           =   2415
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
         TabIndex        =   60
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label6 
         Caption         =   "Client ID "
         Height          =   300
         Left            =   480
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   2280
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8085
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   14775
      Begin VB.TextBox TxtClientIDRow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Height          =   5010
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   14550
         Begin VB.CommandButton CmdAdd2List 
            Caption         =   "Add to List"
            Height          =   375
            Left            =   12840
            TabIndex        =   11
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox CmbExchange 
            Height          =   315
            Left            =   2040
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   210
            Width           =   2445
         End
         Begin VB.TextBox TxtClientName 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2040
            MaxLength       =   200
            TabIndex        =   9
            Top             =   975
            Width           =   6720
         End
         Begin VB.TextBox TxtClientID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   8
            Top             =   570
            Width           =   2445
         End
         Begin VB.ComboBox CmbDepositoryID 
            Height          =   315
            Left            =   2040
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   1365
            Width           =   6735
         End
         Begin VB.CommandButton CmdDepositoryID 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8880
            TabIndex        =   42
            Top             =   1365
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3000
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   14325
            _ExtentX        =   25268
            _ExtentY        =   5292
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label4 
            Caption         =   "Exchange Type"
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
            TabIndex        =   46
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label3 
            Caption         =   "Depository Name"
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
            TabIndex        =   45
            Top             =   1372
            Width           =   1800
         End
         Begin VB.Label Label2 
            Caption         =   "Client Name"
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
            Left            =   120
            TabIndex        =   44
            Top             =   1065
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "ClientID"
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
            Left            =   120
            TabIndex        =   43
            Top             =   660
            Width           =   1800
         End
      End
      Begin VB.TextBox TxtDepositorID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox TxtDepositorName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   200
         TabIndex        =   1
         Top             =   615
         Width           =   7440
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   2160
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1005
         Width           =   7440
      End
      Begin VB.TextBox TxtOrganization 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1830
         Width           =   7440
      End
      Begin VB.TextBox TxtEmailID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   200
         TabIndex        =   4
         Top             =   2220
         Width           =   7440
      End
      Begin VB.TextBox TxtTelephoneNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2595
         Width           =   3000
      End
      Begin VB.TextBox TxtMobileNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6600
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2610
         Width           =   3000
      End
      Begin VB.Label Label5 
         Caption         =   "Client ID Row"
         Height          =   300
         Left            =   7200
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label LblClientID 
         Caption         =   "DepositorID"
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblClientName 
         Caption         =   "Depositor Name         (Non-NCDEX)"
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
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblAddress 
         Caption         =   "Address"
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label LblOrganization 
         Caption         =   "Organization"
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   1845
         Width           =   1800
      End
      Begin VB.Label LblEmailID 
         Caption         =   "EmailID"
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   2220
         Width           =   1800
      End
      Begin VB.Label LblTelephoneNo 
         Caption         =   "Telephone No"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   2595
         Width           =   1800
      End
      Begin VB.Label LblMobileNo 
         Caption         =   "MobileNo"
         Height          =   300
         Left            =   5640
         TabIndex        =   18
         Top             =   2610
         Width           =   840
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   22
      Top             =   8160
      Width           =   14775
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   345
         TabIndex        =   41
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2220
         TabIndex        =   40
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4095
         TabIndex        =   39
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5970
         TabIndex        =   38
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   7845
         TabIndex        =   37
         Top             =   645
         Width           =   1875
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   345
         TabIndex        =   36
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2220
         TabIndex        =   35
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4080
         TabIndex        =   34
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5970
         TabIndex        =   33
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   7845
         TabIndex        =   32
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.TextBox TxtFlag 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox TxtExchangeTypeID 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2280
      TabIndex        =   29
      Top             =   1080
      Width           =   3600
   End
   Begin VB.ComboBox CmbExchangeTypeID 
      Height          =   315
      Left            =   2280
      TabIndex        =   30
      Top             =   1080
      Width           =   3600
   End
   Begin VB.CommandButton CmdExchangeTypeID 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   31
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtDepository 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   49
      Top             =   1080
      Width           =   7440
   End
   Begin VB.ComboBox CmbDepository 
      Height          =   315
      Left            =   2265
      TabIndex        =   48
      Top             =   1080
      Width           =   7455
   End
   Begin VB.CommandButton CmdDepository 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9960
      TabIndex        =   47
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label LblDepositoryID 
      Caption         =   "DepositoryName"
      Height          =   300
      Left            =   360
      TabIndex        =   50
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label LblExchangeTypeID 
      Caption         =   "Exchange Type"
      Height          =   300
      Left            =   360
      TabIndex        =   28
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label LblFlag 
      Caption         =   "Flag"
      Enabled         =   0   'False
      Height          =   300
      Left            =   720
      TabIndex        =   26
      Top             =   4200
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "FrmMst_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, i, InmClientRowID As Variant
Dim mExchangeID, mDepositoryID, mExchangeTypeID, mDepositorID, tmpExchangeID As Double
Dim tmpDepositorID As String
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
   MsgBox "Invalid  selection "
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

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_Client("Where DepositorID is null  ")
Call Grid_Head2
If RsMst_Client.RecordCount = 0 Then
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

Call InData
End Sub
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Mst_Client.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & "Mst_Client.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub



Private Sub MSHFlexGrid2_Click()
i = MSHFlexGrid2.RowSel
If i <= 0 Then Exit Sub
   CmbExchange.Enabled = True
'   If MSHFlexGrid2.TextMatrix(i, 8) <> "" Then
'      CmbExchange.Enabled = False
'   End If
   
   TxtClientID.Text = MSHFlexGrid2.TextMatrix(i, 2)
   TxtClientName.Text = MSHFlexGrid2.TextMatrix(i, 3)
   CmbExchange.Text = MSHFlexGrid2.TextMatrix(i, 1)
   mExchangeTypeID = GetExchangeID(CmbExchange.Text)
   CmbDepositoryID.Text = MSHFlexGrid2.TextMatrix(i, 4)
   mDepositoryID = GetDepositoryID(CmbDepositoryID.Text)
   CmdAdd2List.Caption = "Update List"
   
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Client.MoveNext
If RsMst_Client.EOF Then
   RsMst_Client.MoveLast
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub

Private Sub Option5_Click()
Call Option1_Click
End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Client.MovePrevious
If RsMst_Client.BOF Then
   RsMst_Client.MoveFirst
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_Client.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Client.MoveLast
Call InData
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
Call TableMst_Client
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Client.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 10
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200
'RsMst_Client.MoveFirst
Option4.Value = True
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_Client.RecordCount + 1

For x = 1 To RsMst_Client.RecordCount
     Call ShowGrid(x)
     RsMst_Client.MoveNext
Next

Label21.Refresh
Label21.Caption = RsMst_Client.RecordCount
'TxtSearchClient.Text = ""
End Sub
Private Sub TxtSearchClient_LostFocus()
Dim x As Integer
If TxtSearchClient = "" Then
   Call Grid_Head
   Label21.Caption = ""
   Exit Sub
End If
tmp = "Select * from Mst_Client Where"
If Option1.Value = True Then
   tmp = tmp & " ClientID  Like '%" & TxtSearchClient & "%'"  '
ElseIf Option2.Value = True Then
   tmp = "Select * from Mst_Client MC Inner Join  Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID where ME.ExchangeType like '%" & TxtSearchClient & "%'"
   'tmp = tmp & " DepositoryID Like '%" & TxtSearchClient & "%' "
ElseIf Option5.Value = True Then
   tmp = tmp & " ClientName Like '%" & TxtSearchClient & "%' "
End If
Call Tmp1Table
If TmpRs1.RecordCount = 0 Then
      MsgBox "No such Record Found!!!"
      Exit Sub
End If
MSHFlexGrid1.Rows = TmpRs1.RecordCount + 1
For x = 1 To TmpRs1.RecordCount
    Call ShowGrid(x)
    TmpRs1.MoveNext
Next
End Sub
Private Sub Option1_Click()
Label6.Visible = True
TxtSearchClient.Visible = True
TxtSearchClient = ""
CmdGo.Visible = True
If Option1.Value = True Then
   Label6.Caption = "Search By Client ID"
ElseIf Option2.Value = True Then
   Label6.Caption = "Search By Exchange Type"
ElseIf Option5.Value = True Then
   Label6.Caption = "Search By Client Name"
End If
TxtSearchClient.SetFocus
End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub

Private Sub Option3_Click()
Call Option1_Click
End Sub

Private Sub Option4_Click()
Dim x As Integer
Label6.Visible = False
TxtSearchClient.Visible = False
CmdGo.Visible = False
RsMst_Client.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_Client.RecordCount + 1
For x = 1 To RsMst_Client.RecordCount
    Call ShowGrid(x)
    RsMst_Client.MoveNext
Next
End Sub


'Private Sub ShowGrid(i As Integer)
'MSHFlexGrid1.ColWidth(0) = 1000
'
'For c = 0 To MSHFlexGrid1.Cols - 1
'   MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_Client.Fields(c)), "", RsMst_Client.Fields(c))
'Next
'   MSHFlexGrid1.TextMatrix(i, 3) = GetExchangeName(MSHFlexGrid1.TextMatrix(i, 3))
'   MSHFlexGrid1.TextMatrix(i, 4) = GetDepositorName(MSHFlexGrid1.TextMatrix(i, 4), "N")
'
'
'End Sub
Private Sub ShowGrid(i As Integer)
MSHFlexGrid1.ColWidth(0) = 1000
Label21.Refresh
If Option4.Value = False Then
   Call GridDetails(TmpRs1, i)
   Label21.Caption = TmpRs1.RecordCount
Else
   Call GridDetails(RsMst_Client, i)
   Label21.Caption = RsMst_Client.RecordCount
End If
End Sub
Private Sub GridDetails(RSType As Variant, iu As Integer)
For c = 0 To MSHFlexGrid1.Cols - 1
   MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(RSType.Fields(c)), "", RSType.Fields(c))
Next
   MSHFlexGrid1.TextMatrix(iu, 3) = GetExchangeName(MSHFlexGrid1.TextMatrix(iu, 3))
   MSHFlexGrid1.TextMatrix(iu, 4) = GetDepositorName(MSHFlexGrid1.TextMatrix(iu, 4), "N")

End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12
MSHFlexGrid1.TextMatrix(0, 0) = "ClientID"
MSHFlexGrid1.TextMatrix(0, 1) = "DepositorId"
MSHFlexGrid1.TextMatrix(0, 2) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 3) = "ExchangeType"
MSHFlexGrid1.TextMatrix(0, 4) = "DepositoryID"
MSHFlexGrid1.TextMatrix(0, 5) = "Flag"
MSHFlexGrid1.TextMatrix(0, 6) = "Address"
MSHFlexGrid1.TextMatrix(0, 7) = "Organization"
MSHFlexGrid1.TextMatrix(0, 8) = "Email"
MSHFlexGrid1.TextMatrix(0, 9) = "Tel No"
MSHFlexGrid1.TextMatrix(0, 10) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 11) = "Row ID"


MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 0
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 600
MSHFlexGrid1.ColWidth(6) = 2000
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 1600
MSHFlexGrid1.ColWidth(11) = 0
End Sub
Private Sub Grid_Head2()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 9
MSHFlexGrid2.TextMatrix(0, 0) = "Check Flag"
MSHFlexGrid2.TextMatrix(0, 1) = "ExchangeType"
MSHFlexGrid2.TextMatrix(0, 2) = "ClientID"
MSHFlexGrid2.TextMatrix(0, 3) = "Client Name"
MSHFlexGrid2.TextMatrix(0, 4) = "Depository Name"

MSHFlexGrid2.TextMatrix(0, 5) = "OldExchangeType"
MSHFlexGrid2.TextMatrix(0, 6) = "OldClientID"
MSHFlexGrid2.TextMatrix(0, 7) = "OldDepository Name"
MSHFlexGrid2.TextMatrix(0, 8) = "RowID"


MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(5) = 0
MSHFlexGrid2.ColWidth(6) = 0
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 0

MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 1600
MSHFlexGrid2.ColWidth(3) = 3400
MSHFlexGrid2.ColWidth(4) = 3200

End Sub
Public Sub InData()

TxtDepositorID.Locked = True
TxtClientID.Locked = True
TxtClientName.Locked = True

TxtDepositorName.Locked = True
CmbExchange.Visible = True
CmbExchange.Locked = True
CmbDepositoryID.Visible = True
CmbDepositoryID.Locked = True
CmdDepositoryID.Visible = False

TxtFlag.Locked = True
TxtAddress.Locked = True
TxtOrganization.Locked = True
TxtEmailId.Locked = True
TxtTelephoneNo.Locked = True
TxtMobileNo.Locked = True
CmdAdd2List.Enabled = False
MSHFlexGrid2.Enabled = False
'Call TableMst_Client("where DepositorID is Null")

TxtClientIDRow = IIf(IsNull(RsMst_Client!ClientIDRow), "", RsMst_Client!ClientIDRow)
InmClientRowID = TxtClientIDRow

TxtDepositorID = IIf(IsNull(RsMst_Client!ClientID), "", RsMst_Client!ClientID)
mDepositorID = TxtDepositorID.Text
TxtDepositorName = IIf(IsNull(RsMst_Client!ClientName), "", RsMst_Client!ClientName)

'mExchangeID = "1" 'IIf(IsNull(RsMst_Client!ExchangeTypeID), "", RsMst_Client!ExchangeTypeID)
'TxtExchangeTypeID = GetExchangeName(RsMst_Client!ExchangeTypeID) '(IsNull(RsMst_Client!ExchangeType), "", RsMst_Client!ExchangeType)

'mDepositoryID = "1" 'IIf(IsNull(RsMst_Client!DepositoryID), "", RsMst_Client!DepositoryID)
'TxtDepositoryID = GetDepositorName(RsMst_Client!DepositoryID, "N") '(IsNull(RsMst_Client!DepositoryName), "", RsMst_Client!DepositoryName)

TxtFlag = "0" 'IIf(IsNull(RsMst_Client!Flag), "", RsMst_Client!Flag)
TxtAddress = IIf(IsNull(RsMst_Client!Address), "", RsMst_Client!Address)
TxtOrganization = IIf(IsNull(RsMst_Client!Organization), "", RsMst_Client!Organization)
TxtEmailId = IIf(IsNull(RsMst_Client!EmailID), "", RsMst_Client!EmailID)
TxtTelephoneNo = IIf(IsNull(RsMst_Client!TelephoneNo), "", RsMst_Client!TelephoneNo)
TxtMobileNo = IIf(IsNull(RsMst_Client!MobileNo), "", RsMst_Client!MobileNo)


'RsMst_Client.Find ("DepositorID = '" & TxtDepositorID & "'")
'
'RsMst_Client.MoveFirst

tmp = "select * from Mst_Client where DepositorID = '" & TxtDepositorID & "'"

Call TmpTable
Call Grid_Head2
Dim i As Variant
i = 0
For j = 1 To TmpRs.RecordCount
      i = i + 1
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
      MSHFlexGrid2.TextMatrix(i, 0) = "*"
      MSHFlexGrid2.TextMatrix(i, 1) = GetExchangeName(TmpRs!ExchangeTypeID)
      MSHFlexGrid2.TextMatrix(i, 2) = TmpRs!ClientID 'GetUnitName(RsMst_Godown!UnitTypeID)
      MSHFlexGrid2.TextMatrix(i, 3) = TmpRs!ClientName 'GetCategoryName(RsMst_Godown!CategoryID)
      MSHFlexGrid2.TextMatrix(i, 4) = GetDepositoryName(TmpRs!DepositoryID, TmpRs!ExchangeTypeID)
      MSHFlexGrid2.TextMatrix(i, 5) = TmpRs!ExchangeTypeID
      MSHFlexGrid2.TextMatrix(i, 6) = TmpRs!ClientID
      MSHFlexGrid2.TextMatrix(i, 7) = TmpRs!DepositoryID
      MSHFlexGrid2.TextMatrix(i, 8) = TmpRs!ClientIDRow
      InmClientRowID = InmClientRowID & " , " & TmpRs!ClientIDRow
      TmpRs.MoveNext
Next

TmpRs.Close
'RsMst_Client.MoveFirst
If RsMst_Client.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call Grid_Head2
TxtDepositorName.Locked = False

TxtClientID.Locked = False
TxtClientName.Locked = False

CmbExchange.Locked = False
CmbDepositoryID.Locked = False
CmdDepositoryID.Visible = True

TxtFlag.Locked = False
TxtAddress.Locked = False
TxtOrganization.Locked = False
TxtEmailId.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False
CmdAdd2List.Enabled = True


TxtClientID = ""
TxtClientName = ""
TxtExchangeTypeID = ""
TxtDepositorID = ""
TxtFlag = ""
TxtAddress = ""
TxtOrganization = ""
TxtEmailId = ""
TxtTelephoneNo = ""
TxtMobileNo = ""
TxtDepositorName.Text = ""
CmbExchange.Text = ""
CmbExchange.Enabled = True
CmbDepositoryID.Text = ""
MSHFlexGrid2.Enabled = True
If RsMst_Client.RecordCount > 0 Then
   TxtDepositorName.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtClientID.Locked = False
TxtClientName.Locked = False
'TxtExchangeTypeID.Locked = False
CmbExchange.Visible = True
CmbExchange.Locked = False
'CmdExchangeTypeID.Visible = True
'TxtExchangeTypeID.Visible = False
'TxtDepositoryID.Locked = False
CmbDepositoryID.Visible = True
CmbDepositoryID.Locked = False
CmdDepositoryID.Visible = True
'TxtDepositoryID.Visible = False
TxtFlag.Locked = False
TxtAddress.Locked = False
TxtOrganization.Locked = False
TxtEmailId.Locked = False
TxtTelephoneNo.Locked = False
TxtMobileNo.Locked = False
TxtClientID.Locked = False
TxtClientName.Locked = False
CmdAdd2List.Enabled = True
MSHFlexGrid2.Enabled = True

'CmbExchangeTypeID.Text = TxtExchangeTypeID
'CmbDepositoryID.Text = TxtDepositoryID
End Sub
Private Function CheckForUseage() As Boolean
Dim RetVal As Boolean

RetVal = True

tmp = "select 'GSL' As  'Type', Mst_GSL.ClientIDRow  from Mst_GSL where Mst_GSL.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CDC' As 'Type',Txn_CDC.ClientIDRow  from Txn_CDC where Txn_CDC.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Mst_StorageContract' As 'Type',Mst_StorageContract.ClientIDRow  from Mst_StorageContract where Mst_StorageContract.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CxTF_ClientDetail' As 'Type',Txn_CxTF_ClientDetail.ClientIDRow  from Txn_CxTF_ClientDetail where Txn_CxTF_ClientDetail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_WHR_Detail' As 'Type',Txn_WHR_Detail.ClientIDRow  from Txn_WHR_Detail where Txn_WHR_Detail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_ISINRevalidation' As 'Type',Txn_ISINRevalidation.ClientRowID  from Txn_ISINRevalidation where Txn_ISINRevalidation.ClientRowID in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CxTF_GSL' As 'Type',Txn_CxTF_GSL.ClientIDRow  from Txn_CxTF_GSL where Txn_CxTF_GSL.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CxTF_ClientDetail' As 'Type',Txn_CxTF_ClientDetail.ClientIDRow  from Txn_CxTF_ClientDetail where Txn_CxTF_ClientDetail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CDC_GSL_Detail' As 'Type',Txn_CDC_GSL_Detail.ClientIDRow  from Txn_CDC_GSL_Detail where Txn_CDC_GSL_Detail.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CDC_GSL_Detail' As 'Type',Txn_CDC_GSL_Detail.ClientIDRowSeller  from Txn_CDC_GSL_Detail where Txn_CDC_GSL_Detail.ClientIDRowSeller in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Mst_RRN' As 'Type',Mst_RRN.ClientIDRow  from Mst_RRN where Mst_RRN.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_PledgeCreation' As 'Type',Txn_PledgeCreation.ClientIDRow  from Txn_PledgeCreation where Txn_PledgeCreation.ClientIDRow in( " & InmClientRowID & ") "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_CashReceipt' As 'Type',Txn_CashReceipt.ClientIDRow  from Txn_CashReceipt where Txn_CashReceipt.ClientIDRow in( " & InmClientRowID & ") "

If Gen_DBType = "MDB" Then
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      RetVal = False
   End If
Else
   Call Tmp4Table
   If TmpRs4.RecordCount > 0 Then
      RetVal = False
   End If
End If
CheckForUseage = RetVal
End Function


Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

If CheckForUseage() = False Then
   MsgBox "You can't delete this record since it is used in Transaction !!!  "
   Exit Sub
End If


ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)


If ans = vbYes Then
   
   tmp = "Delete from Mst_Client where DepositorID = '" & TxtDepositorID & "'"
   Call TmpTable
   RsMst_Client.Delete
   RsMst_Client.Update
   If RsMst_Client.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Client.MoveNext
   If RsMst_Client.EOF() Then
      RsMst_Client.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_Client.CancelUpdate
End Sub
Function CheckDuplicateClient() As Boolean
Dim RetVal As Boolean
RetVal = False

For j = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(j, 2) = TxtClientID And _
       MSHFlexGrid2.TextMatrix(j, 1) = CmbExchange.Text And _
       MSHFlexGrid2.TextMatrix(j, 4) = CmbDepositoryID.Text Then
       If CmdAdd2List.Caption = "Add to List" Then
          RetVal = True
          Exit For
       Else
          If j <> MSHFlexGrid2.RowSel Then
             RetVal = True
             Exit For
          End If
       End If
    End If
Next
CheckDuplicateClient = RetVal
End Function
Private Sub CmdAdd2List_Click()

If TxtClientID.Text = "" Then
   MsgBox "Blank ClientID Not Allowed !!! "
   TxtClientID.SetFocus
   Exit Sub
End If
If TxtClientName.Text = "" Then
   MsgBox "Blank ClientName Not Allowed !!! "
   TxtClientName.SetFocus
   Exit Sub
End If

If CmbExchange.Text = "" Then
   MsgBox "Blank Exchange Type Not Allowed !!! "
   CmbExchange.SetFocus
   Exit Sub
End If

If CmbDepositoryID.Text = "" Then
   MsgBox "Blank Depository Type Not Allowed!!!"
   CmbDepositoryID.SetFocus
   Exit Sub
End If


If CheckDuplicateClient() = True Then
   MsgBox "ClientID already Exist !!!!"
   TxtClientID.SetFocus
   Exit Sub
End If
tmp = "select * from Mst_Client where ClientID = '" & TxtClientID & "' and ExchangeTypeID = " & mExchangeTypeID & " and DepositoryID = '" & mDepositoryID & "'"
Call TmpTable
If TmpRs.RecordCount <> 0 Then
   If TmpRs!DepositorId <> TxtDepositorID Then
      MsgBox "ClientID already Exist (Other) !!!! Depositor ID : " & TmpRs!DepositorId & " : " & TmpRs!ClientID & " - Name : " & TmpRs!ClientName
      TxtClientID.SetFocus
      Exit Sub
   End If
End If

If MSHFlexGrid2.Rows > 1 Then
   If CmdAdd2List.Caption = "Update List" Then
      i = MSHFlexGrid2.RowSel
      MSHFlexGrid2.TextMatrix(i, 0) = "*"
   Else
      i = MSHFlexGrid2.Rows - 1
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
      MSHFlexGrid2.TextMatrix(i, 0) = ""
   End If
   MSHFlexGrid2.TextMatrix(i, 1) = CmbExchange.Text
   MSHFlexGrid2.TextMatrix(i, 2) = TxtClientID.Text
   MSHFlexGrid2.TextMatrix(i, 3) = TxtClientName.Text
   MSHFlexGrid2.TextMatrix(i, 4) = CmbDepositoryID.Text
End If

TxtClientID.Text = ""
TxtClientName.Text = ""
CmbExchange.Text = ""
CmbDepositoryID.Text = ""
CmdAdd2List.Caption = "Add to List"
CmbExchange.Enabled = True

TmpRs.Close
   
End Sub
Private Sub SaveCmd_Click()

If TxtDepositorName.Text = "" Then
   MsgBox "Blank DepositorName Not Allowed !!! "
   TxtDepositorName.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsMst_Client.RecordCount > 0 Then
      RsMst_Client.MoveFirst
      RsMst_Client.Find "ClientName = '" & TxtDepositorName.Text & "'"
      If Not RsMst_Client.EOF Then
        MsgBox "Duplicate DepositorName Not Allowed !!!"
         TxtDepositorName.SetFocus
         Exit Sub
      End If
   End If
   RsMst_Client.AddNew
   RsMst_Client!ClientID = GetDepositorID
   TxtDepositorID.Text = RsMst_Client!ClientID
   RsMst_Client!ClientIDRow = GetRowID
   RsMst_Client!g_uid = GetGUID
Else
    RsMst_Client.MoveFirst
    RsMst_Client.Find "ClientName = '" & TxtDepositorName.Text & "'"
    If Not RsMst_Client.EOF Then
       If RsMst_Client!ClientIDRow <> TxtClientIDRow.Text Then
          MsgBox "Duplicate Depositor Name Not Allowed !!! "
          TxtDepositorName.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Client.MoveFirst
    RsMst_Client.Find "ClientIDRow = " & TxtClientIDRow.Text
    If RsMst_Client.EOF Then
       MsgBox "No Such Record Found!!! "
       Exit Sub
    End If
End If

RsMst_Client!ClientID = TxtDepositorID.Text
RsMst_Client!ClientName = TxtDepositorName.Text
RsMst_Client!ExchangeTypeID = "1" 'mExchangeTypeID
RsMst_Client!DepositoryID = "1" 'mDepositoryID
RsMst_Client!Flag = "0" 'TxtFlag.Text
RsMst_Client!Address = TxtAddress.Text
RsMst_Client!Organization = TxtOrganization.Text
RsMst_Client!EmailID = TxtEmailId.Text
RsMst_Client!TelephoneNo = TxtTelephoneNo.Text
RsMst_Client!MobileNo = TxtMobileNo.Text

'RsMst_Client!DepositorID = ""
RsMst_Client.Update

RsMst_Client.Requery
Dim om_ClientIDRow As Variant
om_ClientIDRow = RsMst_Client!ClientIDRow
RsMst_Client.Find "ClientIDRow = " & om_ClientIDRow
Call SaveGridData
RsMst_Client.Requery
RsMst_Client.MoveFirst

RsMst_Client.Find "ClientID = " & TxtDepositorID.Text
If RsMst_Client.EOF Then
   RsMst_Client.MoveFirst
End If
Call InData
Call ClearClientData
Call GButtonLock(Me, True)
End Sub
Private Sub ClearClientData()
   TxtClientID.Text = ""
   TxtClientName.Text = ""
   CmbExchange.Text = ""
   CmbDepository.Text = ""
End Sub

Private Sub SaveGridData()
tmp = "select * from Mst_Client where DepositorID = '" & TxtDepositorID & "'"
Call TmpTable

If MSHFlexGrid2.Rows > 1 Then
   For i = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(i, 1) = "" Then Exit For
       
       'tmp = "ClientID = '" & TxtClientID & "' and ExchangeTypeID = " & tmpExchangeID & " and DepositoryID = '" & tmpDepositorID & "'"
       tmp = "ClientID = '" & MSHFlexGrid2.TextMatrix(i, 6) & "' and ExchangeTypeID = " & Val(MSHFlexGrid2.TextMatrix(i, 5)) & " and DepositoryID = '" & MSHFlexGrid2.TextMatrix(i, 7) & "'"
       TmpRs.Filter = tmp
       If TmpRs.EOF Then
          TmpRs.AddNew
          TmpRs!ClientIDRow = GetRowID
          TmpRs!g_uid = GetGUID
       End If
       TmpRs!ClientID = MSHFlexGrid2.TextMatrix(i, 2)
       TmpRs!ClientName = MSHFlexGrid2.TextMatrix(i, 3) 'GetUnitID(MSHFlexGrid2.TextMatrix(i, 1))
       TmpRs!DepositoryID = GetDepositoryID(MSHFlexGrid2.TextMatrix(i, 4)) 'GetCategoryID(MSHFlexGrid2.TextMatrix(i, 2))
       TmpRs!ExchangeTypeID = GetExchangeID(MSHFlexGrid2.TextMatrix(i, 1))
       TmpRs!DepositorId = TxtDepositorID.Text
       TmpRs.Update
       TmpRs.Requery
       
    Next i
End If
TmpRs.Close
'Call GButtonLock(Me, True)
''RsMst_Client.Requery
''RsMst_Client.MoveFirst
'''RsMst_Client.Find "CMPID = " & TxtCMPID & ""
''If RsMst_CMP.EOF Then
''   RsMst_CMP.MoveLast
''End If
'
'Call InData

End Sub
Private Function GetDepositorID() As Double
Dim RetVal As Double
If Gen_DBType = "MDB" Then
    tmp = "Select max(val(ClientID)) from Mst_Client where DepositorId is Null"
Else
    tmp = "Select max(convert(decimal(10,0),ClientID)) from Mst_Client where DepositorId is Null"
End If
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

RetVal = RetVal + 1
GetDepositorID = RetVal
End Function
Private Function GetRowID() As Double
Dim RetVal As Double
tmp = "Select max(ClientIDRow) from Mst_Client "
Call Tmp1Table
RetVal = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
RetVal = RetVal + 1
TmpRs1.Close
GetRowID = RetVal
End Function
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_Client.MoveFirst
If MSHFlexGrid1.TextMatrix(i, 1) <> "" Then
 RsMst_Client.Find "ClientID = '" & MSHFlexGrid1.TextMatrix(i, 1) & "'"
Else
 RsMst_Client.Find "ClientID = '" & MSHFlexGrid1.TextMatrix(i, 0) & "'"
End If
If Not RsMst_Client.EOF Then
   Call InData
   Frame0.Visible = False
End If

Call TableMst_Client("Where DepositorID is null")
End Sub
Private Sub CmdExchangeTypeID_Click()
    FrmMst_ExchangeType.Show
End Sub


Private Sub CmdDepositoryID_Click()
    FrmMst_Depository.Show
End Sub

Private Sub CmbDepositoryID_GotFocus()
If CmbExchange.Text = "" Then Exit Sub
tmp = CmbDepositoryID.Text
Call TableMst_Depository(" Where ExchangeTypeID = " & mExchangeTypeID)
CmbDepositoryID.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Depository.RecordCount
  If UCase(CmbExchange.Text) = "SPOT" Or UCase(CmbExchange.Text) = "NCDEX" Then
    If RsMst_Depository!DepositoryName <> "NCMSL" Then
      CmbDepositoryID.AddItem RsMst_Depository!DepositoryName
    End If
  ElseIf UCase(CmbExchange.Text) = "NON NCDEX" Then
    If RsMst_Depository!DepositoryName = "NCMSL" Then
      CmbDepositoryID.AddItem RsMst_Depository!DepositoryName
    End If
     'CmbDepositoryID.AddItem RsMst_Depository!DepositoryName
  End If
  RsMst_Depository.MoveNext
Next
CmbDepositoryID.Text = tmp
End Sub
Private Sub CmbDepositoryID_LostFocus()
If CmbDepositoryID.Text = "" Then
   Exit Sub
End If
RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryName = '" & CmbDepositoryID.Text & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid  selection "
   CmbDepositoryID.SetFocus
   Exit Sub
ElseIf UCase(CmbExchange.Text) = "SPOT" Or UCase(CmbExchange.Text) = "NCDEX" Then
   If CmbDepositoryID.Text = "NCMSL" Then
    MsgBox "Invalid  selection "
    CmbDepositoryID.SetFocus
    Exit Sub
   End If
ElseIf UCase(CmbExchange.Text) = "NON NCDEX" Then
   If CmbDepositoryID.Text <> "NCMSL" Then
    MsgBox "Invalid  selection "
    CmbDepositoryID.SetFocus
    Exit Sub
   End If
End If
mDepositoryID = RsMst_Depository!DepositoryID
End Sub

Private Sub TxtClientID_LostFocus()
TxtClientID = Trim(TxtClientID)
End Sub


Private Sub TxtClientName_LostFocus()
TxtClientName = Trim(TxtClientName)
End Sub

Private Sub TxtDepositorName_LostFocus()
TxtDepositorName = Trim(TxtDepositorName)
End Sub

Private Sub TxtSearchClient12_LostFocus()
Dim x As Integer
RsMst_Client.MoveFirst
Call Grid_Head
If TxtSearchClient <> "" Then
    RsMst_Client.Find ("ClientID = '" & Trim(TxtSearchClient) & "'")
    If RsMst_Client.EOF Then
      MsgBox "No Such Client Exist in DataBase!!!"
      TxtSearchClient.SetFocus
      Exit Sub
    End If
  Call TableMst_Client("Where ClientID = '" & Trim(TxtSearchClient) & "'")
  MSHFlexGrid1.Rows = RsMst_Client.RecordCount + 1
  For x = 1 To RsMst_Client.RecordCount
     Call ShowGrid(x)
     RsMst_Client.MoveNext
  Next
Else
 'RsMst_ISIN.Sort
 Call TableMst_Client
 MSHFlexGrid1.Rows = RsMst_Client.RecordCount + 1

 For x = 1 To RsMst_Client.RecordCount
     Call ShowGrid(x)
     RsMst_Client.MoveNext
 Next

End If
End Sub
