VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmUsersMaster 
   Caption         =   "Users  Master"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   10860
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1575
      Left            =   -240
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   15015
      Begin VB.ComboBox CmbSFlag 
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
         ItemData        =   "FrmUsersMaster.frx":0000
         Left            =   6495
         List            =   "FrmUsersMaster.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   405
         Width           =   2550
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   10680
         TabIndex        =   45
         Top             =   180
         Width           =   1005
      End
      Begin VB.ComboBox CmbSUserRole 
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
         Left            =   3870
         Sorted          =   -1  'True
         TabIndex        =   43
         Top             =   435
         Width           =   2550
      End
      Begin VB.TextBox TxtSearchUser 
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
         Left            =   165
         TabIndex        =   41
         Top             =   435
         Width           =   3645
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   9630
         TabIndex        =   40
         Top             =   180
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1620
         Left            =   75
         TabIndex        =   12
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2858
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label15 
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
         Height          =   225
         Left            =   7365
         TabIndex        =   51
         Top             =   150
         Width           =   930
      End
      Begin VB.Label Label13 
         Caption         =   "User Role"
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
         Left            =   4740
         TabIndex        =   44
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label19 
         Caption         =   "Search By User Name"
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
         Left            =   1065
         TabIndex        =   42
         Top             =   165
         Width           =   2895
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
         Left            =   11880
         TabIndex        =   39
         Top             =   240
         Width           =   2340
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
         Left            =   14280
         TabIndex        =   38
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   7125
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3015
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4470
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5805
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1560
         TabIndex        =   30
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2895
         TabIndex        =   29
         Top             =   615
         Width           =   1575
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   4470
         TabIndex        =   28
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5805
         TabIndex        =   27
         Top             =   615
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4605
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7335
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2955
         Width           =   3735
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox TxtEmailID 
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
         IMEMode         =   3  'DISABLE
         Left            =   1695
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1635
         Width           =   5295
      End
      Begin VB.TextBox TxtContact 
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
         IMEMode         =   3  'DISABLE
         Left            =   10560
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
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
         Height          =   1215
         Left            =   10560
         MaxLength       =   255
         TabIndex        =   6
         Top             =   3120
         Width           =   5295
      End
      Begin VB.TextBox TxtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   10560
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox TxtEmpId 
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtUsers 
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
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox TxtUsersID 
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CmbLocation 
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
         Left            =   10560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox TxtLocation 
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
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1680
         Width           =   5295
      End
      Begin VB.ComboBox CmbUserRole 
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1170
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.TextBox TxtUserRole 
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
         TabIndex        =   17
         Top             =   1170
         Width           =   5295
      End
      Begin VB.TextBox TxtFlag 
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2070
         Width           =   2265
      End
      Begin VB.ComboBox CmbFlag 
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
         ItemData        =   "FrmUsersMaster.frx":0021
         Left            =   1665
         List            =   "FrmUsersMaster.frx":002B
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   2070
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label14 
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
         Height          =   285
         Left            =   240
         TabIndex        =   46
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Email ID"
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
         Left            =   270
         TabIndex        =   37
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label Label11 
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
         Height          =   270
         Left            =   240
         TabIndex        =   36
         Top             =   2565
         Width           =   1695
      End
      Begin VB.Label Label10 
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
         Left            =   210
         TabIndex        =   35
         Top             =   3015
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   8880
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Address"
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
         TabIndex        =   23
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Password"
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
         Left            =   9000
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label6 
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
         Left            =   8880
         TabIndex        =   21
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Login ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         TabIndex        =   19
         Top             =   255
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Users Role"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Users   Id"
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmUsersMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mUserRoleID, mLocationID, mPassword As Variant
Dim mSFlag, mFlag As String
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtUsers.Locked = False
TxtAddress.Locked = False
TxtContact.Locked = False
TxtEmpID.Locked = False
TxtPassword.Locked = False
TxtEmailId.Locked = False

TxtUsersID = ""
TxtUsers = ""
TxtAddress = ""
TxtContact = ""
TxtEmpID = ""
TxtPassword = ""
TxtEmailId = ""
CmbUserRole.Text = ""
CmbLocation.Text = ""
TxtCreated = ""
TxtUpdated = ""

If LCase(UserType) = "power" Or LCase(UserType) = "superadmin" Or LCase(UserType) = "admin" Or LCase(UserType) = "headcmg" Then
   CmbUserRole.Enabled = True
Else
   CmbUserRole.Enabled = False
   CmbUserRole.Text = "Account"
   Call CmbUserRole_GotFocus
   Call CmbUserRole_LostFocus
End If

CmbUserRole.Visible = True
TxtUserRole.Visible = False

TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = ""

CmbLocation.Visible = True
TxtLocation.Visible = False

If RsMst_Users.RecordCount > 0 Then
   TxtEmpID.SetFocus
End If

End Sub

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub
If CmbFlag.Text = "Active" Then
   mFlag = "Y"
ElseIf CmbFlag.Text = "De-Active" Then
   mFlag = "N"
Else
   MsgBox "Invalid Selection!!!"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
Call TableMst_Location("")
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
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
RsMst_Location.Find "Locationname = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location.Fields(0)

End Sub

Private Sub CmbSUserRole_GotFocus()
tmp = CmbSUserRole.Text
Call TableMst_UserRole
CmbSUserRole.Clear

If RsMst_UserRole.RecordCount = 0 Then
   Exit Sub
End If
For I = 1 To RsMst_UserRole.RecordCount
    CmbSUserRole.AddItem RsMst_UserRole.Fields(1)
    RsMst_UserRole.MoveNext
Next
CmbSUserRole.Text = tmp
End Sub

Private Sub CmbSUserRole_LostFocus()
If CmbSUserRole.Text = "" Then
   Exit Sub
End If
RsMst_UserRole.MoveFirst
RsMst_UserRole.Find "GroupName = '" & CmbSUserRole.Text & "'"

If RsMst_UserRole.EOF Then
   MsgBox "Invalid User Role "
   CmbSUserRole.SetFocus
   Exit Sub
End If
tmp = CmbSUserRole.Text
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_UserMaster.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_UserMaster.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub DeleteCmd_Click()
If UserType <> "Power" Then
   MsgBox "Access denied!!!!! "
   Exit Sub
End If


Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this Record???", vbYesNo)
If ans = vbYes Then
   RsMst_Users.Delete
   RsMst_Users.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_Users.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Users.MoveNext
   If RsMst_Users.EOF() Then
      RsMst_Users.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Users.CancelUpdate

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
''TxtUsers.Locked = False
CmbUserRole.Visible = True
TxtUserRole.Visible = False
CmbUserRole.Text = TxtUserRole

'If UserType = "Power" Then
If LCase(UserType) = "power" Or LCase(UserType) = "superadmin" Or LCase(UserType) = "admin" Or LCase(UserType) = "headcmg" Then
   CmbUserRole.Enabled = True
Else
   CmbUserRole.Enabled = False
End If


If LCase(UserType) = "headcmg" Then
   If LCase(CmbUserRole.Text) = "power" Then
      CmbUserRole.Enabled = False
   End If
ElseIf LCase(UserType) = "superadmin" Then
   If LCase(CmbUserRole.Text) = "power" Or LCase(CmbUserRole.Text) = "headcmg" Then
      CmbUserRole.Enabled = False
   End If
ElseIf LCase(UserType) = "admin" Then
   If LCase(CmbUserRole.Text) = "power" Or LCase(CmbUserRole.Text) = "headcmg" Or LCase(CmbUserRole.Text) = "superadmin" Then
      CmbUserRole.Enabled = False
   End If
End If

'CmbLocation.Visible = True
'TxtLocation.Visible = False
CmbLocation.Text = TxtLocation

'TxtFlag.Visible = False
'CmbFlag.Visible = True
'CmbFlag.Text = TxtFlag

''TxtAddress.Locked = False
''TxtContact.Locked = False
''TxtEmpID.Locked = False
''TxtEmailID.Locked = False
''If CmbUserRole.Text = "Account" Then
''   TxtPassword.Locked = False
''Else
''   TxtPassword.Locked = True
''End If
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

RsMst_Users.MoveFirst
TxtSearchUser = ""
CmbSUserRole.Text = ""
CmbSFlag.Text = ""
Label22.Caption = ""
'Call StatusBarData(Me)

Call TableMst_Users
If RsMst_Users.RecordCount = 0 Then
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()

CmbUserRole.Visible = False
TxtUserRole.Visible = True

CmbLocation.Visible = False
TxtLocation.Visible = True

TxtUsers.Locked = True
TxtAddress.Locked = True
TxtContact.Locked = True
TxtEmpID.Locked = True
TxtPassword.Locked = True
TxtEmailId.Locked = True
''
TxtFlag.Visible = True
CmbFlag.Visible = False
CmbFlag.Text = ""

tmp = "select * from Mst_UserRole where UserGroupID = " & RsMst_Users!UserGRoupID & ""
Call TmpTable
TxtUsersID = RsMst_Users!UsersID
TxtUsers = RsMst_Users!EmployeeName
TxtUserRole = IIf(IsNull(TmpRs.Fields(1)), "", TmpRs.Fields(1))
TxtAddress = ""
If RsMst_Users!Address <> "" Then
   TxtAddress = RsMst_Users!Address
End If

mLocationID = RsMst_Users!LocationID

TxtContact = IIf(IsNull(RsMst_Users!ContactNo), "", RsMst_Users!ContactNo)
TxtEmpID = RsMst_Users!EmployeeNo
mPassword = RsMst_Users!Password
TxtPassword = RsMst_Users!Password
TxtUsers = RsMst_Users!EmployeeName
TxtEmailId = IIf(IsNull(RsMst_Users!EmailID), "", RsMst_Users!EmailID)
TxtCreated = IIf(IsNull(RsMst_Users!CreatedBy), "", RsMst_Users!CreatedBy) & " :- " & IIf(IsNull(RsMst_Users!CreatedDate), "", RsMst_Users!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Users!UpdatedBy), "", RsMst_Users!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Users!UpdatedDate), "", RsMst_Users!UpdatedDate)

mUserRoleID = TmpRs.Fields(0)
TxtLocation = ""
If mLocationID <> "" Then
   tmp = "select * from Mst_Location where LocationId = " & mLocationID & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      TxtLocation = IIf(IsNull(TmpRs.Fields(1)), "", TmpRs.Fields(1))
   End If
End If
mFlag = IIf(IsNull(RsMst_Users!Flag), "", RsMst_Users!Flag)
If mFlag = "Y" Then
   TxtFlag = "Active"
ElseIf mFlag = "N" Then
   TxtFlag = "De-Active"
Else
   TxtFlag = ""
End If
Call GButtonLock(Me, True)
End Sub

Private Sub NextCmd_Click()
RsMst_Users.MoveNext
If RsMst_Users.EOF Then
   RsMst_Users.MoveLast
End If
Call Indata
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_Users.MoveFirst
Call Indata

End Sub

Private Sub LastCmd_Click()
RsMst_Users.MoveLast
Call Indata

End Sub

Private Sub PreviousCmd_Click()
RsMst_Users.MovePrevious
If RsMst_Users.BOF Then
   RsMst_Users.MoveFirst
End If
Call Indata

End Sub


Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_Users.MoveFirst

If RsMst_Users.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_Users.MoveFirst
RsMst_Users.Find "UsersID = " & Val(mCaId) & ""
If RsMst_Users.EOF Then
   RsMst_Users.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtUsers = "" Then
   MsgBox "Blank User Name not allowed"
   Exit Sub
End If

If CmbUserRole.Text = "" Then
   MsgBox "Blank User Role not allowed"
   Exit Sub
End If

'If CmbUserRole.Text = "Power" Then
'   If UserType <> CmbUserRole.Text Then
'      MsgBox "Access Denie !!!!! "
'      Exit Sub
'   End If
'End If

If TxtEmpID.Text = "" Then
   MsgBox "Blank Employee ID not allowed"
   Exit Sub
End If
''If TxtPassword = "" Then
''   MsgBox "Blank Password not allowed"
''   'TxtPassword.SetFocus
''   Exit Sub
''End If

''If mLocationID = "" Then
''   MsgBox "Blank Location not allowed"
''   'CmbLocation.SetFocus
''End If

If TxtEmailId = "" Then
   MsgBox "Blank EmailID not allowed"
   TxtEmailId.SetFocus
   Exit Sub
End If

If ChkEmail = False Then
   MsgBox "Please provide proper EmailID "
   TxtEmailId.SetFocus
   Exit Sub
End If

''If CmbFlag.Text = "" Then
''   MsgBox "Blank Status not allowed"
''   CmbFlag.SetFocus
''   Exit Sub
''End If

If Edit_Flg = "A" Then
   If RsMst_Users.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_Users.MoveFirst
        RsMst_Users.Find "EmployeeNo = '" & TxtEmpID & "'"
        If Not RsMst_Users.EOF Then
           MsgBox "Duplicate Employee not allowed"
           Exit Sub
        End If
        RsMst_Users.MoveLast
        tmp = RsMst_Users!UsersID + 1
    End If
    RsMst_Users.AddNew
    RsMst_Users!UsersID = tmp
    RsMst_Users!G_UID = GetGUID
    TxtUsersID = RsMst_Users!UsersID
Else
    RsMst_Users.MoveFirst
    RsMst_Users.Find "EmployeeNo = '" & TxtEmpID & "'"
    If Not RsMst_Users.EOF Then
       If RsMst_Users!UsersID <> TxtUsersID Then
          MsgBox "Duplicate Employee not allowed"
          Exit Sub
       End If
    End If
   RsMst_Users.MoveFirst
   RsMst_Users.Find "UsersID = " & TxtUsersID & ""
End If

RsMst_Users!Address = TxtAddress
RsMst_Users!ContactNo = TxtContact
RsMst_Users!EmployeeNo = TxtEmpID
RsMst_Users!Password = IIf(TxtPassword.Text = "", "ncmsl@123", TxtPassword.Text)
RsMst_Users!EmployeeName = TxtUsers
RsMst_Users!UserGRoupID = mUserRoleID
RsMst_Users!LocationID = mLocationID
RsMst_Users!EmailID = TxtEmailId
'RsMst_Users!Flag = mFlag

If Edit_Flg = "A" Then
   RsMst_Users!PassExpiryDate = Date + 14
End If
RsMst_Users!BatchID = Null

If IsNull(RsMst_Users!CreatedBy) = True Or RsMst_Users!CreatedBy = "" Then
   RsMst_Users!CreatedBy = Gen_UserLoginID
   RsMst_Users!CreatedDate = Now
Else
   RsMst_Users!UpdatedBy = Gen_UserLoginID
   RsMst_Users!UpdatedDate = Now
End If

RsMst_Users.Update
RsMst_Users.Requery
RsMst_Users.MoveFirst
RsMst_Users.Find "UsersID = " & TxtUsersID & ""

Call GButtonLock(Me, True)
TxtUsers.Locked = True
Call Indata
End Sub
'---------------
Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If Edit_Flg = "A" Then
      RsMst_Users.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 500
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 700

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head
'
'RsMst_Users.MoveFirst
'TxtSearchUser = ""
'CmbSUserRole.Text = ""
'CmbSFlag.Text = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5

MSHFlexGrid1.TextMatrix(0, 0) = "User Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Login Id"
MSHFlexGrid1.TextMatrix(0, 2) = "User"
MSHFlexGrid1.TextMatrix(0, 3) = "User Role"
MSHFlexGrid1.TextMatrix(0, 4) = "Status"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 4500
MSHFlexGrid1.ColWidth(3) = 2500


End Sub


Private Sub CmbUserRole_GotFocus()
tmp = CmbUserRole.Text
Call TableMst_UserRole
CmbUserRole.Clear

If RsMst_UserRole.RecordCount = 0 Then
   Exit Sub
End If

For I = 1 To RsMst_UserRole.RecordCount
    If LCase(UserType) = "power" Then
       CmbUserRole.AddItem RsMst_UserRole.Fields(1)
    ElseIf LCase(UserType) = "headcmg" Then
       If LCase(RsMst_UserRole.Fields(1)) <> "power" Then
          CmbUserRole.AddItem RsMst_UserRole.Fields(1)
       End If
    ElseIf LCase(UserType) = "superadmin" Then
       If LCase(RsMst_UserRole.Fields(1)) <> "power" And LCase(RsMst_UserRole.Fields(1)) <> "headcmg" Then
          CmbUserRole.AddItem RsMst_UserRole.Fields(1)
       End If
    ElseIf LCase(UserType) = "admin" Then
       If LCase(RsMst_UserRole.Fields(1)) <> "power" And LCase(RsMst_UserRole.Fields(1)) <> "headcmg" And LCase(RsMst_UserRole.Fields(1)) <> "superadmin" Then
          CmbUserRole.AddItem RsMst_UserRole.Fields(1)
       End If
    End If
    RsMst_UserRole.MoveNext
Next

CmbUserRole.Text = tmp

End Sub
Private Sub CmbUserRole_LostFocus()
If CmbUserRole.Text = "" Then
   Exit Sub
End If
RsMst_UserRole.MoveFirst
RsMst_UserRole.Find "GroupName = '" & CmbUserRole.Text & "'"

If RsMst_UserRole.EOF Then
   MsgBox "Invalid User Role "
   CmbUserRole.SetFocus
   Exit Sub
End If
mUserRoleID = RsMst_UserRole.Fields(0)
tmp = CmbUserRole.Text
End Sub


Private Sub TxtEmpId_LostFocus()
''If SaveCmd.Enabled = False Then Exit Sub
''
''If RsMst_Users.RecordCount <= 0 Then
''   Exit Sub
''End If
''RsMst_Users.MoveFirst
''RsMst_Users.Find "Employee_ID = '" & TxtEmpID & "'"
''
''If RsMst_Users.EOF Then
''   Exit Sub
''End If
''
''Call Indata
''Call EditCmd_Click
''mUserRoleId = RsMst_Users!UserGroupID
End Sub

Private Function ChkEmail() As Boolean
Dim Retval As Boolean
Dim StrSplit As Variant
Retval = True

StrSplit = Split(TxtEmailId, "@")
If UBound(StrSplit) <> 1 Then
   Retval = False
End If
ChkEmail = Retval

End Function


Private Sub CmdGo_Click()

tmp = "Select MU.UsersId,MU.EmployeeNo,MU.EmployeeName,MURole.GroupName,MU.Flag from Mst_Users MU " & _
      " Inner Join Mst_UserRole MURole On MU.UserGroupID = MURole.UserGroupID "

Dim wc As Variant
wc = ""
If TxtSearchUser <> "" Then
   wc = " Where MU.EmployeeName Like '%" & TxtSearchUser & "%' "
End If


If CmbSUserRole.Text <> "" Then
   If TxtSearchUser = "" Then
      wc = " Where MURole.GroupName = '" & CmbSUserRole.Text & "'"
   Else
      wc = wc & " And MURole.GroupName = '" & CmbSUserRole.Text & "'"
   End If
End If

If CmbSFlag.Text <> "" Then
   If wc = "" Then
      wc = " Where MU.Flag= '" & mSFlag & "'"
   Else
      wc = wc & " And MU.Flag= '" & mSFlag & "'"
   End If
End If

tmp = tmp & wc

Call TmpTable

Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount


End Sub

Private Sub TxtPassword_LostFocus()

If Edit_Flg = "E" Then Exit Sub
If TxtPassword = "" Then Exit Sub

If Len(TxtPassword) < 7 Then ' Or Len(TxtPassword) > 8 Then
   MsgBox "Password length must be more than 6 Characters!!!" 'MsgBox "Password length must be 6 to 8 Characters!!!"
   TxtPassword.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSFlag_LostFocus()
If CmbSFlag.Text = "" Then Exit Sub
If CmbSFlag.Text = "Active" Then
   mSFlag = "Y"
ElseIf CmbSFlag.Text = "De-Active" Then
   mSFlag = "N"
Else
   MsgBox "Invalid Selection!!!"
   CmbSFlag.SetFocus
   Exit Sub
End If
End Sub

