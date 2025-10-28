VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmUsersRoleMaster 
   Caption         =   "Users Role Master"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   6735
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2535
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1620
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2858
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   6360
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5085
         TabIndex        =   17
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3870
         TabIndex        =   16
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2535
         TabIndex        =   15
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5085
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3870
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2535
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin VB.TextBox TxtUsersRole 
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
         Left            =   2400
         TabIndex        =   4
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox TxtUsersRoleID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Users Role"
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
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Users Role  Id"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmUsersRoleMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
'Call GButtonLock_Pwd(Me, False)
Call GButtonLock(Me, False)
TxtUsersRole.Locked = False
TxtUsersRoleID = ""
TxtUsersRole = ""
If RsMst_UserRole.RecordCount > 0 Then
   TxtUsersRole.SetFocus
End If
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   RsMst_UserRole.Delete
   RsMst_UserRole.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_UserRole.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_UserRole.MoveNext
   If RsMst_UserRole.EOF() Then
      RsMst_UserRole.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_UserRole.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
'Call GButtonLock_Pwd(Me, False)
Call GButtonLock(Me, False)
TxtUsersRole.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_UserRole.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True

End Sub

Private Sub Form_Load()

'If UserType = "Power" Then
'
'Else
'   Call GButtonLockAD(Me)
'   MsgBox "Access denied !!!!!!!!!"
'   Exit Sub
'End If

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If


'Call StatusBarData(Me)

Call TableMst_UserRole
If RsMst_UserRole.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata

Call GButtonLock(Me, True)
'Call GButtonLock_Pwd(Me, True)


End Sub

Public Sub Indata()
TxtUsersRole.Locked = True

TxtUsersRoleID = RsMst_UserRole!UserGroupID
TxtUsersRole = RsMst_UserRole!GroupName
'Call GButtonLock_Pwd(Me, True)

Call GButtonLock(Me, True)

End Sub

Private Sub LastCmd_Click()
RsMst_UserRole.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False

End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_UserRole.MoveFirst

If RsMst_UserRole.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_UserRole.MoveFirst
RsMst_UserRole.Find "UserGroupID = " & Val(mCaId) & ""
If RsMst_UserRole.EOF Then
   RsMst_UserRole.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_UserRole.MoveNext
If RsMst_UserRole.EOF Then
   RsMst_UserRole.MoveLast
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
RsMst_UserRole.MovePrevious
If RsMst_UserRole.BOF Then
   RsMst_UserRole.MoveFirst
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
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtUsersRole = "" Then
   MsgBox "Blank User Role not allowed"
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_UserRole.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_UserRole.MoveFirst
        RsMst_UserRole.Find "GroupName = '" & TxtUsersRole & "'"
        If Not RsMst_UserRole.EOF Then
           MsgBox "Duplicate User Role not allowed"
           Exit Sub
        End If
        RsMst_UserRole.MoveLast
        tmp = RsMst_UserRole!UserGroupID + 1
    End If
    RsMst_UserRole.AddNew
    RsMst_UserRole!UserGroupID = tmp
    RsMst_UserRole!G_UID = GetGUID
Else
    RsMst_UserRole.MoveFirst
    RsMst_UserRole.Find "GroupName = '" & TxtUsersRole & "'"
    If Not RsMst_UserRole.EOF Then
       If RsMst_UserRole!UserGroupID <> TxtUsersRoleID Then
          MsgBox "Duplicate User Role not allowed"
          Exit Sub
       End If
    End If
   RsMst_UserRole.MoveFirst
   RsMst_UserRole.Find "UserGroupID = " & TxtUsersRoleID & ""
End If
RsMst_UserRole!GroupName = TxtUsersRole
'UsersRs!BatchID = Null

If IsNull(RsMst_UserRole!CreatedBy) = True Or RsMst_UserRole!CreatedBy = "" Then
   RsMst_UserRole!CreatedBy = Gen_UserLoginID
   RsMst_UserRole!CreatedDate = Now
Else
   RsMst_UserRole!UpdatedBy = Gen_UserLoginID
   RsMst_UserRole!UpdatedDate = Now
End If
RsMst_UserRole.Update
'Call GButtonLock_Pwd(Me, True)

Call GButtonLock(Me, True)
TxtUsersRole.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
Frame0.Height = 2535
Frame0.Left = 120
Frame0.Top = 120
Frame0.Width = 7100
Frame0.Visible = True
MSHFlexGrid1.Clear

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 500

RsMst_UserRole.MoveFirst
MSHFlexGrid1.Rows = RsMst_UserRole.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "User Role Id"
MSHFlexGrid1.TextMatrix(0, 1) = "User Role"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4500
For I = 1 To RsMst_UserRole.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_UserRole.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_UserRole.Fields(1)
    RsMst_UserRole.MoveNext
Next

End Sub

