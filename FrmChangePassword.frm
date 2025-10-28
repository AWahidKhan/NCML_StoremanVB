VERSION 5.00
Begin VB.Form FrmChangePassword 
   Caption         =   "Change Password"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   7290
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Width           =   7095
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   1575
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox TxtNewPassword 
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
         Left            =   1680
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   4200
         Width           =   2295
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   720
         Width           =   5295
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
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
         Left            =   1680
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox TxtAddress 
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
         Height          =   1215
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox TxtContact 
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox TxtUserRole 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label9 
         Caption         =   "New Password"
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
         TabIndex        =   19
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   6960
         TabIndex        =   14
         Top             =   4920
         Width           =   1335
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
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "User Name"
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
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Employee ID"
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
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Width           =   1695
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
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Old Password"
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
         TabIndex        =   12
         Top             =   3600
         Width           =   1455
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
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Chk_Password, mOldPassword As Variant

Private Sub CmdExit_Click()
If Gen_PassExpire = "Y" Then
   End
Else
   Unload Me
End If
End Sub

Private Sub CmdSave_Click()
Dim ans As Variant
If TxtEmpID = "" Then Exit Sub
If TxtPassword = "" Then
   MsgBox "Old Password can not be blank!!! "
   TxtPassword.SetFocus
   Exit Sub
End If
If TxtNewPassword = "" Then
   MsgBox "New Password can not be blank!!! "
   TxtNewPassword.SetFocus
   Exit Sub
End If
If TxtPassword = TxtNewPassword Then
   MsgBox "Old and new password cannot be same!!!"
   TxtNewPassword.SetFocus
   Exit Sub
End If
ans = MsgBox("Do you want to save new password??", vbYesNo)
If ans = vbYes Then
   RsMst_Users!Password = TxtNewPassword
   RsMst_Users!PassExpiryDate = Date + 14
   If IsNull(RsMst_Users!CreatedBy) = True Or RsMst_Users!CreatedBy = "" Then
      RsMst_Users!CreatedBy = Gen_UserLoginID
      RsMst_Users!CreatedDate = Now
   Else
      RsMst_Users!UpdatedBy = Gen_UserLoginID
      RsMst_Users!UpdatedDate = Now
   End If
   RsMst_Users.Update
Else
   TxtNewPassword.SetFocus
   Exit Sub
End If
End
End Sub

Private Sub Form_Load()


Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_Users

RsMst_Users.Find "UsersID = " & Val(Gen_UserID) & ""

If RsMst_Users.EOF Then
   Exit Sub
End If

'Parag
'tmp = "select * from Mst_UserRole where UserGroupID = " & RsMst_Users!UserGRoupID & ""
tmp = "select * from Mst_UserRole where UserRoleID = " & RsMst_Users!UserRoleID & ""
Call TmpTable
TxtUsersID = RsMst_Users!UsersID
'Parag
'TxtUsers = RsMst_Users!Employee_Name
TxtUsers = RsMst_Users!EmployeeName
TxtAddress = IIf(IsNull(RsMst_Users!Address), "", RsMst_Users!Address)
TxtContact = IIf(IsNull(RsMst_Users!ContactNo), "", RsMst_Users!ContactNo)
'Parag
'TxtEmpID = RsMst_Users!Employee_ID
TxtEmpID = RsMst_Users!EmployeeNo
TxtPassword = ""
Chk_Password = RsMst_Users!Password
'TxtUsers = RsMst_Users!Employee_Name
TxtUsers = RsMst_Users!EmployeeName
TxtUserRole = TmpRs.Fields(1)




End Sub

Private Sub Form_Unload(Cancel As Integer)

If Gen_PassExpire = "Y" Then
   End
End If

End Sub

Private Sub TxtNewPassword_LostFocus()

If TxtNewPassword = "" Then Exit Sub

If Len(TxtNewPassword) < 7 Then 'Or Len(TxtNewPassword) > 8 Then
   MsgBox "Password length must be more than 6 Characters!!!"
   TxtNewPassword.SetFocus
   Exit Sub
End If

If IsAlphaNumeric(TxtNewPassword) = False Then
   TxtNewPassword.SetFocus
   Exit Sub
End If


End Sub

Private Sub TxtPassword_LostFocus()
If TxtPassword = "" Then Exit Sub
If Chk_Password <> TxtPassword Then
   MsgBox "Invalid Password!!! "
   TxtPassword.SetFocus
End If

End Sub


Private Function IsAlphaNumeric(mPwd_ As String) As Boolean
Dim isNum As Boolean
Dim isAlpha As Boolean
Dim isSpecial As Boolean

IsAlphaNumeric = False
isNum = False
isAlpha = False
isSpecial = False

If Len(mPwd_) > 0 Then
   For lng = 1 To Len(mPwd_)
      Select Case Asc(mID$(mPwd_, lng, 1))
      Case 32
         isAlpha = True 'Space
      Case 48 To 57
         isNum = True 'Numbers 0 to 9
      Case 65 To 90
         isAlpha = True 'Letters A to Z
      Case 97 To 122
         isAlpha = True 'Letters a to z
      Case Else
         isSpecial = True
      End Select
   Next lng
End If

If isNum = False Or isAlpha = False Or isSpecial = False Then
   'MsgBox "Password should be AlphaNumeric!!!"
   MsgBox "Password should be contain atleast 1 Character, 1 Number and 1 Special Character!!!"
Else
   IsAlphaNumeric = True
End If
End Function

