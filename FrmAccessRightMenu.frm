VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmAccessRightMenu 
   Caption         =   "AccessRight-Menu"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   12285
      TabIndex        =   30
      Top             =   120
      Width           =   12345
      Begin VB.ComboBox CmbMenuName 
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
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   9615
      End
      Begin VB.ComboBox CmbMainMenu 
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
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   9615
      End
      Begin VB.TextBox TxtMainMenu 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   120
         Width           =   9615
      End
      Begin VB.TextBox TxtMenuName 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   9615
      End
      Begin VB.Label Label9 
         Caption         =   "Menu name"
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
         Left            =   600
         TabIndex        =   34
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Head Menu"
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
         Left            =   600
         TabIndex        =   31
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   8760
      Width           =   12345
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2520
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4935
         TabIndex        =   27
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7350
         TabIndex        =   26
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9765
         TabIndex        =   25
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4935
         TabIndex        =   22
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7350
         TabIndex        =   21
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9765
         TabIndex        =   20
         Top             =   615
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   12315
      TabIndex        =   4
      Top             =   1440
      Width           =   12375
      Begin VB.TextBox TxtSrNo 
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
         TabIndex        =   11
         Top             =   480
         Width           =   855
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   5055
      End
      Begin VB.ComboBox CmbForm 
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
         ItemData        =   "FrmAccessRightMenu.frx":0000
         Left            =   6120
         List            =   "FrmAccessRightMenu.frx":000A
         TabIndex        =   9
         Text            =   "Yes"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox CmbAdd 
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
         ItemData        =   "FrmAccessRightMenu.frx":0017
         Left            =   7560
         List            =   "FrmAccessRightMenu.frx":0021
         TabIndex        =   8
         Text            =   "Yes"
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox CmbEdit 
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
         ItemData        =   "FrmAccessRightMenu.frx":002E
         Left            =   8760
         List            =   "FrmAccessRightMenu.frx":0038
         TabIndex        =   7
         Text            =   "Yes"
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox CmbDelete 
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
         ItemData        =   "FrmAccessRightMenu.frx":0045
         Left            =   9960
         List            =   "FrmAccessRightMenu.frx":004F
         TabIndex        =   6
         Text            =   "Yes"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CmdUpdateList 
         Caption         =   "Update List"
         Height          =   375
         Left            =   11160
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Sr No"
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
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Add"
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
         Left            =   7560
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Form Access"
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
         Left            =   6120
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Edit"
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
         Left            =   8760
         TabIndex        =   13
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Delete"
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
         Left            =   9960
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8040
      TabIndex        =   3
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11520
      TabIndex        =   2
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   255
      Left            =   10320
      TabIndex        =   1
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9360
      TabIndex        =   0
      Top             =   2880
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9551
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
Attribute VB_Name = "FrmAccessRightMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mUserRoleID, mLocationID As Variant
Dim mMainmenuFlg, mMenuID As Variant
Dim mUserGrp As Boolean
Private Sub GridHead()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 7

MSHFlexGrid1.TextMatrix(0, 0) = "Sr No"
MSHFlexGrid1.TextMatrix(0, 1) = "User Role"
MSHFlexGrid1.TextMatrix(0, 2) = "Form Access"
MSHFlexGrid1.TextMatrix(0, 3) = "Add"
MSHFlexGrid1.TextMatrix(0, 4) = "Edit"
MSHFlexGrid1.TextMatrix(0, 5) = "Delete"
MSHFlexGrid1.TextMatrix(0, 6) = "UserRoleID"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 6000
MSHFlexGrid1.ColWidth(2) = 1600
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 0


End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
mMainmenuFlg = ""
mMenuID = ""
CmbMainMenu = ""
CmbMenuName = ""
Call GButtonLock(Me, False)
CmbMainMenu.Visible = True
TxtMainMenu.Visible = False
CmbMenuName.Visible = True
TxtMenuName.Visible = False

Picture2.Enabled = True
Picture1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True


Call TableMst_UserRole(" Order By UserRoleName")

Call GridHead
MSHFlexGrid1.Rows = RsMst_UserRole.RecordCount + 1
If RsMst_UserRole.RecordCount > 0 Then
   For I = 1 To RsMst_UserRole.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = I
       MSHFlexGrid1.TextMatrix(I, 1) = RsMst_UserRole!UserRoleName
       MSHFlexGrid1.TextMatrix(I, 2) = ""
       MSHFlexGrid1.TextMatrix(I, 3) = ""
       MSHFlexGrid1.TextMatrix(I, 4) = ""
       MSHFlexGrid1.TextMatrix(I, 5) = ""
       MSHFlexGrid1.TextMatrix(I, 6) = RsMst_UserRole!UserRoleID
       RsMst_UserRole.MoveNext
   Next

End If

End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 2) = "Yes"
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 2) = ""
    Next
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(I, 1) <> "Report" Then
           MSHFlexGrid1.TextMatrix(I, 2) = "Yes"
           MSHFlexGrid1.TextMatrix(I, 5) = "Yes"
        End If
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 5) = ""
    Next
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(I, 1) <> "Report" Then
           MSHFlexGrid1.TextMatrix(I, 2) = "Yes"
           MSHFlexGrid1.TextMatrix(I, 4) = "Yes"
        End If
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 4) = ""
    Next
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(I, 1) <> "Report" Then
           MSHFlexGrid1.TextMatrix(I, 3) = "Yes"
           MSHFlexGrid1.TextMatrix(I, 2) = "Yes"
        End If
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 3) = ""
    Next
End If
End Sub


Private Sub CmbMainMenu_GotFocus()
tmp1 = CmbMainMenu.Text
tmp = "select Distinct HMenuFlag from Mst_MenuDetail"
Call TmpTable
CmbMainMenu.Clear
CmbMenuName.Clear

If TmpRs.RecordCount = 0 Then
   Exit Sub
End If
For I = 0 To TmpRs.RecordCount - 1
    If TmpRs!HMenuFlag = "A" Then
       CmbMainMenu.AddItem "Admin"
    End If
    If TmpRs!HMenuFlag = "M" Then
       CmbMainMenu.AddItem "Master"
    End If
    If TmpRs!HMenuFlag = "T" Then
       CmbMainMenu.AddItem "Transaction"
    End If
    If TmpRs!HMenuFlag = "R" Then
       CmbMainMenu.AddItem "Report"
    End If
    If TmpRs!HMenuFlag = "C" Then
       CmbMainMenu.AddItem "COG Update"
    End If
    TmpRs.MoveNext
Next
CmbMainMenu.Text = tmp1
End Sub

Private Sub CmbMainMenu_LostFocus()
If CmbMainMenu.Text = "" Then
   Exit Sub
End If
TmpRs.MoveFirst

If LCase(CmbMainMenu.Text) <> LCase("Admin") And LCase(CmbMainMenu.Text) <> LCase("Master") And LCase(CmbMainMenu.Text) <> LCase("Transaction") And LCase(CmbMainMenu.Text) <> LCase("Report") And LCase(CmbMainMenu.Text) <> LCase("COG Update") Then
   MsgBox "Invalid MainMenu"
   CmbMainMenu.SetFocus
   Exit Sub
End If
tmp = CmbMainMenu.Text
End Sub




Private Sub CmbMenuName_GotFocus()
If CmbMainMenu.Text = "" Then Exit Sub
tmp = CmbMenuName.Text
If LCase(CmbMainMenu.Text) = LCase("Admin") Then
   mMainmenuFlg = "A"
End If
If LCase(CmbMainMenu.Text) = LCase("Master") Then
   mMainmenuFlg = "M"
End If
If LCase(CmbMainMenu.Text) = LCase("Transaction") Then
   mMainmenuFlg = "T"
End If
If LCase(CmbMainMenu.Text) = LCase("Report") Then
   mMainmenuFlg = "R"
End If
If LCase(CmbMainMenu.Text) = LCase("COG Update") Then
   mMainmenuFlg = "C"
End If
CmbMenuName.Clear
Call TableMst_MenuDetail("Where HMenuFlag='" & mMainmenuFlg & "'")
If RsMst_MenuDetail.RecordCount > 0 Then
   For I = 1 To RsMst_MenuDetail.RecordCount
       CmbMenuName.AddItem Trim(RsMst_MenuDetail!MenuName)
       RsMst_MenuDetail.MoveNext
   Next
Else
   MsgBox "Menu Names Not Found."
End If
CmbMenuName.Text = tmp
End Sub

Private Sub CmbMenuName_LostFocus()

If CmbMenuName.Text = "" Then
   Exit Sub
End If
RsMst_MenuDetail.MoveFirst
RsMst_MenuDetail.Find "MenuName = '" & CmbMenuName.Text & "'"
If RsMst_MenuDetail.EOF Then
   MsgBox "Invalid Menu Name "
   CmbMenuName.SetFocus
   Exit Sub
End If
mMenuID = RsMst_MenuDetail!MenuID
tmp = CmbMenuName.Text
tmp = "select A.MenuID,A.UserRoleID,UserRoleName,FormAccess " & _
      ",ADD_,EDIT_,Save_,Delete_ from Mst_MenuHasUserRole A " & _
      "inner join Mst_UserRole on A.UserRoleID= Mst_UserRole.UserRoleID " & _
      "where A.MenuID= " & mMenuID & " Order By UserRoleName"
Call TmpTable
Call GridHead
MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = I
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!UserRoleName
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(TmpRs!formaccess = "Y", "Yes", "")
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(TmpRs!ADD_ = "Y", "Yes", "")
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(TmpRs!EDIT_ = "Y", "Yes", "")
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(TmpRs!Delete_ = "Y", "Yes", "")
       MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!UserRoleID
       TmpRs.MoveNext
   Next
End If
Call TableMst_UserRole(" Order By UserRoleName")
If RsMst_UserRole.RecordCount = 0 Then
   Exit Sub
End If

For I = 1 To RsMst_UserRole.RecordCount
    mUserGrp = False
    For r = 1 To MSHFlexGrid1.Rows - 1
        If RsMst_UserRole!UserRoleID = MSHFlexGrid1.TextMatrix(r, 6) Then
           mUserGrp = True
           Exit For
        Else
           mUserGrp = False
        End If
    Next
    If mUserGrp = False Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = RsMst_UserRole!UserRoleName
       MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 6) = RsMst_UserRole!UserRoleID
    End If
    RsMst_UserRole.MoveNext
Next
End Sub

Private Sub CmdUpdateList_Click()
Dim flg As Variant
FAG = 0
x = MSHFlexGrid1.RowSel
If x <= 0 Then Exit Sub

If CmbAdd.Text = "Yes" Or CmbEdit.Text = "Yes" Or CmbDelete.Text = "Yes" Then
   CmbForm.Text = "Yes"
End If
If MSHFlexGrid1.TextMatrix(x, 1) = "Report" Then
   CmbAdd.Text = "No"
   CmbEdit.Text = "No"
   CmbDelete.Text = "No"
End If

MSHFlexGrid1.TextMatrix(x, 2) = IIf(CmbForm.Text = "Yes", "Yes", "")
MSHFlexGrid1.TextMatrix(x, 3) = IIf(CmbAdd.Text = "Yes", "Yes", "")
MSHFlexGrid1.TextMatrix(x, 4) = IIf(CmbEdit.Text = "Yes", "Yes", "")
MSHFlexGrid1.TextMatrix(x, 5) = IIf(CmbDelete.Text = "Yes", "Yes", "")

End Sub

Private Sub EditCmd_Click()
Picture2.Enabled = True
Picture1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True

Edit_Flg = "E"
Call GButtonLock(Me, False)
CmbMainMenu.Visible = True
TxtMainMenu.Visible = False
CmbMenuName.Visible = True
TxtMenuName.Visible = False
CmbMainMenu.Text = TxtMainMenu
CmbMenuName.Text = TxtMenuName
End Sub

Private Sub Form_Load()
Picture2.Enabled = False
Picture1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False

Gen_AzureDbPath = "N"
tmp = "select Distinct MenuID " & _
      "From Mst_MenuHasUserRole"
Call Tmp2Table
If TmpRs2.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
'
'
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied!!!!!!!!!"
   Exit Sub
End If

Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()

Picture2.Enabled = False
Picture1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False

CmbMainMenu.Visible = False
TxtMainMenu.Visible = True
CmbMenuName.Visible = False
TxtMenuName.Visible = True


tmp = "select * from Mst_MenuDetail where MenuID = " & TmpRs2!MenuID & ""
Call TmpTable
mMenuID = TmpRs!MenuID
TxtMenuName = TmpRs!MenuName
mMainmenuFlg = TmpRs!HMenuFlag
If TmpRs!HMenuFlag = "A" Then
      TxtMainMenu = "Admin"
    End If
    If TmpRs!HMenuFlag = "M" Then
       TxtMainMenu = "Master"
    End If
    If TmpRs!HMenuFlag = "T" Then
       TxtMainMenu = "Transaction"
    End If
    If TmpRs!HMenuFlag = "R" Then
       TxtMainMenu = "Report"
    End If
TmpRs.Close

tmp = "select a.UserRoleID,UserRoleName,a.FormAccess, a.ADD_, a.EDIT_, a.Delete_,a.MenuID From Mst_MenuHasUserRole a " & _
      "inner join Mst_UserRole on a.UserRoleID= Mst_UserRole.UserRoleID where a.MenuID=" & mMenuID & " Order By UserRoleName"
Call TmpTable
Call GridHead

MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
For I = 1 To TmpRs.RecordCount
        MSHFlexGrid1.TextMatrix(I, 0) = I
        MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!UserRoleName
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(TmpRs!formaccess = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 3) = IIf(TmpRs!ADD_ = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(TmpRs!EDIT_ = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(TmpRs!Delete_ = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!UserRoleID
        TmpRs.MoveNext
Next
Call TableMst_UserRole(" Order By UserRoleName")
If RsMst_UserRole.RecordCount = 0 Then
   Exit Sub
End If

For I = 1 To RsMst_UserRole.RecordCount
    mUserGrp = True
    For r = 1 To MSHFlexGrid1.Rows - 1
        If RsMst_UserRole!UserRoleID = MSHFlexGrid1.TextMatrix(r, 6) Then
           mUserGrp = True
           Exit For
        Else
           mUserGrp = False
        End If
    Next
    If mUserGrp = False Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = RsMst_UserRole!UserRoleName
       MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 6) = RsMst_UserRole!UserRoleID
    End If
    RsMst_UserRole.MoveNext
Next

Call GButtonLock(Me, True)
End Sub

Private Sub NextCmd_Click()
TmpRs2.MoveNext
If TmpRs2.EOF Then
   TmpRs2.MoveLast
End If
Call Indata
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
TmpRs2.MoveFirst
Call Indata

End Sub

Private Sub LastCmd_Click()
TmpRs2.MoveLast
Call Indata

End Sub



Private Sub PreviousCmd_Click()
TmpRs2.MovePrevious
If TmpRs2.BOF Then
   TmpRs2.MoveFirst
End If
Call Indata

End Sub


Private Sub MSHFlexGrid1_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
If x <= 0 Then Exit Sub


TxtSrNo = MSHFlexGrid1.TextMatrix(x, 0)
TxtUserRole = MSHFlexGrid1.TextMatrix(x, 1)
CmbForm.Text = MSHFlexGrid1.TextMatrix(x, 2)
CmbAdd.Text = MSHFlexGrid1.TextMatrix(x, 3)
CmbEdit.Text = MSHFlexGrid1.TextMatrix(x, 4)
CmbDelete.Text = MSHFlexGrid1.TextMatrix(x, 5)

End Sub

Private Sub SaveCmd_Click()
If CmbMainMenu.Text = "" Then
   MsgBox "Blank Head Menu not allowed !!!! "
   Exit Sub
End If
If CmbMenuName.Text = "" Then
   MsgBox "Blank Menu Name not allowed !!!! "
   Exit Sub
End If
If mMainmenuFlg = "" Then
   MsgBox "Blank Head Menu not allowed !!!! "
   Exit Sub
End If
If mMenuID = "" Then
   MsgBox "Blank Menu Name not allowed !!!! "
   Exit Sub
End If
Call Mst_MenuHasUserRoleTable(" Where MenuID = " & mMenuID & "")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 6) <> "" Then
        If RsMenuUserRole.RecordCount > 0 Then
           RsMenuUserRole.MoveFirst
           RsMenuUserRole.Find "UserRoleID = " & MSHFlexGrid1.TextMatrix(I, 6) & ""
        End If
        If RsMenuUserRole.EOF Then
           RsMenuUserRole.AddNew
           RsMenuUserRole!UserRoleID = MSHFlexGrid1.TextMatrix(I, 6)
           RsMenuUserRole!MenuID = mMenuID
           RsMenuUserRole!G_UID = GetGUID
        End If
        RsMenuUserRole!formaccess = IIf(MSHFlexGrid1.TextMatrix(I, 2) = "Yes", "Y", "N")
        RsMenuUserRole!ADD_ = IIf(MSHFlexGrid1.TextMatrix(I, 3) = "Yes", "Y", "N")
        RsMenuUserRole!EDIT_ = IIf(MSHFlexGrid1.TextMatrix(I, 4) = "Yes", "Y", "N")
        RsMenuUserRole!Delete_ = IIf(MSHFlexGrid1.TextMatrix(I, 5) = "Yes", "Y", "N")
        RsMenuUserRole.Update
    End If
Next
Call GButtonLock(Me, True)

'Call Mst_MenuHasUserRoleTable(" order by MenuID,UserRoleID")
TmpRs2.Requery
TmpRs2.MoveFirst
TmpRs2.Find "MenuID = " & mMenuID

'Call InData


End Sub
'---------------
Private Sub SearchCmd_Click()
'
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If Edit_Flg = "A" Then
      TmpRs2.MoveFirst
   End If
   Call Indata
   Exit Sub
End If
'
'Frame0.Height = 6615
'Frame0.Left = 120
'Frame0.Top = 120
'Frame0.Width = 7100
'Frame0.Visible = True
'MSHFlexGrid1.Clear
'
'MSHFlexGrid1.Width = Frame0.Width - 200
'MSHFlexGrid1.Height = Frame0.Height - 500
'
'TmpRs2.MoveFirst
'MSHFlexGrid1.Rows = TmpRs2.RecordCount + 1
'MSHFlexGrid1.TextMatrix(0, 0) = "USer Id"
'MSHFlexGrid1.TextMatrix(0, 1) = "User"
'
'MSHFlexGrid1.ColWidth(0) = 1200
'MSHFlexGrid1.ColWidth(1) = 4500
'For i = 1 To TmpRs2.RecordCount
'    MSHFlexGrid1.TextMatrix(i, 0) = TmpRs2.Fields(0)
'    MSHFlexGrid1.TextMatrix(i, 1) = TmpRs2.Fields(2)
'    TmpRs2.MoveNext
'Next

End Sub


