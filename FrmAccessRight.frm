VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmAccessRight 
   Caption         =   "Access Right Management"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9360
      TabIndex        =   20
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   255
      Left            =   10320
      TabIndex        =   19
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   11520
      TabIndex        =   18
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   8040
      TabIndex        =   17
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1275
      ScaleWidth      =   12315
      TabIndex        =   16
      Top             =   1080
      Width           =   12375
      Begin VB.CommandButton CmdUpdateList 
         Caption         =   "Update List"
         Height          =   375
         Left            =   10920
         TabIndex        =   35
         Top             =   480
         Width           =   1215
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
         ItemData        =   "FrmAccessRight.frx":0000
         Left            =   9480
         List            =   "FrmAccessRight.frx":000A
         TabIndex        =   34
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
         ItemData        =   "FrmAccessRight.frx":0017
         Left            =   8400
         List            =   "FrmAccessRight.frx":0021
         TabIndex        =   32
         Text            =   "Yes"
         Top             =   480
         Width           =   1095
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
         ItemData        =   "FrmAccessRight.frx":002E
         Left            =   7320
         List            =   "FrmAccessRight.frx":0038
         TabIndex        =   31
         Text            =   "Yes"
         Top             =   480
         Width           =   1095
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
         ItemData        =   "FrmAccessRight.frx":0045
         Left            =   6000
         List            =   "FrmAccessRight.frx":004F
         TabIndex        =   30
         Text            =   "Yes"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtMName 
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox TxtHMenu 
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
         TabIndex        =   28
         Top             =   480
         Width           =   2175
      End
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
         TabIndex        =   27
         Top             =   480
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
         Left            =   9600
         TabIndex        =   33
         Top             =   120
         Width           =   1215
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
         Left            =   8520
         TabIndex        =   26
         Top             =   120
         Width           =   855
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
         Left            =   6000
         TabIndex        =   25
         Top             =   120
         Width           =   1335
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
         Left            =   7440
         TabIndex        =   24
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Menu Name"
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
         Left            =   3600
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   120
         Width           =   1815
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
         TabIndex        =   21
         Top             =   120
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5775
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10186
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
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   8760
      Width           =   12345
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9765
         TabIndex        =   11
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7350
         TabIndex        =   10
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4935
         TabIndex        =   9
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9765
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7350
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4935
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   12285
      TabIndex        =   0
      Top             =   120
      Width           =   12345
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   10335
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   120
         Width           =   10455
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
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmAccessRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mUserRoleID, mLocationID As Variant
Private Sub GridHead()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8

MSHFlexGrid1.TextMatrix(0, 0) = "Sr No"
MSHFlexGrid1.TextMatrix(0, 1) = "Head Menu"
MSHFlexGrid1.TextMatrix(0, 2) = "Menu Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Form Access"
MSHFlexGrid1.TextMatrix(0, 4) = "Add"
MSHFlexGrid1.TextMatrix(0, 5) = "Edit"
MSHFlexGrid1.TextMatrix(0, 6) = "Delete"
MSHFlexGrid1.TextMatrix(0, 7) = "MenuID"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 0


End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
CmbUserRole.Visible = True
TxtUserRole.Visible = False

Picture2.Enabled = True
Picture1.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True


Call Mst_MenuDetailTable("WHERE     Flag = 'Y' ORDER BY HMenuFlag, MenuID")

Call GridHead
MSHFlexGrid1.Rows = RsMenuDetail.RecordCount + 2
If RsMenuDetail.RecordCount > 0 Then
   For I = 1 To RsMenuDetail.RecordCount
   
       MSHFlexGrid1.TextMatrix(I, 0) = I
       
       If RsMenuDetail!HMenuFlag = "A" Then
          MSHFlexGrid1.TextMatrix(I, 1) = "Admin"
       End If
       If RsMenuDetail!HMenuFlag = "M" Then
          MSHFlexGrid1.TextMatrix(I, 1) = "Master"
       End If
       If RsMenuDetail!HMenuFlag = "T" Then
          MSHFlexGrid1.TextMatrix(I, 1) = "Transaction"
       End If
       If RsMenuDetail!HMenuFlag = "R" Then
          MSHFlexGrid1.TextMatrix(I, 1) = "Report"
       End If
       
       MSHFlexGrid1.TextMatrix(I, 2) = RsMenuDetail!MenuName
       MSHFlexGrid1.TextMatrix(I, 3) = ""
       MSHFlexGrid1.TextMatrix(I, 4) = ""
       MSHFlexGrid1.TextMatrix(I, 5) = ""
       MSHFlexGrid1.TextMatrix(I, 6) = ""
       MSHFlexGrid1.TextMatrix(I, 7) = RsMenuDetail!MenuID
       RsMenuDetail.MoveNext
   Next

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

Private Sub Check1_Click()
If Check1.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 3) = "Yes"
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 3) = ""
    Next
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(I, 1) <> "Report" Then
           MSHFlexGrid1.TextMatrix(I, 3) = "Yes"
           MSHFlexGrid1.TextMatrix(I, 6) = "Yes"
        End If
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 6) = ""
    Next
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(I, 1) <> "Report" Then
           MSHFlexGrid1.TextMatrix(I, 3) = "Yes"
           MSHFlexGrid1.TextMatrix(I, 5) = "Yes"
        End If
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 5) = ""
    Next
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    For I = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(I, 1) <> "Report" Then
           MSHFlexGrid1.TextMatrix(I, 4) = "Yes"
           MSHFlexGrid1.TextMatrix(I, 3) = "Yes"
        End If
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 4) = ""
    Next
End If
End Sub

Private Sub CmbUserRole_Click()

tmp = "select * from Mst_UserRole where GroupName  = '" & CmbUserRole.Text & "'"
Call TmpTable
mUserRoleID = TmpRs.Fields(0)
TxtUserRole = IIf(IsNull(TmpRs.Fields(1)), "", TmpRs.Fields(1))
TmpRs.Close

tmp = "select b.HMenuFlag, b.MenuName,  b.MenuID "
tmp = tmp & " From Mst_MenuDetail b"
tmp = tmp & " left join Mst_MenuHasUserRole a   on  b.MenuID = a.MenuID"
tmp = tmp & " Where a.UserGroupID = " & mUserRoleID & ""
tmp = tmp & " Union"
tmp = tmp & " select b.HMenuFlag, b.MenuName, b.MenuID"
tmp = tmp & " from Mst_MenuDetail b"
tmp = tmp & " order by b.HMenuFlag,b.MenuID"

Call Tmp4Table



tmp = "select b.HMenuFlag, b.MenuName, a.FormAccess, a.ADD_, a.EDIT_, a.Delete_, b.MenuID From Mst_MenuHasUserRole a " & _
"inner join Mst_MenuDetail b on a.MenuID = b.MenuID Where a.UserGroupID = " & mUserRoleID & " order by b.HMenuFlag,b.MenuID"
Call TmpTable

Call GridHead

MSHFlexGrid1.Rows = TmpRs4.RecordCount + 1
For I = 1 To TmpRs4.RecordCount
        MSHFlexGrid1.TextMatrix(I, 0) = I
        If TmpRs4!HMenuFlag = "A" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Admin"
        End If
        If TmpRs4!HMenuFlag = "M" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Master"
        End If
        If TmpRs4!HMenuFlag = "T" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Transaction"
        End If
        If TmpRs4!HMenuFlag = "R" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Report"
        End If
        MSHFlexGrid1.TextMatrix(I, 2) = TmpRs4!MenuName
        If TmpRs.RecordCount > 0 Then
           TmpRs.MoveFirst
           TmpRs.Find "MenuID = " & TmpRs4!MenuID
           If Not TmpRs.EOF Then
              MSHFlexGrid1.TextMatrix(I, 3) = IIf(TmpRs!formaccess = "Y", "Yes", "")
              MSHFlexGrid1.TextMatrix(I, 4) = IIf(TmpRs!ADD_ = "Y", "Yes", "")
              MSHFlexGrid1.TextMatrix(I, 5) = IIf(TmpRs!EDIT_ = "Y", "Yes", "")
              MSHFlexGrid1.TextMatrix(I, 6) = IIf(TmpRs!Delete_ = "Y", "Yes", "")
           End If
        End If
        MSHFlexGrid1.TextMatrix(I, 7) = TmpRs4!MenuID
        TmpRs4.MoveNext
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

MSHFlexGrid1.TextMatrix(x, 3) = IIf(CmbForm.Text = "Yes", "Yes", "")
MSHFlexGrid1.TextMatrix(x, 4) = IIf(CmbAdd.Text = "Yes", "Yes", "")
MSHFlexGrid1.TextMatrix(x, 5) = IIf(CmbEdit.Text = "Yes", "Yes", "")
MSHFlexGrid1.TextMatrix(x, 6) = IIf(CmbDelete.Text = "Yes", "Yes", "")
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
CmbUserRole.Visible = True
TxtUserRole.Visible = False
CmbUserRole.Text = TxtUserRole

End Sub

Private Sub Form_Load()
Picture2.Enabled = False
Picture1.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False

'Call StatusBarData(Me)
Call Mst_MenuHasUserRoleTable(" order by UserGroupID,MenuID")

If RsMenuUserRole.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If


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

CmbUserRole.Visible = False
TxtUserRole.Visible = True


tmp = "select * from Mst_UserRole where UserGroupID = " & RsMenuUserRole!UserGRoupID & ""
Call TmpTable
mUserRoleID = TmpRs.Fields(0)
TxtUserRole = IIf(IsNull(TmpRs.Fields(1)), "", TmpRs.Fields(1))
TmpRs.Close

tmp = "select b.HMenuFlag, b.MenuName, a.FormAccess, a.ADD_, a.EDIT_, a.Delete_, b.MenuID From Mst_MenuHasUserRole a " & _
"inner join Mst_MenuDetail b on a.MenuID = b.MenuID Where a.UserGroupID = " & mUserRoleID & " order by b.HMenuFlag,b.MenuID"
Call TmpTable
Call GridHead

MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
For I = 1 To TmpRs.RecordCount
        MSHFlexGrid1.TextMatrix(I, 0) = I
        If TmpRs!HMenuFlag = "A" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Admin"
        End If
        If TmpRs!HMenuFlag = "M" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Master"
        End If
        If TmpRs!HMenuFlag = "T" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Transaction"
        End If
        If TmpRs!HMenuFlag = "R" Then
           MSHFlexGrid1.TextMatrix(I, 1) = "Report"
        End If
        MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!MenuName
        MSHFlexGrid1.TextMatrix(I, 3) = IIf(TmpRs!formaccess = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(TmpRs!ADD_ = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(TmpRs!EDIT_ = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 6) = IIf(TmpRs!Delete_ = "Y", "Yes", "")
        MSHFlexGrid1.TextMatrix(I, 7) = TmpRs!MenuID
        TmpRs.MoveNext
Next

Call GButtonLock(Me, True)
End Sub

Private Sub NextCmd_Click()
RsMenuUserRole.MoveNext
If RsMenuUserRole.EOF Then
   RsMenuUserRole.MoveLast
End If
Call Indata
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMenuUserRole.MoveFirst
Call Indata

End Sub

Private Sub LastCmd_Click()
RsMenuUserRole.MoveLast
Call Indata

End Sub

Private Sub PreviousCmd_Click()
RsMenuUserRole.MovePrevious
If RsMenuUserRole.BOF Then
   RsMenuUserRole.MoveFirst
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
TxtHMenu = MSHFlexGrid1.TextMatrix(x, 1)
TxtMName = MSHFlexGrid1.TextMatrix(x, 2)

CmbForm.Text = MSHFlexGrid1.TextMatrix(x, 3)
CmbAdd.Text = MSHFlexGrid1.TextMatrix(x, 4)
CmbEdit.Text = MSHFlexGrid1.TextMatrix(x, 5)
CmbDelete.Text = MSHFlexGrid1.TextMatrix(x, 6)

End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If mUserRoleID = "" Then
   MsgBox "Blank User Role not allowed !!!! "
   Exit Sub
End If
Call Mst_MenuHasUserRoleTable(" Where UserGroupID = " & mUserRoleID)
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 7) <> "" Then
        If RsMenuUserRole.RecordCount > 0 Then
           RsMenuUserRole.MoveFirst
           RsMenuUserRole.Find "MenuID = " & MSHFlexGrid1.TextMatrix(I, 7)
        End If
        If RsMenuUserRole.EOF Then
           RsMenuUserRole.AddNew
           RsMenuUserRole!MenuID = MSHFlexGrid1.TextMatrix(I, 7)
           RsMenuUserRole!UserGRoupID = mUserRoleID
           RsMenuUserRole!G_UID = GetGUID
        End If
        RsMenuUserRole!formaccess = IIf(MSHFlexGrid1.TextMatrix(I, 3) = "Yes", "Y", "N")
        RsMenuUserRole!ADD_ = IIf(MSHFlexGrid1.TextMatrix(I, 4) = "Yes", "Y", "N")
        RsMenuUserRole!EDIT_ = IIf(MSHFlexGrid1.TextMatrix(I, 5) = "Yes", "Y", "N")
        RsMenuUserRole!Delete_ = IIf(MSHFlexGrid1.TextMatrix(I, 6) = "Yes", "Y", "N")
        RsMenuUserRole.Update
    End If
Next
Call GButtonLock(Me, True)

RsMenuUserRole.Close

Call Mst_MenuHasUserRoleTable(" order by UserGroupID,MenuID")
RsMenuUserRole.Find "UserGroupID = " & mUserRoleID
'Call InData


End Sub
'---------------
Private Sub SearchCmd_Click()
'
'If SearchCmd.Caption = "Cancel" Then
'   Call GButtonLock(Me, True)
'   If Edit_Flg = "A" Then
'      RsMenuUserRole.MoveFirst
'   End If
'   Call InData
'   Exit Sub
'End If
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
'RsMenuUserRole.MoveFirst
'MSHFlexGrid1.Rows = RsMenuUserRole.RecordCount + 1
'MSHFlexGrid1.TextMatrix(0, 0) = "USer Id"
'MSHFlexGrid1.TextMatrix(0, 1) = "User"
'
'MSHFlexGrid1.ColWidth(0) = 1200
'MSHFlexGrid1.ColWidth(1) = 4500
'For i = 1 To RsMenuUserRole.RecordCount
'    MSHFlexGrid1.TextMatrix(i, 0) = RsMenuUserRole.Fields(0)
'    MSHFlexGrid1.TextMatrix(i, 1) = RsMenuUserRole.Fields(2)
'    RsMenuUserRole.MoveNext
'Next

End Sub

Private Sub CmbUserRole_GotFocus()
tmp = CmbUserRole.Text
Call TableMst_UserRole
CmbUserRole.Clear

If RsMst_UserRole.RecordCount = 0 Then
   Exit Sub
End If
For I = 1 To RsMst_UserRole.RecordCount
    CmbUserRole.AddItem RsMst_UserRole.Fields(1)
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
If SaveCmd.Enabled = False Then Exit Sub

If RsMenuUserRole.RecordCount <= 0 Then
   Exit Sub
End If
RsMenuUserRole.MoveFirst
RsMenuUserRole.Find "EmployeeNo = '" & TxtEmpID & "'"

If RsMenuUserRole.EOF Then
   Exit Sub
End If

Call Indata
Call EditCmd_Click
mUserRoleID = RsMenuUserRole!UserGRoupID
End Sub


