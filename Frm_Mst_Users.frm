VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_Users 
   Caption         =   "Mst_Users"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12495
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   4000
      Left            =   8000
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   80
         TabIndex        =   10
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   1
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   8490
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5865
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10000
      Begin VB.TextBox TxtUsersID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   3600
      End
      Begin VB.TextBox TxtEmployee_ID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   16
         Top             =   615
         Width           =   3600
      End
      Begin VB.TextBox TxtEmployee_Name 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   18
         Top             =   990
         Width           =   3600
      End
      Begin VB.TextBox TxtContactNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   20
         Top             =   1365
         Width           =   3600
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   22
         Top             =   1740
         Width           =   3600
      End
      Begin VB.TextBox TxtDepartment 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   24
         Top             =   2115
         Width           =   3600
      End
      Begin VB.TextBox TxtDesignation 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   26
         Top             =   2490
         Width           =   3600
      End
      Begin MSComCtl2.DTPicker DateOfJoin 
         Height          =   360
         Left            =   1800
         TabIndex        =   28
         Top             =   2865
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   59899905
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtPassword1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   30
         Top             =   3240
         Width           =   3600
      End
      Begin VB.TextBox TxtUserGroupID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   32
         Top             =   3615
         Width           =   3600
      End
      Begin VB.ComboBox CmbUserGroupID 
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   3615
         Width           =   4000
      End
      Begin VB.TextBox TxtCreatedBy 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   35
         Top             =   3990
         Width           =   3600
      End
      Begin MSComCtl2.DTPicker CreatedDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   37
         Top             =   4365
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   59899905
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker UpdatedDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   39
         Top             =   4740
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   59899905
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtUpdatedBy 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   41
         Top             =   5115
         Width           =   3600
      End
      Begin VB.Label LblUsersID 
         Caption         =   "Users ID"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblEmployee_ID 
         Caption         =   "Employee ID"
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblEmployee_Name 
         Caption         =   "Employee Name"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label LblContactNo 
         Caption         =   "Contact No"
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   1365
         Width           =   1800
      End
      Begin VB.Label LblAddress 
         Caption         =   "Address"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label LblDepartment 
         Caption         =   "Department"
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   2115
         Width           =   1800
      End
      Begin VB.Label LblDesignation 
         Caption         =   "Designation"
         Height          =   300
         Left            =   240
         TabIndex        =   25
         Top             =   2490
         Width           =   1800
      End
      Begin VB.Label LblDateOfJoin 
         Caption         =   "Date Of Join"
         Height          =   300
         Left            =   240
         TabIndex        =   27
         Top             =   2865
         Width           =   1800
      End
      Begin VB.Label LblPassword1 
         Caption         =   "Password"
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   1800
      End
      Begin VB.Label LblUserGroupID 
         Caption         =   "User Role "
         Height          =   300
         Left            =   240
         TabIndex        =   31
         Top             =   3615
         Width           =   1800
      End
      Begin VB.Label LblCreatedBy 
         Caption         =   "Created By"
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   3990
         Width           =   1800
      End
      Begin VB.Label LblCreatedDate 
         Caption         =   "Created Date"
         Height          =   300
         Left            =   240
         TabIndex        =   36
         Top             =   4365
         Width           =   1800
      End
      Begin VB.Label LblUpdatedDate 
         Caption         =   "Updated Date"
         Height          =   300
         Left            =   240
         TabIndex        =   38
         Top             =   4740
         Width           =   1800
      End
      Begin VB.Label LblUpdatedBy 
         Caption         =   "Updated By"
         Height          =   300
         Left            =   240
         TabIndex        =   40
         Top             =   5115
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   42
      Top             =   7200
      Width           =   10000
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7320
         TabIndex        =   43
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7320
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5520
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3720
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1920
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmMst_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mUserGroupID As Variant
Private Sub Form_Load()
Call TableMst_Users
If RsMst_Users.RecordCount = 0 Then
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
Call Gen_Create_Xls("Mst_Users.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Mst_Users.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Users.MoveNext
If RsMst_Users.EOF Then
   RsMst_Users.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Users.MovePrevious
If RsMst_Users.BOF Then
   RsMst_Users.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_Users.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Users.MoveLast
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
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 350
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_Users.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_Users.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsMst_Users.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsMst_Users.Fields(C)), "", RsMst_Users.Fields(C))
    Next
    RsMst_Users.MoveNext
Next
End Sub
Public Sub InData()

TxtUsersID.Locked = True
TxtEmployee_ID.Locked = True
TxtEmployee_Name.Locked = True
TxtContactNo.Locked = True
TxtAddress.Locked = True
TxtDepartment.Locked = True
TxtDesignation.Locked = True
TxtDateOfJoin.Locked = True
TxtPassword1.Locked = True
TxtUserGroupID.Locked = True
CmbUserGroupID.Visible = False
TxtUserGroupID.Visible = True
TxtCreatedBy.Locked = True
TxtCreatedDate.Locked = True
TxtUpdatedDate.Locked = True
TxtUpdatedBy.Locked = True

TxtUsersID = IIf(IsNull(RsMst_Users!UsersID), "", RsMst_Users!UsersID)
TxtEmployee_ID = IIf(IsNull(RsMst_Users!Employee_ID), "", RsMst_Users!Employee_ID)
TxtEmployee_Name = IIf(IsNull(RsMst_Users!Employee_Name), "", RsMst_Users!Employee_Name)
TxtContactNo = IIf(IsNull(RsMst_Users!ContactNo), "", RsMst_Users!ContactNo)
TxtAddress = IIf(IsNull(RsMst_Users!Address), "", RsMst_Users!Address)
TxtDepartment = IIf(IsNull(RsMst_Users!Department), "", RsMst_Users!Department)
TxtDesignation = IIf(IsNull(RsMst_Users!Designation), "", RsMst_Users!Designation)
TxtDateOfJoin = IIf(IsNull(RsMst_Users!DateOfJoin), "", RsMst_Users!DateOfJoin)
TxtPassword1 = IIf(IsNull(RsMst_Users!Password1), "", RsMst_Users!Password1)
TxtUserGroupID = IIf(IsNull(RsMst_Users!GroupName), "", RsMst_Users!GroupName)
mUserGroupID = IIf(IsNull(RsMst_Users!UserGroupID), "", RsMst_Users!UserGroupID)
TxtCreatedBy = IIf(IsNull(RsMst_Users!CreatedBy), "", RsMst_Users!CreatedBy)
TxtCreatedDate = IIf(IsNull(RsMst_Users!CreatedDate), "", RsMst_Users!CreatedDate)
TxtUpdatedDate = IIf(IsNull(RsMst_Users!UpdatedDate), "", RsMst_Users!UpdatedDate)
TxtUpdatedBy = IIf(IsNull(RsMst_Users!UpdatedBy), "", RsMst_Users!UpdatedBy)
If RsMst_Users.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtUsersID_LostFocus()
If TxtUsersID = "" Then Exit Sub
If IsNumeric(TxtUsersID) = False Then
   MsgBox "Invalid number format !!!!"
   TxtUsersID.SetFocus
   Exit Sub
End If
TxtUsersID = Val(TxtUsersID)
End Sub
Private Sub TxtDateOfJoin_LostFocus()
End Sub
Private Sub TxtCreatedDate_LostFocus()
End Sub
Private Sub TxtUpdatedDate_LostFocus()
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtUsersID.Locked = True
TxtEmployee_ID.Locked = False
TxtEmployee_Name.Locked = False
TxtContactNo.Locked = False
TxtAddress.Locked = False
TxtDepartment.Locked = False
TxtDesignation.Locked = False
TxtDateOfJoin.Locked = False
TxtPassword1.Locked = False
TxtUserGroupID.Locked = False
CmbUserGroupID.Visible = True
TxtUserGroupID.Visible = False
TxtCreatedBy.Locked = False
TxtCreatedDate.Locked = False
TxtUpdatedDate.Locked = False
TxtUpdatedBy.Locked = False

TxtUsersID = ""

TxtEmployee_ID = ""

TxtEmployee_Name = ""

TxtContactNo = ""

TxtAddress = ""

TxtDepartment = ""

TxtDesignation = ""

TxtDateOfJoin = ""

TxtPassword1 = ""

TxtUserGroupID = ""

TxtCreatedBy = ""

TxtCreatedDate = ""

TxtUpdatedDate = ""

TxtUpdatedBy = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtUsersID.Locked = True
TxtEmployee_ID.Locked = False
TxtEmployee_Name.Locked = False
TxtContactNo.Locked = False
TxtAddress.Locked = False
TxtDepartment.Locked = False
TxtDesignation.Locked = False
TxtDateOfJoin.Locked = False
TxtPassword1.Locked = False
TxtUserGroupID.Locked = False
CmbUserGroupID.Visible = True
TxtUserGroupID.Visible = False
TxtCreatedBy.Locked = False
TxtCreatedDate.Locked = False
TxtUpdatedDate.Locked = False
TxtUpdatedBy.Locked = False

CmbUserGroupID.Text = TxtUserGroupID
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsMst_Users.Delete
   RsMst_Users.Update
   If RsMst_Users.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Users.MoveNext
   If RsMst_Users.EOF() Then
      RsMst_Users.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_Users.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtEmployee_ID.Text = "" Then
   MsgBox "Blank Employee_ID Not Allowed !!! "
   TxtEmployee_ID.SetFocus
   Exit Sub
End If


If TxtEmployee_Name.Text = "" Then
   MsgBox "Blank Employee_Name Not Allowed !!! "
   TxtEmployee_Name.SetFocus
   Exit Sub
End If


If TxtPassword1.Text = "" Then
   MsgBox "Blank Password1 Not Allowed !!! "
   TxtPassword1.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Users.RecordCount > 0 Then
      RsMst_Users.MoveFirst
      RsMst_Users.Find "Employee_Name= '" & TxtEmployee_Name.Text & "'"
      If Not RsMst_Users.EOF Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         TxtEmployee_Name.SetFocus
         Exit Sub
      End If
   End If
      RsMst_Users.AddNew
      RsMst_Users!UsersID = GetUsersID
      TxtUsersID.Text = RsMst_Users!UsersID
Else
    RsMst_Users.MoveFirst
    RsMst_Users.Find "Employee_Name= '" & TxtEmployee_Name.Text & "'"
    If Not RsMst_Users.EOF Then
       If RsMst_Users!UsersID <> TxtUsersID.Text Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          TxtEmployee_Name.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Users.MoveFirst
    RsMst_Users.Find " UsersID= " & TxtUsersID.Text
End If

RsMst_Users!UsersID = TxtUsersID.Text
RsMst_Users!Employee_ID = TxtEmployee_ID.Text
RsMst_Users!Employee_Name = TxtEmployee_Name.Text
RsMst_Users!ContactNo = TxtContactNo.Text
RsMst_Users!Address = TxtAddress.Text
RsMst_Users!Department = TxtDepartment.Text
RsMst_Users!Designation = TxtDesignation.Text
RsMst_Users!DateOfJoin = TxtDateOfJoin.Text
RsMst_Users!Password1 = TxtPassword1.Text
RsMst_Users!UserGroupID = mUserGroupID
RsMst_Users!CreatedBy = TxtCreatedBy.Text
RsMst_Users!CreatedDate = TxtCreatedDate.Text
RsMst_Users!UpdatedDate = TxtUpdatedDate.Text
RsMst_Users!UpdatedBy = TxtUpdatedBy.Text
RsMst_Users.Update
Call InData
End Sub
Private Function GetUsersID() As Double
Dim RetVal As Double
tmp = "Select max(UsersID) from Mst_Users"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetUsersID = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 14
MSHFlexGrid1.TextMatrix(0, 0) = "Users ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Employee ID"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Contact No"
MSHFlexGrid1.TextMatrix(0, 4) = "Address"
MSHFlexGrid1.TextMatrix(0, 5) = "Department"
MSHFlexGrid1.TextMatrix(0, 6) = "Designation"
MSHFlexGrid1.TextMatrix(0, 7) = "Date Of Join"
MSHFlexGrid1.TextMatrix(0, 8) = "Password"
MSHFlexGrid1.TextMatrix(0, 9) = "User Role "
MSHFlexGrid1.TextMatrix(0, 10) = "Created By"
MSHFlexGrid1.TextMatrix(0, 11) = "Created Date"
MSHFlexGrid1.TextMatrix(0, 12) = "Updated Date"
MSHFlexGrid1.TextMatrix(0, 13) = "Updated By"
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_Users.MoveFirst
RsMst_Users.Find "UsersID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_Users.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdUserGroupID_Click()
    Frm_Mst_UserRole.Show
End Sub
'----
Private Sub CmbUserGroupID_GotFocus()
tmp = CmbUserGroupID.Text
Call TableMst_UserRole
CmbUserGroupID.Clear
If RsMst_UserRole.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_UserRole.RecordCount
    CmbUserGroupID.AddItem RsMst_UserRole!GroupName
    RsMst_UserRole.MoveNext
Next
CmbUserGroupID.Text = tmp
End Sub
'----
Private Sub CmbUserGroupID_LostFocus()
If CmbUserGroupID.Text = "" Then
   Exit Sub
End If
RsMst_UserRole.MoveFirst
RsMst_UserRole.Find "GroupName = '" & CmbUserGroupID.Text & "'"
If RsMst_UserRole.EOF Then
   MsgBox "Invalid  selection "
   CmbUserGroupID.SetFocus
   Exit Sub
End If
mUserGroupID = RsMst_UserRole!UserGroupID
End Sub
'----
