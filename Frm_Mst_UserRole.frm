VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_UserRole 
   Caption         =   "Mst_UserRole"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12495
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1125
      Left            =   9960
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
      Height          =   2625
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10005
      Begin VB.TextBox TxtUserGroupID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   1320
      End
      Begin VB.TextBox TxtGroupName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   16
         Top             =   675
         Width           =   6720
      End
      Begin VB.TextBox TxtCreatedBy 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   18
         Top             =   1155
         Width           =   3600
      End
      Begin MSComCtl2.DTPicker CreatedDate 
         Height          =   360
         Left            =   7200
         TabIndex        =   20
         Top             =   1170
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   48693249
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker UpdatedDate 
         Height          =   360
         Left            =   7200
         TabIndex        =   22
         Top             =   1740
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   48693249
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtUpdatedBy 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   24
         Top             =   1635
         Width           =   3600
      End
      Begin VB.Label LblUserGroupID 
         Caption         =   "User Group ID"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblGroupName 
         Caption         =   "Group Name"
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   675
         Width           =   1800
      End
      Begin VB.Label LblCreatedBy 
         Caption         =   "Created By"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   1155
         Width           =   1800
      End
      Begin VB.Label LblCreatedDate 
         Caption         =   "Created Date"
         Height          =   300
         Left            =   5640
         TabIndex        =   19
         Top             =   1170
         Width           =   1800
      End
      Begin VB.Label LblUpdatedDate 
         Caption         =   "Updated Date"
         Height          =   300
         Left            =   5640
         TabIndex        =   21
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label LblUpdatedBy 
         Caption         =   "Updated By"
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   1635
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   10000
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   7920
         TabIndex        =   26
         Top             =   660
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   7920
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   6120
         TabIndex        =   2
         Top             =   660
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4200
         TabIndex        =   4
         Top             =   660
         Width           =   1935
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2160
         TabIndex        =   6
         Top             =   660
         Width           =   2055
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   8
         Top             =   660
         Width           =   2055
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmMst_UserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub Form_Load()
Call TableMst_UserRole
If RsMst_UserRole.RecordCount = 0 Then
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
Call Gen_Create_Xls("Mst_UserRole.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Mst_UserRole.xls")
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
RsMst_UserRole.MoveNext
If RsMst_UserRole.EOF Then
   RsMst_UserRole.MoveLast
   LF_Flag = "Y"
End If
Call Indata
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
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_UserRole.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_UserRole.MoveLast
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
RsMst_UserRole.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_UserRole.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsMst_UserRole.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsMst_UserRole.Fields(C)), "", RsMst_UserRole.Fields(C))
    Next
    RsMst_UserRole.MoveNext
Next
End Sub
Public Sub Indata()

TxtUserGroupID.Locked = True
TxtGroupName.Locked = True
TxtCreatedBy.Locked = True
TxtCreatedDate.Locked = True
TxtUpdatedDate.Locked = True
TxtUpdatedBy.Locked = True

TxtUserGroupID = IIf(IsNull(RsMst_UserRole!UserGroupID), "", RsMst_UserRole!UserGroupID)
TxtGroupName = IIf(IsNull(RsMst_UserRole!GroupName), "", RsMst_UserRole!GroupName)
TxtCreatedBy = IIf(IsNull(RsMst_UserRole!CreatedBy), "", RsMst_UserRole!CreatedBy)
TxtCreatedDate = IIf(IsNull(RsMst_UserRole!CreatedDate), "", RsMst_UserRole!CreatedDate)
TxtUpdatedDate = IIf(IsNull(RsMst_UserRole!UpdatedDate), "", RsMst_UserRole!UpdatedDate)
TxtUpdatedBy = IIf(IsNull(RsMst_UserRole!UpdatedBy), "", RsMst_UserRole!UpdatedBy)
If RsMst_UserRole.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtCreatedDate_LostFocus()
End Sub
Private Sub TxtUpdatedDate_LostFocus()
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtUserGroupID.Locked = True
TxtGroupName.Locked = False
TxtCreatedBy.Locked = False
TxtCreatedDate.Locked = False
TxtUpdatedDate.Locked = False
TxtUpdatedBy.Locked = False

TxtUserGroupID = ""
TxtGroupName = ""
TxtCreatedBy = ""
TxtCreatedDate = ""
TxtUpdatedDate = ""
TxtUpdatedBy = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtUserGroupID.Locked = True
TxtGroupName.Locked = False
TxtCreatedBy.Locked = False
TxtCreatedDate.Locked = False
TxtUpdatedDate.Locked = False
TxtUpdatedBy.Locked = False

End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsMst_UserRole.Delete
   RsMst_UserRole.Update
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
MsgBox "Child Record Present "
RsMst_UserRole.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If RsMst_UserRole.RecordCount > 0 Then
      RsMst_UserRole.MoveFirst
      RsMst_UserRole.Find "= '" & Txt.Text & "'"
      If Not RsMst_UserRole.EOF Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsMst_UserRole.AddNew
      RsMst_UserRole!UserGroupID = GetUserGroupID
      TxtUserGroupID.Text = RsMst_UserRole!UserGroupID
Else
    RsMst_UserRole.MoveFirst
    RsMst_UserRole.Find "= '" & Txt.Text & "'"
    If Not RsMst_UserRole.EOF Then
       If RsMst_UserRole!UserGroupID <> TxtUserGroupID.Text Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsMst_UserRole.MoveFirst
    RsMst_UserRole.Find " UserGroupID= " & TxtUserGroupID.Text
End If

RsMst_UserRole!UserGroupID = TxtUserGroupID.Text
RsMst_UserRole!GroupName = TxtGroupName.Text
RsMst_UserRole!CreatedBy = TxtCreatedBy.Text
RsMst_UserRole!CreatedDate = TxtCreatedDate.Text
RsMst_UserRole!UpdatedDate = TxtUpdatedDate.Text
RsMst_UserRole!UpdatedBy = TxtUpdatedBy.Text
RsMst_UserRole.Update
Call Indata
End Sub
Private Function GetUserGroupID() As Double
Dim RetVal As Double
tmp = "Select max(UserGroupID) from Mst_UserRole"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetUserGroupID = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 6
MSHFlexGrid1.TextMatrix(0, 0) = "User Group ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Group Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Created By"
MSHFlexGrid1.TextMatrix(0, 3) = "Created Date"
MSHFlexGrid1.TextMatrix(0, 4) = "Updated Date"
MSHFlexGrid1.TextMatrix(0, 5) = "Updated By"
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_UserRole.MoveFirst
RsMst_UserRole.Find "UserGroupID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_UserRole.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
