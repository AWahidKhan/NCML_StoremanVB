VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_MenuHasUserRole 
   Caption         =   "Mst_MenuHasUserRole"
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
      Begin VB.TextBox TxtMenuID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   3600
      End
      Begin VB.TextBox TxtUserGroupID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Top             =   615
         Width           =   3600
      End
      Begin VB.TextBox TxtFormAccess 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   18
         Top             =   990
         Width           =   3600
      End
      Begin VB.TextBox TxtADD_ 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1365
         Width           =   3600
      End
      Begin VB.TextBox TxtEDIT_ 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   22
         Top             =   1740
         Width           =   3600
      End
      Begin VB.TextBox TxtSave_ 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   24
         Top             =   2115
         Width           =   3600
      End
      Begin VB.TextBox TxtDelete_ 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   26
         Top             =   2490
         Width           =   3600
      End
      Begin VB.TextBox TxtCreatedBy 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   28
         Top             =   2865
         Width           =   3600
      End
      Begin MSComCtl2.DTPicker CreatedDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   30
         Top             =   3240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   59703297
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker UpdatedDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   32
         Top             =   3615
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   59703297
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtUpdatedBy 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   34
         Top             =   3990
         Width           =   3600
      End
      Begin VB.Label LblMenuID 
         Caption         =   "Menu ID"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblUserGroupID 
         Caption         =   "User Group ID"
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblFormAccess 
         Caption         =   "Form Access"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label LblADD_ 
         Caption         =   "ADD"
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   1365
         Width           =   1800
      End
      Begin VB.Label LblEDIT_ 
         Caption         =   "EDIT"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label LblSave_ 
         Caption         =   "Save"
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   2115
         Width           =   1800
      End
      Begin VB.Label LblDelete_ 
         Caption         =   "Delete"
         Height          =   300
         Left            =   240
         TabIndex        =   25
         Top             =   2490
         Width           =   1800
      End
      Begin VB.Label LblCreatedBy 
         Caption         =   "Created By"
         Height          =   300
         Left            =   240
         TabIndex        =   27
         Top             =   2865
         Width           =   1800
      End
      Begin VB.Label LblCreatedDate 
         Caption         =   "Created Date"
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   1800
      End
      Begin VB.Label LblUpdatedDate 
         Caption         =   "Updated Date"
         Height          =   300
         Left            =   240
         TabIndex        =   31
         Top             =   3615
         Width           =   1800
      End
      Begin VB.Label LblUpdatedBy 
         Caption         =   "Updated By"
         Height          =   300
         Left            =   240
         TabIndex        =   33
         Top             =   3990
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   35
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
         TabIndex        =   36
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
         TabIndex        =   37
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
Attribute VB_Name = "FrmMst_MenuHasUserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub Form_Load()
Call TableMst_MenuHasUserRole
If RsMst_MenuHasUserRole.RecordCount = 0 Then
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
Call Gen_Create_Xls("Mst_MenuHasUserRole.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Mst_MenuHasUserRole.xls")
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
RsMst_MenuHasUserRole.MoveNext
If RsMst_MenuHasUserRole.EOF Then
   RsMst_MenuHasUserRole.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_MenuHasUserRole.MovePrevious
If RsMst_MenuHasUserRole.BOF Then
   RsMst_MenuHasUserRole.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_MenuHasUserRole.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_MenuHasUserRole.MoveLast
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
RsMst_MenuHasUserRole.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_MenuHasUserRole.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsMst_MenuHasUserRole.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsMst_MenuHasUserRole.Fields(C)), "", RsMst_MenuHasUserRole.Fields(C))
    Next
    RsMst_MenuHasUserRole.MoveNext
Next
End Sub
Public Sub InData()

TxtMenuID.Locked = True
TxtUserGroupID.Locked = True
TxtFormAccess.Locked = True
TxtADD_.Locked = True
TxtEDIT_.Locked = True
TxtSave_.Locked = True
TxtDelete_.Locked = True
TxtCreatedBy.Locked = True
TxtCreatedDate.Locked = True
TxtUpdatedDate.Locked = True
TxtUpdatedBy.Locked = True

TxtMenuID = IIf(IsNull(RsMst_MenuHasUserRole!MenuID), "", RsMst_MenuHasUserRole!MenuID)
TxtUserGroupID = IIf(IsNull(RsMst_MenuHasUserRole!UserGroupID), "", RsMst_MenuHasUserRole!UserGroupID)
TxtFormAccess = IIf(IsNull(RsMst_MenuHasUserRole!FormAccess), "", RsMst_MenuHasUserRole!FormAccess)
TxtADD_ = IIf(IsNull(RsMst_MenuHasUserRole!ADD_), "", RsMst_MenuHasUserRole!ADD_)
TxtEDIT_ = IIf(IsNull(RsMst_MenuHasUserRole!EDIT_), "", RsMst_MenuHasUserRole!EDIT_)
TxtSave_ = IIf(IsNull(RsMst_MenuHasUserRole!Save_), "", RsMst_MenuHasUserRole!Save_)
TxtDelete_ = IIf(IsNull(RsMst_MenuHasUserRole!Delete_), "", RsMst_MenuHasUserRole!Delete_)
TxtCreatedBy = IIf(IsNull(RsMst_MenuHasUserRole!CreatedBy), "", RsMst_MenuHasUserRole!CreatedBy)
TxtCreatedDate = IIf(IsNull(RsMst_MenuHasUserRole!CreatedDate), "", RsMst_MenuHasUserRole!CreatedDate)
TxtUpdatedDate = IIf(IsNull(RsMst_MenuHasUserRole!UpdatedDate), "", RsMst_MenuHasUserRole!UpdatedDate)
TxtUpdatedBy = IIf(IsNull(RsMst_MenuHasUserRole!UpdatedBy), "", RsMst_MenuHasUserRole!UpdatedBy)
If RsMst_MenuHasUserRole.RecordCount = 1 Then
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

TxtMenuID.Locked = True
TxtUserGroupID.Locked = False
TxtFormAccess.Locked = False
TxtADD_.Locked = False
TxtEDIT_.Locked = False
TxtSave_.Locked = False
TxtDelete_.Locked = False
TxtCreatedBy.Locked = False
TxtCreatedDate.Locked = False
TxtUpdatedDate.Locked = False
TxtUpdatedBy.Locked = False

TxtMenuID = ""

TxtUserGroupID = ""

TxtFormAccess = ""

TxtADD_ = ""

TxtEDIT_ = ""

TxtSave_ = ""

TxtDelete_ = ""

TxtCreatedBy = ""

TxtCreatedDate = ""

TxtUpdatedDate = ""

TxtUpdatedBy = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtMenuID.Locked = True
TxtUserGroupID.Locked = False
TxtFormAccess.Locked = False
TxtADD_.Locked = False
TxtEDIT_.Locked = False
TxtSave_.Locked = False
TxtDelete_.Locked = False
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
   RsMst_MenuHasUserRole.Delete
   RsMst_MenuHasUserRole.Update
   If RsMst_MenuHasUserRole.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_MenuHasUserRole.MoveNext
   If RsMst_MenuHasUserRole.EOF() Then
      RsMst_MenuHasUserRole.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_MenuHasUserRole.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If RsMst_MenuHasUserRole.RecordCount > 0 Then
      RsMst_MenuHasUserRole.MoveFirst
      RsMst_MenuHasUserRole.Find "= '" & Txt.Text & "'"
      If Not RsMst_MenuHasUserRole.EOF Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsMst_MenuHasUserRole.AddNew
      RsMst_MenuHasUserRole!MenuID = GetMenuID
      TxtMenuID.Text = RsMst_MenuHasUserRole!MenuID
Else
    RsMst_MenuHasUserRole.MoveFirst
    RsMst_MenuHasUserRole.Find "= '" & Txt.Text & "'"
    If Not RsMst_MenuHasUserRole.EOF Then
       If RsMst_MenuHasUserRole!MenuID <> TxtMenuID.Text Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsMst_MenuHasUserRole.MoveFirst
    RsMst_MenuHasUserRole.Find " MenuID= " & TxtMenuID.Text
End If

RsMst_MenuHasUserRole!MenuID = TxtMenuID.Text
RsMst_MenuHasUserRole!UserGroupID = TxtUserGroupID.Text
RsMst_MenuHasUserRole!FormAccess = TxtFormAccess.Text
RsMst_MenuHasUserRole!ADD_ = TxtADD_.Text
RsMst_MenuHasUserRole!EDIT_ = TxtEDIT_.Text
RsMst_MenuHasUserRole!Save_ = TxtSave_.Text
RsMst_MenuHasUserRole!Delete_ = TxtDelete_.Text
RsMst_MenuHasUserRole!CreatedBy = TxtCreatedBy.Text
RsMst_MenuHasUserRole!CreatedDate = TxtCreatedDate.Text
RsMst_MenuHasUserRole!UpdatedDate = TxtUpdatedDate.Text
RsMst_MenuHasUserRole!UpdatedBy = TxtUpdatedBy.Text
RsMst_MenuHasUserRole.Update
Call InData
End Sub
Private Function GetMenuID() As Double
Dim RetVal As Double
tmp = "Select max(MenuID) from Mst_MenuHasUserRole"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetMenuID = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.TextMatrix(0, 0) = "Menu ID"
MSHFlexGrid1.TextMatrix(0, 1) = "User Group ID"
MSHFlexGrid1.TextMatrix(0, 2) = "Form Access"
MSHFlexGrid1.TextMatrix(0, 3) = "ADD"
MSHFlexGrid1.TextMatrix(0, 4) = "EDIT"
MSHFlexGrid1.TextMatrix(0, 5) = "Save"
MSHFlexGrid1.TextMatrix(0, 6) = "Delete"
MSHFlexGrid1.TextMatrix(0, 7) = "Created By"
MSHFlexGrid1.TextMatrix(0, 8) = "Created Date"
MSHFlexGrid1.TextMatrix(0, 9) = "Updated Date"
MSHFlexGrid1.TextMatrix(0, 10) = "Updated By"
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_MenuHasUserRole.MoveFirst
RsMst_MenuHasUserRole.Find "MenuID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_MenuHasUserRole.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
