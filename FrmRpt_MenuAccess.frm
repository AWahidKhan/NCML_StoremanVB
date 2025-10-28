VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_MenuAccess 
   Caption         =   "Menu Access Detail"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   105
      TabIndex        =   8
      Top             =   30
      Width           =   15000
      Begin VB.ComboBox CmbMenuType 
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
         ItemData        =   "FrmRpt_MenuAccess.frx":0000
         Left            =   5160
         List            =   "FrmRpt_MenuAccess.frx":0013
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   562
         Width           =   2895
      End
      Begin VB.ComboBox CmbUserRole 
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
         Left            =   8160
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   562
         Width           =   3015
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
         Height          =   360
         Left            =   12480
         TabIndex        =   5
         Top             =   562
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   11265
         TabIndex        =   4
         Top             =   562
         Width           =   1215
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   13695
         TabIndex        =   6
         Top             =   562
         Width           =   1215
      End
      Begin VB.ComboBox CmbMenuName 
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
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   562
         Width           =   4935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Menu Type"
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
         Left            =   6097
         TabIndex        =   14
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   9210
         TabIndex        =   12
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Record Found :-"
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
         Left            =   11280
         TabIndex        =   11
         Top             =   240
         Width           =   2280
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "--"
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
         Left            =   13680
         TabIndex        =   10
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   2100
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8955
      Left            =   105
      TabIndex        =   0
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         AllowUserResizing=   3
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   8565
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_MenuAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMenuID, mUserRoleID As Variant

Public Sub Grid_Head()
Label5.Caption = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Cols = 8
   .Rows = 2
      
   .TextMatrix(0, 0) = "Menu Name"
   .TextMatrix(0, 1) = "Menu Type"
   .TextMatrix(0, 2) = "User Role"
   .TextMatrix(0, 3) = "Access"
   .TextMatrix(0, 4) = "Add"
   .TextMatrix(0, 5) = "Edit"
   .TextMatrix(0, 6) = "Delete"
   .TextMatrix(0, 7) = "Description"
   
   .ColWidth(0) = 3500
   .ColWidth(1) = 1500
   .ColWidth(2) = 1500
   .ColWidth(3) = 900
   .ColWidth(4) = 900
   .ColWidth(5) = 900
   .ColWidth(6) = 900
   .ColWidth(7) = 4000
   
   .RowHeight(0) = 700
   .WordWrap = True
End With
End Sub

Private Sub CmbMenuName_GotFocus()
tmp = CmbMenuName.Text
Call TableMst_MenuDetail

CmbMenuName.Clear
If RsMst_MenuDetail.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_MenuDetail.RecordCount
   CmbMenuName.AddItem RsMst_MenuDetail!MenuName
   RsMst_MenuDetail.MoveNext
Next
CmbMenuName.Text = tmp
End Sub

Private Sub CmbMenuName_LostFocus()
If Trim(CmbMenuName.Text) = "" Then Exit Sub

RsMst_MenuDetail.MoveFirst
RsMst_MenuDetail.Find "MenuName = '" & CmbMenuName.Text & "'"
If RsMst_MenuDetail.EOF Then
   MsgBox "Invalid selection "
   CmbMenuName.SetFocus
   Exit Sub
End If
mMenuID = RsMst_MenuDetail!MenuID
End Sub

Private Sub CmbMenuType_LostFocus()
If Trim(CmbMenuType.Text) = "" Then Exit Sub
If LCase(CmbMenuType.Text) <> "admin" And LCase(CmbMenuType.Text) <> "master" And LCase(CmbMenuType.Text) <> "transaction" And LCase(CmbMenuType.Text) <> "report" And LCase(CmbMenuType.Text) <> "cog update" Then
   MsgBox "Invalid Selection!!!"
   CmbMenuType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbUserRole_GotFocus()
tmp = CmbUserRole.Text

Call TableMst_UserRole

CmbUserRole.Clear
If RsMst_UserRole.RecordCount = 0 Then
   MsgBox "No User Role found"
   CmbUserRole.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_UserRole.RecordCount
   CmbUserRole.AddItem RsMst_UserRole!groupname
   RsMst_UserRole.MoveNext
Next
CmbUserRole.Text = tmp
End Sub

Private Sub CmbUserRole_LostFocus()
If Trim(CmbUserRole.Text) = "" Then Exit Sub

RsMst_UserRole.MoveFirst
RsMst_UserRole.Find "GroupName = '" & CmbUserRole.Text & "'"
If RsMst_UserRole.EOF Then
   MsgBox "Invalid selection!!!"
   CmbUserRole.SetFocus
   Exit Sub
End If

mUserRoleID = RsMst_UserRole!UserGRoupID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbUserRole.Text = ""
CmbMenuName.Text = ""
CmbMenuType.Text = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_MenuAccess.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_MenuAccess.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
       oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
Call Grid_Head
MsgBox "Please wait till next Message!!!"

If CmbUserRole.Text <> "" Then
   wc = " Where a.usergroupid = " & mUserRoleID
End If
If CmbMenuName.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And a.menuid = " & mMenuID
   Else
      wc = " Where a.menuid = " & mMenuID
   End If
End If
If CmbMenuType.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And b.hmenuflag = '" & Left$(CmbMenuType.Text, 1) & "'"
   Else
      wc = " Where b.hmenuflag = '" & Left$(CmbMenuType.Text, 1) & "'"
   End If
End If

tmp = " select b.menuname,  case b.hmenuflag  when 'a' then 'Admin' when 'c' then 'COG Update' when 'm' then 'Master'"
tmp = tmp & " when 'r' then 'Report' when 't' then 'Transaction' end as 'menutype', c.groupname,"
tmp = tmp & " case a.formaccess when 'y' then 'Yes' when 'n' then 'No' end as 'formaccess',"
tmp = tmp & " case a.add_ when 'y' then 'Yes' when 'n' then 'No' end as 'add',"
tmp = tmp & " case a.edit_ when 'y' then 'Yes' when 'n' then 'No' end as 'edit',"
tmp = tmp & " case a.delete_ when 'y' then 'Yes' when 'n' then 'No' end as 'delete',"
tmp = tmp & " b.detaildesc"
tmp = tmp & " from Mst_MenuHasUserRole a"
tmp = tmp & " inner join Mst_MenuDetail b on  a.menuid = b.menuid"
tmp = tmp & " inner join mst_userrole c on a.usergroupid = c.usergroupid"
tmp = tmp & wc
tmp = tmp & " order by b.menuname, b.hmenuflag, c.groupname"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 3
   CmdExcel.Enabled = True
   Label5.Caption = TmpRs.RecordCount
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For I = 1 To TmpRs.RecordCount
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!MenuName), "", TmpRs!MenuName)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!MenuType), "", TmpRs!MenuType)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!groupname), "", TmpRs!groupname)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!formaccess), "", TmpRs!formaccess)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!Add), "", TmpRs!Add)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!Edit), "", TmpRs!Edit)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!Delete), "", TmpRs!Delete)
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!detaildesc), "", TmpRs!detaildesc)
         .Row = I
         .Col = 0
         .CellAlignment = flexAlignLeftCenter
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
   MsgBox "No record found!!!"
End If

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Call CmdClear_Click
End Sub

