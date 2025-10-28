VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_MenuSearchPopup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7710
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11580
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Close 
      Caption         =   "Close"
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
      Left            =   5715
      TabIndex        =   3
      Top             =   75
      Width           =   975
   End
   Begin VB.TextBox TxtSearch 
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
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   4455
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "Search"
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
      Left            =   4620
      TabIndex        =   0
      Top             =   75
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7125
      Left            =   75
      TabIndex        =   2
      Top             =   510
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12568
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
Attribute VB_Name = "Frm_MenuSearchPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grid_Head()
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 5
   
   .TextMatrix(0, 0) = "Menu ID"
   .TextMatrix(0, 1) = "Menu Name"
   .TextMatrix(0, 2) = "Menu Path"
   .TextMatrix(0, 3) = "Form File Name"
   .TextMatrix(0, 4) = "Self Menu"
   
   .ColWidth(0) = 0
   .ColWidth(1) = 2500
   .ColWidth(2) = 8500
   .ColWidth(3) = 0
   .ColWidth(4) = 0
   
End With
End Sub

Private Sub Cmd_Close_Click()
Unload Me
End Sub

Private Sub CmdSearch_Click()
If TxtSearch.Text = "" Then Exit Sub

Call Grid_Head

Gen_DbPathFromAzure = "N"

tmp = " select b.MenuName, b.MenuID, b.FormName, b.selfmenu From  Mst_MenuHasUserRole a"
tmp = tmp & " inner join Mst_MenuDetail b on a.MenuID = b.MenuID"
tmp = tmp & " Where b.Flag = 'Y' and a.FormAccess = 'Y' and a.UserGroupID = " & Gen_UserTypeID & " and b.menuname like '%" & TxtSearch.Text & "%'"
Call Tmp6Table



If TmpRs6.RecordCount > 0 Then
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs6.RecordCount
      For I = 1 To TmpRs6.RecordCount
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs6!MenuID), 0, TmpRs6!MenuID)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs6!MenuName), "", TmpRs6!MenuName)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs6!FormName), "", TmpRs6!FormName)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs6!SelfMenu), "", TmpRs6!SelfMenu)
         Call GetPathforMenu(I)
         
         TmpRs6.MoveNext
      Next
   End With
Else
   MsgBox "No record found!!!"
End If

Gen_DbPathFromAzure = "Y"

End Sub

Private Sub GetPathforMenu(x_ As Variant)
mMenuID_ = Val(MSHFlexGrid1.TextMatrix(x_, 0))

mMasterLabel:

tmp = "select mastermenuid, menuname from mst_menudetail where  menuid = " & mMenuID_
Call Tmp7Table
If MSHFlexGrid1.TextMatrix(x_, 2) = "" Then
   MSHFlexGrid1.TextMatrix(x_, 2) = IIf(IsNull(TmpRs7!MenuName), "", TmpRs7!MenuName)
Else
   MSHFlexGrid1.TextMatrix(x_, 2) = IIf(IsNull(TmpRs7!MenuName), "", TmpRs7!MenuName) & "-->" & MSHFlexGrid1.TextMatrix(x_, 2)
End If
If IIf(IsNull(TmpRs7!MasterMenuID), 0, TmpRs7!MasterMenuID) <> 0 Then
   mMenuID_ = TmpRs7!MasterMenuID
   GoTo mMasterLabel
End If
End Sub

Private Sub Form_Load()
Call Grid_Head
End Sub

Private Sub Form_Resize()
Me.Top = (MDIFrmMain.Height - Me.Height) / 2
Me.Left = (MDIFrmMain.Width - Me.Width) / 2
End Sub

Private Sub MSHFlexGrid1_Click()
If MSHFlexGrid1.Rows = 2 Then Exit Sub
If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 4) = 0 Then
   Dim frm As Form
   Set frm = CallByName(Forms, "Add", VbMethod, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 3))
   frm.Caption = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 2)
   frm.Show
   If Gen_FormAccess = True Then
      Call TableLog_FormAccessDetail(" Where LogID = 0 ")
      RsLog_FormAccessDetail.AddNew
'      RsLog_FormAccessDetail!LogID = GetMaxLogID
      RsLog_FormAccessDetail!UserId = Gen_UserID
      RsLog_FormAccessDetail!MenuID = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 0)
      RsLog_FormAccessDetail.Update
   End If
End If
'Unload Me
Me.Hide
End Sub

