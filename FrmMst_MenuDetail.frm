VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_MenuDetail 
   Caption         =   "Menu Detail"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   9225
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   5655
      Left            =   240
      TabIndex        =   12
      Top             =   9360
      Visible         =   0   'False
      Width           =   12015
      Begin VB.OptionButton Option3 
         Caption         =   "Search By Form Description"
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
         Left            =   5520
         TabIndex        =   43
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   9300
         TabIndex        =   40
         Top             =   765
         Width           =   1000
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Search By Menu Name"
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
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Search By Form Name"
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
         TabIndex        =   38
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxtSearchForm 
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
         Left            =   2580
         TabIndex        =   32
         Top             =   750
         Width           =   5595
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   8295
         TabIndex        =   31
         Top             =   765
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   75
         TabIndex        =   13
         Top             =   1275
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Form Name"
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
         Left            =   75
         TabIndex        =   33
         Top             =   795
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Access Details"
      Height          =   9120
      Left            =   10560
      TabIndex        =   44
      Top             =   0
      Width           =   4590
      Begin VB.CommandButton CmdAccessRightMenu 
         Caption         =   "Go to Access Right Menu Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   105
         TabIndex        =   46
         Top             =   8340
         Width           =   3135
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   7860
         Left            =   75
         TabIndex        =   45
         Top             =   225
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   13864
         _Version        =   393216
         Rows            =   3
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   10410
      Begin VB.TextBox TxtDetailDesc 
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
         Height          =   1875
         Left            =   2400
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   6015
         Width           =   7815
      End
      Begin VB.TextBox TxtFormDesc 
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
         Height          =   1395
         Left            =   2400
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4560
         Width           =   7815
      End
      Begin VB.TextBox TxtSequenceNo 
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
         TabIndex        =   8
         Top             =   4110
         Width           =   3975
      End
      Begin VB.ComboBox CmbMasterMenu 
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
         ItemData        =   "FrmMst_MenuDetail.frx":0000
         Left            =   2400
         List            =   "FrmMst_MenuDetail.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   3600
         Width           =   3975
      End
      Begin VB.TextBox TxtMasterMenu 
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
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3600
         Width           =   3975
      End
      Begin VB.CheckBox ChkSelfMenu 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   3120
         Width           =   195
      End
      Begin VB.TextBox TxtHMenuFlag 
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
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox TxtFlag 
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
         MaxLength       =   1
         TabIndex        =   4
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox TxtFormName 
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
         TabIndex        =   3
         Top             =   1680
         Width           =   7815
      End
      Begin VB.TextBox TxtMenuName 
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
         TabIndex        =   2
         Top             =   1200
         Width           =   7815
      End
      Begin VB.TextBox TxtFormFormName 
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
         TabIndex        =   1
         Top             =   720
         Width           =   7815
      End
      Begin VB.TextBox TxtMenuID 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Detail Description"
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
         TabIndex        =   42
         Top             =   6825
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Form Description"
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
         TabIndex        =   41
         Top             =   5130
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Sequence No"
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
         Left            =   240
         TabIndex        =   37
         Top             =   4170
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Master Menu"
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
         TabIndex        =   36
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Self Menu"
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
         Left            =   240
         TabIndex        =   34
         Top             =   3187
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Header Menu"
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
         TabIndex        =   30
         Top             =   2700
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Flag"
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
         TabIndex        =   29
         Top             =   2220
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Form Name"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1747
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Menu Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   1282
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Menu ID"
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
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Form File Name"
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
         Left            =   240
         TabIndex        =   25
         Top             =   780
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   8040
      Width           =   10440
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2295
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4230
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6165
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   8100
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2295
         TabIndex        =   18
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4230
         TabIndex        =   17
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   6165
         TabIndex        =   16
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   8100
         TabIndex        =   15
         Top             =   615
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmMst_MenuDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mMasterMenuID As Variant
Dim mSearch As Boolean

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock_Pwd(Me, False)
CmdAccessRightMenu.Enabled = False
TxtMenuID = ""
TxtFormFormName = ""
TxtMenuName = ""
TxtFormName = ""
TxtFlag = ""
TxtHMenuFlag = ""
CmbMasterMenu.Visible = True
TxtMasterMenu.Visible = False
CmbMasterMenu.Text = ""
ChkSelfMenu.Value = 0
TxtSequenceNo = ""
TxtFormDesc = ""
TxtDetailDesc = ""

'Frame1.Enabled = True

TxtFormDesc.Locked = False
TxtDetailDesc.Locked = False
TxtFlag.Locked = False
TxtFormFormName.Locked = False
TxtFormName.Locked = False
TxtHMenuFlag.Locked = False
TxtMasterMenu.Locked = False
TxtMenuName.Locked = False
TxtSequenceNo.Locked = False
ChkSelfMenu.Enabled = True
If RsMenuDetail.RecordCount > 0 Then
   TxtFormFormName.SetFocus
Else
   SearchCmd.Enabled = False
End If

Call Grid_HeadAccess

End Sub
Private Sub ShowAccessDetails()

Call Grid_HeadAccess
CmdAccessRightMenu.Enabled = False
If ChkSelfMenu.Value = 0 Then
   CmdAccessRightMenu.Enabled = True
   
   tmp = " Select Distinct MUR.GroupName,MHUR.FormAccess,MHUR.ADD_,MHUR.EDIT_,MHUR.Delete_"
   tmp = tmp & " From Mst_MenuHasUserRole MHUR"
   tmp = tmp & " Left Join Mst_MenuDetail MM On MHUR.MenuID = MM.MenuID"
   tmp = tmp & " Left Join Mst_UserRole MUR On MHUR.UserGroupID = MUR.UserGroupID"
   tmp = tmp & " Where MM.MenuID = " & TxtMenuID & ""
   
   Call TmpTable
   
   For i = 1 To TmpRs.RecordCount
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       MSHFlexGrid2.TextMatrix(i, 0) = IIf(IsNull(TmpRs!groupname), "N", TmpRs!groupname)
       MSHFlexGrid2.TextMatrix(i, 1) = IIf(IsNull(TmpRs!formaccess), "N", TmpRs!formaccess)
       MSHFlexGrid2.TextMatrix(i, 2) = IIf(IsNull(TmpRs!ADD_), "N", TmpRs!ADD_)
       MSHFlexGrid2.TextMatrix(i, 3) = IIf(IsNull(TmpRs!EDIT_), "N", TmpRs!EDIT_)
       MSHFlexGrid2.TextMatrix(i, 4) = IIf(IsNull(TmpRs!Delete_), "N", TmpRs!Delete_)
       TmpRs.MoveNext
   Next

End If

End Sub
Private Sub CmbMasterMenu_GotFocus()

tmp1 = CmbMasterMenu.Text
'Call TableMst_Designation("")

tmp = "Select * from Mst_MenuDetail Where SelfMenu = 1 "
Call TmpTable
CmbMasterMenu.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Menu found"
   CmbMasterMenu.SetFocus
   Exit Sub
End If
For i = 1 To TmpRs.RecordCount
    CmbMasterMenu.AddItem TmpRs!MenuName
    TmpRs.MoveNext
Next
CmbMasterMenu.Text = tmp1

End Sub

Private Sub CmbMasterMenu_LostFocus()
If CmbMasterMenu.Text = "" Then
   TxtSequenceNo = ""
   mMasterMenuID = Null
   Exit Sub
End If

TmpRs.MoveFirst
TmpRs.Find "MenuName = '" & CmbMasterMenu.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Menu Selection "
   CmbMasterMenu.SetFocus
   Exit Sub
End If
mMasterMenuID = TmpRs!MenuID

tmp = "Select Max(SequenceNo) as Seq from Mst_MenuDetail Where MasterMenuID = " & mMasterMenuID & ""
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   TxtSequenceNo = IIf(IsNull(TmpRs1!Seq), 0, TmpRs1!Seq) + 1
Else
   TxtSequenceNo = 1
End If


End Sub

Private Sub CmdAccessRightMenu_Click()
FrmAccessRightMenu.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_MenuDetail.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_MenuDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

Dim wc As Variant
wc = ""
'tmp = "Select MenuID,FormFileName from Mst_MenuDetail "
Call Grid_Head

tmp = "Select MM.MenuID,MM.FormFileName,MM.MenuName,MM.FormName,MM.Flag,MM.SelfMenu,MM1.MenuName as 'MasterMenu',MM.HMenuFlag,MM.SequenceNo,MM.FormDesc,MM.DetailDesc From Mst_MenuDetail MM"
tmp = tmp & " Left Join Mst_MenuDetail MM1 On MM.MasterMenuID = MM1.MenuID"
'tmp = tmp & " Where MM.SelfMenu = 1 Or MM1.MenuName is not null "
'tmp = tmp & " Order By MM.SelfMenu desc,MM1.MasterMenu,MM.SequenceNo "
'tmp = tmp & " Order By  MM.SelfMenu Desc, MM.MasterMenuID, MM1.MasterMenu, MM.SequenceNo"

If TxtSearchForm <> "" Then
   If Option1.Value = True Then
      wc = " Where MM.FormFileName Like '%" & TxtSearchForm & "%' "
   ElseIf Option2.Value = True Then
      wc = " Where MM.MenuName Like '%" & TxtSearchForm & "%' "
   Else
      wc = " Where MM.FormDesc Like '%" & TxtSearchForm & "%' "
   End If
End If

tmp = tmp & wc
tmp = tmp & "  Order by MM.MenuID "

Call TmpTable

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "r" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Report"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "m" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Master"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "t" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Transaction"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "c" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "COG Process"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 7)) = "a" Then
       MSHFlexGrid1.TextMatrix(iu, 7) = "Admin"
    End If
    TmpRs.MoveNext
Next

End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   RsMenuDetail.Delete
   RsMenuDetail.Update
   'MsgBox "Record Successfully Removed"
   If RsMenuDetail.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMenuDetail.MoveNext
   If RsMenuDetail.EOF() Then
      RsMenuDetail.MoveLast
   End If
   Call InData
End If

Exit Sub

msgError:
MsgBox "Child record Present "
RsMenuDetail.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
'Frame1.Enabled = True

TxtFormDesc.Locked = False
TxtDetailDesc.Locked = False
TxtFlag.Locked = False
TxtFormFormName.Locked = False
TxtFormName.Locked = False
TxtHMenuFlag.Locked = False
TxtMasterMenu.Locked = False
TxtMenuName.Locked = False
TxtSequenceNo.Locked = False
ChkSelfMenu.Enabled = True
CmbMasterMenu.Visible = True
TxtMasterMenu.Visible = False
CmbMasterMenu.Text = TxtMasterMenu
CmdAccessRightMenu.Enabled = False
Call GButtonLock_Pwd(Me, False)
'TxtMenuID.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMenuDetail.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True

End Sub

Private Sub Form_Load()
'Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   Call GButtonLock_PwdAD(Me)
'   MsgBox "Access denied !!!!!!!!!"
'   Exit Sub
'End If
If UserType = "Power" Then

Else
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
'Call StatusBarData(Me)

Call Mst_MenuDetailTable("  Order by HmenuFlag,MenuID ")
If RsMenuDetail.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call InData


Call GButtonLock_Pwd(Me, True)
mSearch = False



End Sub

Public Sub InData()

'TxtMenuID.Locked = True

TxtMenuID = RsMenuDetail!MenuID
TxtFormFormName = RsMenuDetail!FormFileName

TxtMenuName = IIf(IsNull(RsMenuDetail!MenuName), "", RsMenuDetail!MenuName)

TxtFormName = IIf(IsNull(RsMenuDetail!FormName), "", RsMenuDetail!FormName)
TxtFlag = IIf(IsNull(RsMenuDetail!Flag), "", RsMenuDetail!Flag)
TxtHMenuFlag = IIf(IsNull(RsMenuDetail!HMenuFlag), "", RsMenuDetail!HMenuFlag)

CmbMasterMenu.Visible = False
TxtMasterMenu.Visible = True
'CmbMasterMenu.Text = ""
'ChkSelfMenu.Value = 0

mMasterMenuID = IIf(IsNull(RsMenuDetail!MasterMenuID), 0, RsMenuDetail!MasterMenuID)
ChkSelfMenu = IIf(IsNull(RsMenuDetail!SelfMenu), 0, RsMenuDetail!SelfMenu)

TxtSequenceNo = IIf(IsNull(RsMenuDetail!SequenceNo), "", RsMenuDetail!SequenceNo)
TxtFormDesc = IIf(IsNull(RsMenuDetail!FormDesc), "", RsMenuDetail!FormDesc)
TxtDetailDesc = IIf(IsNull(RsMenuDetail!detaildesc), "", RsMenuDetail!detaildesc)
If mMasterMenuID <> 0 Then
   tmp = "Select MenuName from Mst_MenuDetail Where MenuID = " & mMasterMenuID & ""
   Call TmpTable
   
   TxtMasterMenu = TmpRs!MenuName 'GetMenuName(mMasterMenu)
Else
   TxtMasterMenu = ""
End If

'Frame1.Enabled = False
TxtFormDesc.Locked = True
TxtDetailDesc.Locked = True
TxtFlag.Locked = True
TxtFormFormName.Locked = True
TxtFormName.Locked = True
TxtHMenuFlag.Locked = True
TxtMasterMenu.Locked = True
TxtMenuName.Locked = True
TxtSequenceNo.Locked = True
ChkSelfMenu.Enabled = False
Call GButtonLock_Pwd(Me, True)

Call Grid_HeadAccess
Call ShowAccessDetails



End Sub

Private Sub LastCmd_Click()
RsMenuDetail.MoveLast
Call InData
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMenuDetail.MoveFirst

If RsMenuDetail.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMenuDetail.MoveFirst
RsMenuDetail.Find "MenuID = " & Val(mCaId) & ""
If RsMenuDetail.EOF Then
   RsMenuDetail.MoveLast
End If
Frame0.Visible = False
Call InData

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMenuDetail.MoveNext
If RsMenuDetail.EOF Then
   RsMenuDetail.MoveLast
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
   NextCmd.Enabled = False
   LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   Label19.Caption = Option1.Caption
ElseIf Option2.Value = True Then
   Label19.Caption = Option2.Caption
Else
   Label19.Caption = Option3.Caption
End If

End Sub


Private Sub Option2_Click()
If Option1.Value = True Then
   Label19.Caption = Option1.Caption
ElseIf Option2.Value = True Then
   Label19.Caption = Option2.Caption
Else
   Label19.Caption = Option3.Caption
End If
End Sub

Private Sub Option3_Click()
If Option1.Value = True Then
   Label19.Caption = Option1.Caption
ElseIf Option2.Value = True Then
   Label19.Caption = Option2.Caption
Else
   Label19.Caption = Option3.Caption
End If
End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMenuDetail.MovePrevious
If RsMenuDetail.BOF Then
   RsMenuDetail.MoveFirst
   LF_Flag = "Y"
End If
Call InData
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

If TxtFormFormName = "" Then
   MsgBox "Blank Bank Name not allowed"
   TxtFormFormName.SetFocus
   Exit Sub
End If

If TxtMenuName = "" Then
   MsgBox "Blank Menu Name not allowed"
   TxtMenuName.SetFocus
   Exit Sub
End If

If TxtFormName = "" Then
   MsgBox "Blank Form Name not allowed"
   TxtFormName.SetFocus
   Exit Sub
End If

'If TxtSequenceNo = "" Then
'   MsgBox "Blank Sequence No not allowed"
'   TxtSequenceNo.SetFocus
'   Exit Sub
'End If

If Edit_Flg = "A" Then
   If RsMenuDetail.RecordCount = 0 Then
      tmp = 1
   Else
      RsMenuDetail.MoveFirst
      RsMenuDetail.Find "FormFileName = '" & TxtFormFormName & "'"
      If Not RsMenuDetail.EOF Then
         MsgBox "Duplicate Bank Name not allowed"
         Exit Sub
      End If
      RsMenuDetail.MoveLast
     
      tmp = "select max(MenuID) from Mst_MenuDetail"
      Call TmpTable
      tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
   End If
   RsMenuDetail.AddNew
   RsMenuDetail!MenuID = tmp
   RsMenuDetail!G_UID = GetGUID
   TxtMenuID = tmp
   
Else
   RsMenuDetail.MoveFirst
   RsMenuDetail.Find "FormFileName = '" & TxtFormFormName & "'"
   If Not RsMenuDetail.EOF Then
      If RsMenuDetail!MenuID <> TxtMenuID Then
         MsgBox "Duplicate Menu Name not allowed"
         Exit Sub
      End If
   End If
   RsMenuDetail.MoveFirst
   RsMenuDetail.Find "MenuID = " & TxtMenuID & ""
End If

RsMenuDetail!FormFileName = TxtFormFormName
RsMenuDetail!MenuName = TxtMenuName

RsMenuDetail!FormName = TxtFormName
RsMenuDetail!Flag = TxtFlag
RsMenuDetail!HMenuFlag = TxtHMenuFlag

RsMenuDetail!SelfMenu = ChkSelfMenu.Value
If CmbMasterMenu.Text <> "" Then
   RsMenuDetail!MasterMenuID = mMasterMenuID
Else
   RsMenuDetail!MasterMenuID = Null
End If
RsMenuDetail!SequenceNo = Val(TxtSequenceNo)
RsMenuDetail!FormDesc = TxtFormDesc.Text
RsMenuDetail!detaildesc = TxtDetailDesc.Text

RsMenuDetail.Update
RsMenuDetail.Requery
RsMenuDetail.MoveFirst
RsMenuDetail.Find "MenuID = " & TxtMenuID & ""
If RsMenuDetail.EOF Then
   RsMenuDetail.MoveFirst
End If
Call GButtonLock_Pwd(Me, True)
TxtMenuID.Locked = True
Call InData
End Sub

Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock_Pwd(Me, True)
   If RsMenuDetail.EOF Then
      RsMenuDetail.MoveFirst
   End If
   Call InData

   Exit Sub
End If

Frame0.Height = Me.Height - 800
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400
RsMenuDetail.MoveFirst

If mSearch = False Then
   Option1.Value = True
   Call Grid_Head
   TxtSearchForm = ""
End If

mSearch = True
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.WordWrap = True
'MSHFlexGrid1.TextMatrix(0, 0) = "Menu Id"
'MSHFlexGrid1.TextMatrix(0, 1) = "Form File Name"

MSHFlexGrid1.TextMatrix(0, 0) = "Menu ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Form File Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Menu Name"
MSHFlexGrid1.TextMatrix(0, 3) = "FormName"
MSHFlexGrid1.TextMatrix(0, 4) = "Flag"
MSHFlexGrid1.TextMatrix(0, 5) = "Self Menu"
MSHFlexGrid1.TextMatrix(0, 6) = "Master Menu"
MSHFlexGrid1.TextMatrix(0, 7) = "Form Type"
MSHFlexGrid1.TextMatrix(0, 8) = "Sequence No"
MSHFlexGrid1.TextMatrix(0, 9) = "Form Description"
MSHFlexGrid1.TextMatrix(0, 10) = "Detail Description"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 2400
MSHFlexGrid1.ColWidth(2) = 2400
MSHFlexGrid1.ColWidth(3) = 2400
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 3000
MSHFlexGrid1.ColWidth(10) = 3000

End Sub


Private Sub TxtMenuID_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtMenuID)
If ans = False Then
   MsgBox "Special character not allowed!!! "
   TxtMenuID.SetFocus
End If
End Sub

Private Sub Grid_HeadAccess()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 5
MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.RowHeight(0) = 800

MSHFlexGrid2.TextMatrix(0, 0) = "User Role"
MSHFlexGrid2.TextMatrix(0, 1) = "Form Access"
MSHFlexGrid2.TextMatrix(0, 2) = "Add"
MSHFlexGrid2.TextMatrix(0, 3) = "Edit"
MSHFlexGrid2.TextMatrix(0, 4) = "Delete"

MSHFlexGrid2.ColWidth(0) = 1400
MSHFlexGrid2.ColWidth(1) = 1000
MSHFlexGrid2.ColWidth(2) = 600
MSHFlexGrid2.ColWidth(3) = 600
MSHFlexGrid2.ColWidth(4) = 600

End Sub

