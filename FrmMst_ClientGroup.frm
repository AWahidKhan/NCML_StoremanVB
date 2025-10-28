VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ClientGroup 
   Caption         =   "Client Group Master"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   10110
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   975
      Left            =   7680
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
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
         Height          =   390
         Left            =   7005
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   225
         Width           =   1005
      End
      Begin VB.TextBox TxtSClientGroup 
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
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   2895
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   225
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1260
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2223
         _Version        =   393216
         Rows            =   3
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
      Begin VB.Label Label20 
         Caption         =   "Total Records Found :"
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
         Left            =   8205
         TabIndex        =   31
         Top             =   270
         Width           =   2415
      End
      Begin VB.Label LblTotalRecord 
         Caption         =   "-"
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
         Left            =   10605
         TabIndex        =   30
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Client Group"
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
         Left            =   120
         TabIndex        =   29
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   17
      Top             =   8400
      Width           =   8040
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         TabIndex        =   23
         Top             =   630
         Width           =   1575
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1669
         TabIndex        =   3
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1669
         TabIndex        =   22
         Top             =   630
         Width           =   1575
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3218
         TabIndex        =   21
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3218
         TabIndex        =   20
         Top             =   630
         Width           =   1575
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4767
         TabIndex        =   19
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4767
         TabIndex        =   18
         Top             =   630
         Width           =   1575
      End
      Begin VB.CommandButton SearchCmd 
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
         Height          =   400
         Left            =   6315
         TabIndex        =   5
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   6315
         TabIndex        =   4
         Top             =   630
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   8055
      Begin VB.CheckBox ChkAllClient 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5415
         TabIndex        =   35
         Top             =   510
         Width           =   255
      End
      Begin VB.CommandButton CmdAddToGrid 
         Caption         =   "Add To Grid"
         Height          =   375
         Left            =   6210
         TabIndex        =   34
         Top             =   465
         Width           =   1335
      End
      Begin VB.TextBox TxtClientName 
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
         Left            =   1425
         TabIndex        =   32
         Top             =   435
         Width           =   3855
      End
      Begin VB.CheckBox ChkCheckAll 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   15
         Top             =   915
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4695
         Left            =   105
         TabIndex        =   2
         Top             =   1230
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   8281
         _Version        =   393216
         ScrollBars      =   2
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "All"
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
         Left            =   5730
         TabIndex        =   36
         Top             =   510
         Width           =   225
      End
      Begin VB.Label Label4 
         Caption         =   "Client Name"
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
         TabIndex        =   33
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select / Deselect"
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
         Left            =   450
         TabIndex        =   16
         Top             =   915
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8055
      Begin VB.TextBox TxtCreated 
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox TxtUpdated 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1665
         Width           =   3735
      End
      Begin VB.TextBox TxtClientGroup 
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
         Width           =   3735
      End
      Begin VB.TextBox TxtClientGroupID 
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
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
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
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
         Top             =   1665
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Client Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Client Group ID"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmMst_ClientGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
TxtClientGroup = ""
TxtClientGroupID = ""
TxtClientGroup.Locked = False

TxtClientName.Text = ""
ChkAllClient.Value = 0

TxtClientName.Enabled = True
ChkAllClient.Enabled = True
CmdAddToGrid.Enabled = True

TxtCreated = ""
TxtUpdated = ""
Call Grid_Head1
If RsMst_ClientGroup.RecordCount > 0 Then
   TxtClientGroup.SetFocus
Else
   SearchCmd.Enabled = False
End If

End Sub

Private Sub ChkCheckAll_Click()
If SaveCmd.Enabled = False Then
   ChkCheckAll.Value = 0
   Exit Sub
End If
If MSHFlexGrid2.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.Row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.Row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
   Next
End If

End Sub


Private Sub CmdAddToGrid_Click()
Call GetGridData
TxtClientName.Text = ""
ChkAllClient.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ClientGroup.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ClientGroup.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
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

Call Grid_Head

tmp = " Select * from Mst_ClientGroup"


If TxtSClientGroup.Text <> "" Then
   wc = " Where ClientGroupName Like '%" & TxtSClientGroup.Text & "%' "
End If

tmp = tmp & wc & " Order By ClientGroupName"

Call TmpTable

For I = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!ClientGroupID
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!ClientGroupName
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    TmpRs.MoveNext
Next
LblTotalRecord.Caption = TmpRs.RecordCount

End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.AllowUserResizing = flexResizeBoth
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Client GroupID"
MSHFlexGrid1.TextMatrix(0, 1) = "Client Group"
MSHFlexGrid1.TextMatrix(0, 2) = "Flag"
MSHFlexGrid1.RowHeight(0) = 500

MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 6000
MSHFlexGrid1.ColWidth(2) = 0

End Sub



Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError


tmp = "Select * from Mst_Client where ClientGroupID = " & Val(TxtClientGroupID) & ""

Call TmpTable
   
If TmpRs.RecordCount > 0 Then
   MsgBox "Group mapped to Client. Can't Delete!"
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then

   RsMst_ClientGroup.Delete
   RsMst_ClientGroup.Update
   
   If RsMst_ClientGroup.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ClientGroup.MoveNext
   If RsMst_ClientGroup.EOF() Then
      RsMst_ClientGroup.MoveLast
   End If
   Call Indata
   
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_ClientGroup.CancelUpdate

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True
TxtClientGroup.SetFocus

TxtClientName.Enabled = True
ChkAllClient.Enabled = True
CmdAddToGrid.Enabled = True

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_ClientGroup.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
TxtSClientGroup.Text = ""
LblTotalRecord.Caption = ""

Call TableMst_ClientGroup
If RsMst_ClientGroup.RecordCount = 0 Then
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



Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
If MSHFlexGrid1.TextMatrix(x, 0) = "" Then Exit Sub
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_ClientGroup.MoveFirst

If RsMst_ClientGroup.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_ClientGroup.MoveFirst
RsMst_ClientGroup.Find "ClientGroupID = " & Val(mCaId) & ""
If RsMst_ClientGroup.EOF Then
   RsMst_ClientGroup.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid2
    If .TextMatrix(.Row, 2) = "" Then Exit Sub
    If .Col = 0 Then
        If .TextMatrix(.Row, 0) = strUnChecked Then
           .TextMatrix(.Row, 0) = strChecked
        Else
           .TextMatrix(.Row, 0) = strUnChecked
          
        End If
    End If
End With
End Sub
Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtClientGroup = "" Then
   MsgBox "Blank ClientGroup  not Allowed"
   TxtClientGroup.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsMst_ClientGroup.RecordCount > 0 Then
        RsMst_ClientGroup.MoveFirst
        RsMst_ClientGroup.Find "ClientGroupName = '" & TxtClientGroup & "'"
        If Not RsMst_ClientGroup.EOF Then
           MsgBox "Duplicate ClientGroup not Allowed"
           TxtClientGroup.SetFocus
           Exit Sub
        End If
    End If
    RsMst_ClientGroup.AddNew
    RsMst_ClientGroup!ClientGroupID = GetMaxClientGroupID
    RsMst_ClientGroup!G_UID = GetGUID
    TxtClientGroupID = RsMst_ClientGroup!ClientGroupID
Else
    RsMst_ClientGroup.MoveFirst
    RsMst_ClientGroup.Find "ClientGroupName = '" & TxtClientGroup & "'"
    If Not RsMst_ClientGroup.EOF Then
       If RsMst_ClientGroup!ClientGroupID <> TxtClientGroupID Then
          MsgBox "Duplicate ClientGroup not Allowed"
          TxtClientGroup.SetFocus
          Exit Sub
       End If
    End If
   RsMst_ClientGroup.MoveFirst
   RsMst_ClientGroup.Find "ClientGroupID = " & TxtClientGroupID & ""
End If
RsMst_ClientGroup!ClientGroupName = TxtClientGroup

If IsNull(RsMst_ClientGroup!CreatedBy) = True Or RsMst_ClientGroup!CreatedBy = "" Then
   RsMst_ClientGroup!CreatedBy = Gen_UserLoginID
   RsMst_ClientGroup!CreatedDate = Now
Else
   RsMst_ClientGroup!UpdatedBy = Gen_UserLoginID
   RsMst_ClientGroup!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If

RsMst_ClientGroup.Update
RsMst_ClientGroup.Requery

Call UpdateClientMaster
RsMst_ClientGroup.MoveFirst
RsMst_ClientGroup.Find "ClientGroupID = " & TxtClientGroupID & ""

Call GButtonLock(Me, True)
TxtClientGroup.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_ClientGroup.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_ClientGroup.EOF Then
      RsMst_ClientGroup.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_ClientGroup.MoveFirst
TxtSClientGroup.SetFocus

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'TxtSClientGroup.Text = ""
'LblTotalRecord.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtClientGroup_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtClientGroup")
End If
End Sub

Private Sub TxtClientGroup_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtClientGroup)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtClientGroup.SetFocus
End If
TxtClientGroup.Text = Trim(TxtClientGroup.Text)
End Sub

Private Sub TxtClientGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtClientGroupID")
End If
End Sub
Public Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .FixedCols = 0
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "ClientID"
    .TextMatrix(0, 2) = "Client ID"
    .TextMatrix(0, 3) = "Client Name"
    .TextMatrix(0, 4) = "Exchange Type"
    .ColAlignment(2) = vbLeftJustify
    .ColAlignment(3) = vbLeftJustify
    .ColWidth(0) = 800
    .ColWidth(1) = 0
    .ColWidth(2) = 1200
    .ColWidth(3) = 4000
    .ColWidth(4) = 1500
End With
End Sub

Private Sub UpdateClientMaster()

tmp = "Update Mst_Client Set ClientGroupID =Null Where ClientGroupID =" & TxtClientGroupID

Call TmpTable

Call TableMst_Client(" Where ClientGroupID Is Null")

For I = 1 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
       If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
          If RsMst_Client.RecordCount > 0 Then
             RsMst_Client.MoveFirst
          End If
          RsMst_Client.Find "ClientIDRow= " & Val(MSHFlexGrid2.TextMatrix(I, 1))
          If Not RsMst_Client.EOF Then
             RsMst_Client!ClientGroupID = Val(TxtClientGroupID)
             RsMst_Client.Update
          End If
       End If
    End If
Next
End Sub
Private Sub InDataGrid()

Call Grid_Head1

tmp = " Select ClientIdRow,ClientID,ClientName,ExchangeType From Mst_Client inner join Mst_ExchangeType on Mst_Client.ExchangeTypeID=Mst_ExchangeType.ExchangeTypeID Where ClientGroupID=" & Val(TxtClientGroupID) & ""

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!ClientIDRow
       MSHFlexGrid2.TextMatrix(I, 2) = TmpRs1!ClientID
       MSHFlexGrid2.TextMatrix(I, 3) = TmpRs1!ClientName
       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs1!ExchangeType
       MSHFlexGrid2.Row = I
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellFontName = "Wingdings"
       MSHFlexGrid2.CellFontSize = 12
       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid2.Text = strChecked
       TmpRs1.MoveNext
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   Next
End If
End Sub
Private Function GetMaxClientGroupID() As Double
Dim Retval As Double
tmp = "Select max(ClientGroupID) from Mst_ClientGroup"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClientGroupID = Retval
End Function
Private Sub GetGridData()

If ChkAllClient.Value = 1 Then
   tmp = " Select ClientIdRow,ClientID,ClientName,ExchangeType From Mst_Client inner join Mst_ExchangeType on Mst_Client.ExchangeTypeID=Mst_ExchangeType.ExchangeTypeID Where ClientGroupID Is Null  Order By ClientName"
Else
  tmp = " Select ClientIdRow,ClientID,ClientName ,ExchangeType From Mst_Client inner join Mst_ExchangeType on Mst_Client.ExchangeTypeID=Mst_ExchangeType.ExchangeTypeID Where ClientGroupID Is Null And ClientName like '%" & TxtClientName.Text & "%' Order By ClientName"
End If

Call TmpTable


If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       If CheckClientID(TmpRs!ClientIDRow) = False Then
          mRow = MSHFlexGrid2.Rows - 1
          MSHFlexGrid2.TextMatrix(mRow, 1) = TmpRs!ClientIDRow
          MSHFlexGrid2.TextMatrix(mRow, 2) = TmpRs!ClientID
          MSHFlexGrid2.TextMatrix(mRow, 3) = TmpRs!ClientName
          MSHFlexGrid2.TextMatrix(mRow, 4) = TmpRs!ExchangeType
          MSHFlexGrid2.Row = mRow
          MSHFlexGrid2.Col = 0
          MSHFlexGrid2.CellFontName = "Wingdings"
          MSHFlexGrid2.CellFontSize = 12
          MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid2.Text = strUnChecked
          TmpRs.MoveNext
          MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       End If
   Next
Else
  MsgBox "No Record Found For Grid!!!"
  Exit Sub
End If
End Sub
Private Function CheckClientID(ClientIDRow_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False

For r = 1 To MSHFlexGrid2.Rows - 1
    If Val(MSHFlexGrid2.TextMatrix(r, 1)) = ClientIDRow_ Then
       Retval = True
       Exit For
    End If
Next
CheckClientID = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ClientGroup.MoveNext
If RsMst_ClientGroup.EOF Then
   RsMst_ClientGroup.MoveLast
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
RsMst_ClientGroup.MovePrevious
If RsMst_ClientGroup.BOF Then
   RsMst_ClientGroup.MoveFirst
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
Public Sub Indata()
Frame1.Enabled = False

ChkCheckAll.Value = 0

TxtClientName.Text = ""
ChkAllClient.Value = 0
TxtClientName.Enabled = False
ChkAllClient.Enabled = False
CmdAddToGrid.Enabled = False

TxtClientGroupID = RsMst_ClientGroup!ClientGroupID
TxtClientGroup = RsMst_ClientGroup!ClientGroupName

TxtCreated = IIf(IsNull(RsMst_ClientGroup!CreatedBy), "", RsMst_ClientGroup!CreatedBy) & " :- " & IIf(IsNull(RsMst_ClientGroup!CreatedDate), "", RsMst_ClientGroup!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ClientGroup!UpdatedBy), "", RsMst_ClientGroup!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ClientGroup!UpdatedDate), "", RsMst_ClientGroup!UpdatedDate)

'Call GetGridData

Call InDataGrid
Call GButtonLock(Me, True)
End Sub


Private Sub LastCmd_Click()
RsMst_ClientGroup.MoveLast
Call Indata
End Sub
