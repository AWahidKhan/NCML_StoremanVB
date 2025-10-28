VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_State 
   Caption         =   "State Master"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2925
      Left            =   0
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   14880
      Begin VB.ComboBox CmbSStateIndiaName 
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
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   3675
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   8775
         TabIndex        =   30
         Top             =   492
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchState 
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
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   3645
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   7770
         TabIndex        =   28
         Top             =   495
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   31
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "State Name"
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
         Left            =   1342
         TabIndex        =   35
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label Label23 
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
         Left            =   9885
         TabIndex        =   34
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
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
         Left            =   12285
         TabIndex        =   33
         Top             =   540
         Width           =   75
      End
      Begin VB.Label Label2 
         Caption         =   "StateIndia  Name"
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
         Left            =   4897
         TabIndex        =   32
         Top             =   165
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   7200
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   13
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   8
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   12
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   7
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   405
         Left            =   2895
         TabIndex        =   11
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   6
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   10
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   9
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   5
         Top             =   240
         Width           =   1400
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   7200
      Begin VB.ComboBox CmbRegion 
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
         Left            =   2085
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   4635
      End
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2490
         Width           =   4650
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2955
         Width           =   4635
      End
      Begin VB.TextBox TxtStateID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox TxtStateName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2085
         MaxLength       =   100
         TabIndex        =   1
         Top             =   684
         Width           =   4635
      End
      Begin VB.ComboBox CmbStateIndiaName 
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
         Left            =   2085
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1110
         Width           =   4635
      End
      Begin VB.TextBox TxtStateIndiaName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   22
         Top             =   1110
         Width           =   4635
      End
      Begin VB.ComboBox CmbFlag 
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
         ItemData        =   "Frm_Mst_State.frx":0000
         Left            =   2085
         List            =   "Frm_Mst_State.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2055
         Width           =   4635
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   18
         Top             =   2055
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.TextBox TxtRegion 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   37
         Top             =   1560
         Width           =   4635
      End
      Begin VB.Label Label3 
         Caption         =   "Region"
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
         Left            =   120
         TabIndex        =   38
         Top             =   1605
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   2085
         Width           =   480
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
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3015
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "State India Name"
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
         Left            =   120
         TabIndex        =   23
         Top             =   1155
         Width           =   1800
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
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   2565
         Width           =   1695
      End
      Begin VB.Label LblStateID 
         Caption         =   "StateID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   308
         Width           =   1800
      End
      Begin VB.Label LblStateName 
         Caption         =   "State Name"
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
         Left            =   120
         TabIndex        =   17
         Top             =   750
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmMst_State"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mStateIndiaID, mRegionID As Variant

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub

If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStateIndiaName_GotFocus()
tmp = CmbStateIndiaName.Text
Call TableMst_StateIndia
CmbStateIndiaName.Clear
If RsMst_StateIndia.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_StateIndia.RecordCount
    CmbStateIndiaName.AddItem RsMst_StateIndia!StateIndiaName
    RsMst_StateIndia.MoveNext
Next
CmbStateIndiaName.Text = tmp
End Sub

Private Sub CmbStateIndiaName_LostFocus()
If CmbStateIndiaName.Text = "" Then
   Exit Sub
End If
RsMst_StateIndia.MoveFirst
RsMst_StateIndia.Find "StateIndiaName = '" & CmbStateIndiaName.Text & "'"
If RsMst_StateIndia.EOF Then
   MsgBox "Invalid selection "
   CmbStateIndiaName.SetFocus
   Exit Sub
End If
mStateIndiaID = RsMst_StateIndia!StateIndiaID
 
End Sub


Private Sub CmbRegion_GotFocus()
tmp = CmbRegion.Text
Call TableMst_Region
CmbRegion.Clear
If RsMst_Region.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Region.RecordCount
    CmbRegion.AddItem RsMst_Region!Region
    RsMst_Region.MoveNext
Next
CmbRegion.Text = tmp
End Sub

Private Sub CmbRegion_LostFocus()
If CmbRegion.Text = "" Then
   Exit Sub
End If
RsMst_Region.MoveFirst
RsMst_Region.Find "Region = '" & CmbRegion.Text & "'"
If RsMst_Region.EOF Then
   MsgBox "Invalid selection "
   CmbRegion.SetFocus
   Exit Sub
End If
mRegionID = RsMst_Region!RegionID
 
End Sub

Private Sub CmdGo_Click()

tmp = "Select MST.StateId,MST.StateName,MSTI.StateIndiaName,MR.Region,MST.Flag from Mst_State  MST"
tmp = tmp & " Left Join Mst_StateIndia MSTI On MST.StateIndiaID=MSTI.StateIndiaID "
tmp = tmp & " Left Join Mst_Region MR On MST.RegionID=MR.RegionID "

Dim wc As Variant
wc = ""
If TxtSearchState <> "" Then
   If wc = "" Then
      wc = " Where MST.StateName Like '%" & TxtSearchState & "%' "
   Else
      wc = wc & " and MST.StateName Like '%" & TxtSearchState & "%' "
   End If
End If

If CmbSStateIndiaName <> "" Then
   If wc = "" Then
      wc = " Where MSTI.StateIndiaName ='" & CmbSStateIndiaName & "' "
   Else
      wc = wc & " and MSTI.StateIndiaName ='" & CmbSStateIndiaName & "' "
   End If
End If

tmp = tmp & wc & " Order By MST.StateName "

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

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
TxtSearchState = ""
Label22.Caption = ""


Call TableMst_State
If RsMst_State.RecordCount = 0 Then
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
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_State.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_State.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_State.MoveNext
If RsMst_State.EOF Then
   RsMst_State.MoveLast
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
RsMst_State.MovePrevious
If RsMst_State.BOF Then
   RsMst_State.MoveFirst
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
Private Sub FirstCmd_Click()
RsMst_State.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_State.MoveLast
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
   If RsMst_State.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_State.EOF Then
      RsMst_State.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


'If SearchCmd.Caption = "Cancel" Then
'   If RsMst_Employee.RecordCount = 0 Then Exit Sub
'   Call GButtonLock(Me, True)
'   If RsMst_Employee.EOF Then
'      RsMst_Employee.MoveFirst
'   End If
'   Call Indata
'   Exit Sub
'End If


Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 600
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_State.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchState = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Rows = RsMst_State.RecordCount + 1
''MSHFlexGrid1.ColWidth(0) = 1000
'For i = 1 To RsMst_State.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_State.Fields(c)), "", RsMst_State.Fields(c))
'    Next
'    RsMst_State.MoveNext
'Next
End Sub
Public Sub Indata()

TxtStateID.Locked = True
TxtStateName.Locked = True
CmbFlag.Visible = False
TxtFlag.Visible = True
CmbStateIndiaName.Visible = False
TxtStateIndiaName.Visible = True
CmbRegion.Visible = False
TxtRegion.Visible = True


TxtStateID = IIf(IsNull(RsMst_State!StateID), "", RsMst_State!StateID)
TxtStateName = IIf(IsNull(RsMst_State!StateName), "", RsMst_State!StateName)


TxtCreated = IIf(IsNull(RsMst_State!CreatedBy), "", RsMst_State!CreatedBy) & " :- " & IIf(IsNull(RsMst_State!CreatedDate), "", RsMst_State!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_State!UpdatedBy), "", RsMst_State!UpdatedBy) & " :- " & IIf(IsNull(RsMst_State!UpdatedDate), "", RsMst_State!UpdatedDate)

'If RsMst_State!Flag = "A" Then
'   TxtFlag.Text = "Active"
'ElseIf RsMst_State!Flag = "D" Then
'   TxtFlag.Text = "Deactive"
'End If

If LCase(RsMst_State!Flag) = "a" Then
   TxtFlag.Text = "Active"
ElseIf LCase(RsMst_State!Flag) = "d" Then
   TxtFlag.Text = "De-Active"
Else
   TxtFlag.Text = ""
End If
'TxtFlag = IIf(IsNull(RsMst_State!Flag), "", RsMst_State!Flag)
mStateIndiaID = RsMst_State!StateIndiaID
mRegionID = RsMst_State!RegionID

If IsNull(RsMst_State!StateIndiaID) = False Then
   TxtStateIndiaName = GetStateIndiaName(RsMst_State!StateIndiaID)
Else
   TxtStateIndiaName = ""
End If

If IsNull(RsMst_State!RegionID) = False Then
   TxtRegion = GetRegion(RsMst_State!RegionID)
Else
   TxtRegion = ""
End If

If RsMst_State.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtFlag")
End If
End Sub

Private Sub TxtStateID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtStateID")
End If

End Sub

Private Sub TxtStateID_LostFocus()
If TxtStateID = "" Then Exit Sub
If IsNumeric(TxtStateID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtStateID.SetFocus
   Exit Sub
End If
TxtStateID = Val(TxtStateID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtStateID.Locked = True
TxtStateName.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True

TxtStateID = ""
TxtStateName = ""

CmbStateIndiaName.Visible = True
TxtStateIndiaName.Visible = False
CmbStateIndiaName.Text = ""
CmbFlag.Text = ""

CmbRegion.Visible = True
TxtRegion.Visible = False
CmbRegion.Text = ""

TxtCreated = ""
TxtUpdated = ""
If RsMst_State.RecordCount > 0 Then
   TxtStateName.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
CmbStateIndiaName.Visible = True
TxtStateIndiaName.Visible = False
CmbStateIndiaName.Text = TxtStateIndiaName.Text
TxtStateID.Locked = True
TxtStateName.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = TxtFlag.Text

CmbRegion.Visible = True
TxtRegion.Visible = False
CmbRegion.Text = TxtRegion.Text


End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_State.Delete
   RsMst_State.Update
   If RsMst_State.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_State.MoveNext
   If RsMst_State.EOF() Then
      RsMst_State.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_State.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtStateName.Text = "" Then
   MsgBox "Blank StateName Not Allowed !!! "
   TxtStateName.SetFocus
   Exit Sub
End If

If CmbStateIndiaName.Text = "" Then
   MsgBox "Blank State India Name Not Allowed !!! "
   CmbStateIndiaName.SetFocus
   Exit Sub
End If

If CmbRegion.Text = "" Then
   MsgBox "Blank Region Not Allowed !!! "
   CmbRegion.SetFocus
   Exit Sub
End If


If CmbFlag.Text = "" Then
   MsgBox "Blank Flag Not Allowed !!! "
   CmbFlag.SetFocus
   Exit Sub
End If




If Edit_Flg = "A" Then
   If RsMst_State.RecordCount > 0 Then
      RsMst_State.MoveFirst
      RsMst_State.Find "StateName= '" & TxtStateName.Text & "'"
      If Not RsMst_State.EOF Then
        MsgBox "Duplicate State Name Not Allowed !!! "
         TxtStateName.SetFocus
         Exit Sub
      End If
   End If
      RsMst_State.AddNew
      RsMst_State!StateID = GetStateID
      RsMst_State!G_UID = GetGUID
      TxtStateID.Text = RsMst_State!StateID
      RsMst_State!StateIndiaID = mStateIndiaID
Else
    RsMst_State.MoveFirst
    RsMst_State.Find "StateName= '" & TxtStateName.Text & "'"
    If Not RsMst_State.EOF Then
       If RsMst_State!StateID <> TxtStateID.Text Then
          MsgBox "Duplicate State Name Not Allowed !!! "
          TxtStateName.SetFocus
          Exit Sub
       End If
    End If
    RsMst_State.MoveFirst
    RsMst_State.Find " StateID= " & TxtStateID.Text
End If

RsMst_State!StateID = TxtStateID.Text
RsMst_State!StateName = TxtStateName.Text
RsMst_State!Flag = Left(CmbFlag.Text, 1)
RsMst_State!StateIndiaID = mStateIndiaID
RsMst_State!RegionID = mRegionID

If IsNull(RsMst_State!CreatedBy) = True Or RsMst_State!CreatedBy = "" Then
   RsMst_State!CreatedBy = Gen_UserLoginID
   RsMst_State!CreatedDate = Now
Else
   RsMst_State!UpdatedBy = Gen_UserLoginID
   RsMst_State!UpdatedDate = Now
End If

RsMst_State.Update
RsMst_State.Requery

RsMst_State.MoveFirst
RsMst_State.Find "StateID = " & TxtStateID

Call Indata

End Sub
Private Function GetStateID() As Double
Dim Retval As Double
tmp = "Select max(StateID) from Mst_State"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetStateID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5 '4
MSHFlexGrid1.TextMatrix(0, 0) = "StateID"
MSHFlexGrid1.TextMatrix(0, 1) = "State Name"
MSHFlexGrid1.TextMatrix(0, 2) = "State India Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Region"
MSHFlexGrid1.TextMatrix(0, 4) = "Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 3500
MSHFlexGrid1.ColWidth(3) = 3000
MSHFlexGrid1.ColWidth(4) = 1000

End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_State.MoveFirst
RsMst_State.Find "StateID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_State.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtStateName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtStateName")
End If
End Sub
Private Sub CmbSStateIndiaName_GotFocus()
tmp = CmbSStateIndiaName.Text
Call TableMst_StateIndia
CmbSStateIndiaName.Clear
If RsMst_StateIndia.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_StateIndia.RecordCount
    CmbSStateIndiaName.AddItem RsMst_StateIndia!StateIndiaName
    RsMst_StateIndia.MoveNext
Next
CmbSStateIndiaName.Text = tmp
End Sub

Private Sub CmbSStateIndiaName_LostFocus()
If CmbSStateIndiaName.Text = "" Then
   Exit Sub
End If
RsMst_StateIndia.MoveFirst
RsMst_StateIndia.Find "StateIndiaName = '" & CmbSStateIndiaName.Text & "'"
If RsMst_StateIndia.EOF Then
   MsgBox "Invalid selection "
   CmbSStateIndiaName.SetFocus
   Exit Sub
End If
mStateIndiaID = RsMst_StateIndia!StateIndiaID
 
End Sub
