VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ExpenseGroup 
   Caption         =   "Expense Group"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   6720
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2415
      Left            =   6840
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   150
         TabIndex        =   19
         Top             =   195
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1740
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3069
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   6480
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5100
         TabIndex        =   4
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3885
         TabIndex        =   17
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2550
         TabIndex        =   16
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1335
         TabIndex        =   15
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5100
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3885
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2550
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1335
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6495
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3735
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox TxtExpenseGroup 
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
      Begin VB.TextBox TxtExpenseGroupID 
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtNAVCode 
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
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label8 
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
         TabIndex        =   24
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         TabIndex        =   23
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Expense Group"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Expense Group ID"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "NAV Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   1237
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_ExpenseGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtExpenseGroup.Locked = False
TxtNAVCode.Locked = False
TxtExpenseGroup = ""
TxtExpenseGroupID = ""
TxtNAVCode = ""
TxtCreated = ""
TxtUpdated = ""

If RsMst_ExpenseGroup.RecordCount > 0 Then
   TxtExpenseGroup.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ExpenseGroup.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ExpenseGroup.xls")
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
Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   'tmp = "Select * from Mst_Employee Where (ExpenseGroupID = " & TxtExpenseGroupID & ")"
   'Call Tmp1Table
   'If TmpRs1.RecordCount > 0 Then
   '   MsgBox "Child Record Present!!!"
   '   Exit Sub
   'End If
   RsMst_ExpenseGroup.Delete
   RsMst_ExpenseGroup.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_ExpenseGroup.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ExpenseGroup.MoveNext
   If RsMst_ExpenseGroup.EOF() Then
      RsMst_ExpenseGroup.MoveLast
   End If
   Call Indata
   'TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_ExpenseGroup.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtExpenseGroup.Locked = False
TxtNAVCode.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_ExpenseGroup.MoveFirst
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
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_ExpenseGroup(" Order by ExpenseGroupID ")

If RsMst_ExpenseGroup.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtExpenseGroup.Locked = True
TxtNAVCode.Locked = True
TxtExpenseGroupID = RsMst_ExpenseGroup!ExpenseGroupID
TxtExpenseGroup = RsMst_ExpenseGroup!ExpenseGroup
TxtNAVCode = IIf(IsNull(RsMst_ExpenseGroup!NAVCode), "", RsMst_ExpenseGroup!NAVCode)
TxtCreated = IIf(IsNull(RsMst_ExpenseGroup!CreatedBy), "", RsMst_ExpenseGroup!CreatedBy) & " :- " & IIf(IsNull(RsMst_ExpenseGroup!CreatedDate), "", RsMst_ExpenseGroup!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ExpenseGroup!UpdatedBy), "", RsMst_ExpenseGroup!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ExpenseGroup!UpdatedDate), "", RsMst_ExpenseGroup!UpdatedDate)
Call GButtonLock(Me, True)
End Sub



Private Sub LastCmd_Click()
RsMst_ExpenseGroup.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_ExpenseGroup.MoveFirst

If RsMst_ExpenseGroup.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_ExpenseGroup.MoveFirst
RsMst_ExpenseGroup.Find "ExpenseGroupID = " & Val(mCaId) & ""
If RsMst_ExpenseGroup.EOF Then
   RsMst_ExpenseGroup.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ExpenseGroup.MoveNext
If RsMst_ExpenseGroup.EOF Then
   RsMst_ExpenseGroup.MoveLast
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
RsMst_ExpenseGroup.MovePrevious
If RsMst_ExpenseGroup.BOF Then
   RsMst_ExpenseGroup.MoveFirst
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

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtExpenseGroup = "" Then
   MsgBox "Blank Expense Group  not allowed !!!"
   TxtExpenseGroup.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_ExpenseGroup.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_ExpenseGroup.MoveFirst
        RsMst_ExpenseGroup.Find "ExpenseGroup = '" & TxtExpenseGroup & "'"
        If Not RsMst_ExpenseGroup.EOF Then
           MsgBox "Duplicate ExpenseGroup not Allowed"
           Exit Sub
        End If
        RsMst_ExpenseGroup.MoveLast
        
        tmp = "select max(ExpenseGroupID) from Mst_ExpenseGroup"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_ExpenseGroup.AddNew
    RsMst_ExpenseGroup!ExpenseGroupID = tmp
    TxtExpenseGroupID = tmp
Else
    RsMst_ExpenseGroup.MoveFirst
    RsMst_ExpenseGroup.Find "ExpenseGroup = '" & TxtExpenseGroup & "'"
    If Not RsMst_ExpenseGroup.EOF Then
       If RsMst_ExpenseGroup!ExpenseGroupID <> TxtExpenseGroupID Then
          MsgBox "Duplicate ExpenseGroup not Allowed"
          Exit Sub
       End If
    End If
   RsMst_ExpenseGroup.MoveFirst
   RsMst_ExpenseGroup.Find "ExpenseGroupID = " & TxtExpenseGroupID & ""
End If
RsMst_ExpenseGroup!ExpenseGroup = TxtExpenseGroup
RsMst_ExpenseGroup!NAVCode = TxtNAVCode

If IsNull(RsMst_ExpenseGroup!CreatedBy) = True Or RsMst_ExpenseGroup!CreatedBy = "" Then
   RsMst_ExpenseGroup!CreatedBy = Gen_UserLoginID
   RsMst_ExpenseGroup!CreatedDate = Now
Else
   RsMst_ExpenseGroup!UpdatedBy = Gen_UserLoginID
   RsMst_ExpenseGroup!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_ExpenseGroup.Update
RsMst_ExpenseGroup.Requery
RsMst_ExpenseGroup.MoveFirst
RsMst_ExpenseGroup.Find "ExpenseGroupID = " & TxtExpenseGroupID & ""

Call GButtonLock(Me, True)
TxtExpenseGroup.Locked = True
TxtNAVCode.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_ExpenseGroup.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_ExpenseGroup.EOF Then
      RsMst_ExpenseGroup.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500
Frame0.Left = 120
Frame0.Top = 120

Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 500
MSHFlexGrid1.Height = Frame0.Height - 700

RsMst_ExpenseGroup.MoveFirst
MSHFlexGrid1.Rows = RsMst_ExpenseGroup.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "Expense Group Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Expense Group"
MSHFlexGrid1.TextMatrix(0, 2) = "NAV Code"

MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 5500
MSHFlexGrid1.ColWidth(2) = 3500
For I = 1 To RsMst_ExpenseGroup.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_ExpenseGroup.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_ExpenseGroup.Fields(1)
    MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(RsMst_ExpenseGroup.Fields(2)), "", RsMst_ExpenseGroup.Fields(2))
    RsMst_ExpenseGroup.MoveNext
Next

End Sub

Private Sub TxtExpenseGroup_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtExpenseGroup")
End If
End Sub

Private Sub TxtExpenseGroup_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtExpenseGroup)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtExpenseGroup.SetFocus
End If
End Sub


Private Sub TxtExpenseGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtExpenseGroupID")
End If
End Sub
Private Sub TxtNAVCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtNAVCode")
End If
End Sub

