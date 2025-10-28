VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Category 
   Caption         =   "Category"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2085
      Left            =   10320
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3360
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   13
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   12
         Top             =   480
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
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7245
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
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1005
         Width           =   5280
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
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1410
         Width           =   5280
      End
      Begin VB.TextBox TxtCategoryID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1320
      End
      Begin VB.TextBox TxtCategory 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   1
         Top             =   615
         Width           =   5280
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   1065
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1463
         Width           =   975
      End
      Begin VB.Label LblCategoryID 
         Caption         =   "Category ID"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblCategory 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   615
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   7245
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   4
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   3
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   11
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
         Height          =   400
         Left            =   2895
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   2
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   8
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   5
         Top             =   240
         Width           =   1400
      End
   End
End
Attribute VB_Name = "FrmMst_Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_Category
If RsMst_Category.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Category.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Category.xls")
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
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Category.MoveNext
If RsMst_Category.EOF Then
   RsMst_Category.MoveLast
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
RsMst_Category.MovePrevious
If RsMst_Category.BOF Then
   RsMst_Category.MoveFirst
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
RsMst_Category.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Category.MoveLast
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
   If RsMst_Category.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_Category.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_Category.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For I = 1 To RsMst_Category.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(RsMst_Category.Fields(c)), "", RsMst_Category.Fields(c))
    Next
    RsMst_Category.MoveNext
Next
End Sub
Public Sub Indata()

TxtCategoryID.Locked = True
TxtCategory.Locked = True

TxtCategoryID = IIf(IsNull(RsMst_Category!CategoryID), "", RsMst_Category!CategoryID)
TxtCategory = IIf(IsNull(RsMst_Category!Category), "", RsMst_Category!Category)

TxtCreated = IIf(IsNull(RsMst_Category!CreatedBy), "", RsMst_Category!CreatedBy) & " :- " & IIf(IsNull(RsMst_Category!CreatedDate), "", RsMst_Category!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Category!UpdatedBy), "", RsMst_Category!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Category!UpdatedDate), "", RsMst_Category!UpdatedDate)


If RsMst_Category.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtCategory_LostFocus()
TxtCategory = Trim(TxtCategory)
End Sub

Private Sub TxtCategoryID_LostFocus()
If TxtCategoryID = "" Then Exit Sub
If IsNumeric(TxtCategoryID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtCategoryID.SetFocus
   Exit Sub
End If
TxtCategoryID = Val(TxtCategoryID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtCategoryID.Locked = True
TxtCategory.Locked = False

TxtCategoryID = ""

TxtCategory = ""
TxtCreated = ""
TxtUpdated = ""


End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtCategoryID.Locked = True
TxtCategory.Locked = False

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_Category.Delete
   RsMst_Category.Update
   If RsMst_Category.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Category.MoveNext
   If RsMst_Category.EOF() Then
      RsMst_Category.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Category.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtCategory.Text = "" Then
   MsgBox "Blank Category Not Allowed !!! "
   TxtCategory.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Category.RecordCount > 0 Then
      RsMst_Category.MoveFirst
      RsMst_Category.Find "Category= '" & TxtCategory.Text & "'"
      If Not RsMst_Category.EOF Then
        MsgBox "Duplicate Category Name Not Allowed !!! "
         TxtCategory.SetFocus
         Exit Sub
      End If
   End If
      RsMst_Category.AddNew
      RsMst_Category!CategoryID = GetCategoryID
      RsMst_Category!G_UID = GetGUID
      TxtCategoryID.Text = RsMst_Category!CategoryID
Else
    RsMst_Category.MoveFirst
    RsMst_Category.Find "Category= '" & TxtCategory.Text & "'"
    If Not RsMst_Category.EOF Then
       If RsMst_Category!CategoryID <> TxtCategoryID.Text Then
          MsgBox "Duplicate Category Name Not Allowed !!! "
          TxtCategory.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Category.MoveFirst
    RsMst_Category.Find " CategoryID= " & TxtCategoryID.Text
End If

RsMst_Category!CategoryID = TxtCategoryID.Text
RsMst_Category!Category = TxtCategory.Text

If IsNull(RsMst_Category!CreatedBy) = True Or RsMst_Category!CreatedBy = "" Then
   RsMst_Category!CreatedBy = Gen_UserLoginID
   RsMst_Category!CreatedDate = Now
Else
   RsMst_Category!UpdatedBy = Gen_UserLoginID
   RsMst_Category!UpdatedDate = Now
End If


RsMst_Category.Update
Call Indata
End Sub
Private Function GetCategoryID() As Double
Dim Retval As Double
tmp = "Select max(CategoryID) from Mst_Category"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCategoryID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 2
MSHFlexGrid1.TextMatrix(0, 0) = "Category ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Category"

MSHFlexGrid1.ColWidth(1) = 6000

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_Category.MoveFirst
RsMst_Category.Find "CategoryID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_Category.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
