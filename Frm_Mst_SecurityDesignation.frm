VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_SecurityDesignation 
   Caption         =   "Security Designation"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   3930
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   150
         TabIndex        =   12
         Top             =   195
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1740
         Left            =   120
         TabIndex        =   13
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
      TabIndex        =   7
      Top             =   2250
      Width           =   6825
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   615
         Width           =   1305
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1436
         TabIndex        =   21
         Top             =   615
         Width           =   1305
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2752
         TabIndex        =   20
         Top             =   615
         Width           =   1305
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   4068
         TabIndex        =   19
         Top             =   615
         Width           =   1305
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5385
         TabIndex        =   18
         Top             =   615
         Width           =   1305
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   375
         Left            =   5385
         TabIndex        =   10
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4068
         TabIndex        =   9
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2752
         TabIndex        =   8
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1436
         TabIndex        =   2
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2250
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6825
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
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1695
         Width           =   4185
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
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1210
         Width           =   4185
      End
      Begin VB.TextBox TxtSecurityDesignation 
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
         Left            =   2505
         TabIndex        =   1
         Top             =   725
         Width           =   4185
      End
      Begin VB.TextBox TxtSecurityDesignationID 
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
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
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
         TabIndex        =   17
         Top             =   1695
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
         TabIndex        =   16
         Top             =   1215
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Security Designation"
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
         TabIndex        =   6
         Top             =   780
         Width           =   2250
      End
      Begin VB.Label Label1 
         Caption         =   "Security Designation ID"
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
         TabIndex        =   5
         Top             =   300
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmMst_SecurityDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtSecurityDesignation.Locked = False
TxtSecurityDesignation = ""
TxtSecurityDesignationID = ""
TxtCreated = ""
TxtUpdated = ""

If RsMst_SecurityDesignation.RecordCount > 0 Then
   TxtSecurityDesignation.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_SecurityDesignation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_SecurityDesignation.xls")
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
   'tmp = "Select * from Mst_Employee Where (SecurityDesignationID = " & TxtSecurityDesignationID & ")"
   'Call Tmp1Table
   'If TmpRs1.RecordCount > 0 Then
   '   MsgBox "Child Record Present!!!"
   '   Exit Sub
   'End If
   RsMst_SecurityDesignation.Delete
   RsMst_SecurityDesignation.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_SecurityDesignation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SecurityDesignation.MoveNext
   If RsMst_SecurityDesignation.EOF() Then
      RsMst_SecurityDesignation.MoveLast
   End If
   Call Indata
   'TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_SecurityDesignation.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtSecurityDesignation.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_SecurityDesignation.MoveFirst
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

Call TableMst_SecurityDesignation(" Order by SecurityDesignationID ")

If RsMst_SecurityDesignation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtSecurityDesignation.Locked = True
TxtSecurityDesignationID = RsMst_SecurityDesignation!SecurityDesignationID
TxtSecurityDesignation = RsMst_SecurityDesignation!SecurityDesignation
TxtCreated = IIf(IsNull(RsMst_SecurityDesignation!CreatedBy), "", RsMst_SecurityDesignation!CreatedBy) & " :- " & IIf(IsNull(RsMst_SecurityDesignation!CreatedDate), "", RsMst_SecurityDesignation!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_SecurityDesignation!UpdatedBy), "", RsMst_SecurityDesignation!UpdatedBy) & " :- " & IIf(IsNull(RsMst_SecurityDesignation!UpdatedDate), "", RsMst_SecurityDesignation!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_SecurityDesignation.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mSecDesigID As Variant

x = MSHFlexGrid1.RowSel
mSecDesigID = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_SecurityDesignation.MoveFirst

If RsMst_SecurityDesignation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_SecurityDesignation.MoveFirst
RsMst_SecurityDesignation.Find "SecurityDesignationID = " & Val(mSecDesigID) & ""
If RsMst_SecurityDesignation.EOF Then
   RsMst_SecurityDesignation.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SecurityDesignation.MoveNext
If RsMst_SecurityDesignation.EOF Then
   RsMst_SecurityDesignation.MoveLast
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
RsMst_SecurityDesignation.MovePrevious
If RsMst_SecurityDesignation.BOF Then
   RsMst_SecurityDesignation.MoveFirst
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

If TxtSecurityDesignation = "" Then
   MsgBox "Blank Security Designation not allowed !!!"
   TxtSecurityDesignation.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_SecurityDesignation.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_SecurityDesignation.MoveFirst
        RsMst_SecurityDesignation.Find "SecurityDesignation = '" & TxtSecurityDesignation & "'"
        If Not RsMst_SecurityDesignation.EOF Then
           MsgBox "Duplicate SecurityDesignation not Allowed"
           Exit Sub
        End If
        RsMst_SecurityDesignation.MoveLast
        
        tmp = "select max(SecurityDesignationID) from Mst_SecurityDesignation"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_SecurityDesignation.AddNew
    RsMst_SecurityDesignation!SecurityDesignationID = tmp
    TxtSecurityDesignationID = tmp
Else
    RsMst_SecurityDesignation.MoveFirst
    RsMst_SecurityDesignation.Find "SecurityDesignation = '" & TxtSecurityDesignation & "'"
    If Not RsMst_SecurityDesignation.EOF Then
       If RsMst_SecurityDesignation!SecurityDesignationID <> TxtSecurityDesignationID Then
          MsgBox "Duplicate Security Designation not Allowed"
          Exit Sub
       End If
    End If
   RsMst_SecurityDesignation.MoveFirst
   RsMst_SecurityDesignation.Find "SecurityDesignationID = " & TxtSecurityDesignationID & ""
End If
RsMst_SecurityDesignation!SecurityDesignation = TxtSecurityDesignation

If IsNull(RsMst_SecurityDesignation!CreatedBy) = True Or RsMst_SecurityDesignation!CreatedBy = "" Then
   RsMst_SecurityDesignation!CreatedBy = Gen_UserLoginID
   RsMst_SecurityDesignation!CreatedDate = Now
Else
   RsMst_SecurityDesignation!UpdatedBy = Gen_UserLoginID
   RsMst_SecurityDesignation!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_SecurityDesignation.Update
RsMst_SecurityDesignation.Requery
RsMst_SecurityDesignation.MoveFirst
RsMst_SecurityDesignation.Find "SecurityDesignationID = " & TxtSecurityDesignationID & ""

Call GButtonLock(Me, True)
TxtSecurityDesignation.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_SecurityDesignation.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_SecurityDesignation.EOF Then
      RsMst_SecurityDesignation.MoveFirst
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

RsMst_SecurityDesignation.MoveFirst
MSHFlexGrid1.Rows = RsMst_SecurityDesignation.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "Security Designation Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Security Designation"

MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 5500

For I = 1 To RsMst_SecurityDesignation.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_SecurityDesignation.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_SecurityDesignation.Fields(1)
    RsMst_SecurityDesignation.MoveNext
Next

End Sub

Private Sub TxtSecurityDesignation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSecurityDesignation")
End If
End Sub

Private Sub TxtSecurityDesignation_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtSecurityDesignation)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtSecurityDesignation.SetFocus
End If
End Sub


Private Sub TxtSecurityDesignationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSecurityDesignationID")
End If
End Sub

