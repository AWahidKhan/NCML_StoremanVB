VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_AuditType 
   Caption         =   "Audit Type"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   7080
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   19
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   20
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
   Begin VB.Frame Frame1 
      Height          =   2610
      Left            =   120
      TabIndex        =   12
      Top             =   75
      Width           =   6240
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4320
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1515
         Width           =   4320
      End
      Begin VB.TextBox TxtAuditTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox TxtAuditType 
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
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   645
         Width           =   4320
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1590
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1965
         Visible         =   0   'False
         Width           =   4200
      End
      Begin VB.Label Label5 
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
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Height          =   270
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label LblStateID 
         Caption         =   "Audit Type ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   278
         Width           =   1800
      End
      Begin VB.Label LblStateName 
         Caption         =   "Audit Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   668
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
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
         TabIndex        =   15
         Top             =   1980
         Visible         =   0   'False
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   5
      Top             =   2670
      Width           =   6240
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   4920
         TabIndex        =   3
         Top             =   645
         Width           =   1200
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   3720
         TabIndex        =   11
         Top             =   645
         Width           =   1200
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   405
         Left            =   2520
         TabIndex        =   10
         Top             =   645
         Width           =   1200
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1320
         TabIndex        =   8
         Top             =   645
         Width           =   1200
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   7
         Top             =   645
         Width           =   1200
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   0
         Top             =   240
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FrmMst_AuditType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtAuditType.Locked = False
TxtAuditType = ""
TxtAuditTypeID = ""
TxtFlag = ""
TxtCreated = ""
TxtUpdated = ""

If RsMst_AuditType.RecordCount > 0 Then
   TxtAuditType.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub


Private Sub DeleteCmd_Click()



Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   RsMst_AuditType.Delete
   RsMst_AuditType.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_AuditType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_AuditType.MoveNext
   If RsMst_AuditType.EOF() Then
      RsMst_AuditType.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_AuditType.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtAuditType.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_AuditType.MoveFirst
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


Call TableMst_AuditType("  Order by AuditTypeID ")
If RsMst_AuditType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtAuditType.Locked = True
'RsMst_AuditType
TxtAuditTypeID = RsMst_AuditType!AuditTypeID
TxtAuditType = RsMst_AuditType!AuditType
TxtCreated = IIf(IsNull(RsMst_AuditType!CreatedBy), "", RsMst_AuditType!CreatedBy) & " :- " & IIf(IsNull(RsMst_AuditType!CreatedDate), "", RsMst_AuditType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_AuditType!UpdatedBy), "", RsMst_AuditType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_AuditType!UpdatedDate), "", RsMst_AuditType!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_AuditType.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False

End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_AuditType.MoveFirst

If RsMst_AuditType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_AuditType.MoveFirst
RsMst_AuditType.Find "AuditTypeID = " & Val(mCaId) & ""
If RsMst_AuditType.EOF Then
   RsMst_AuditType.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_AuditType.MoveNext
If RsMst_AuditType.EOF Then
   RsMst_AuditType.MoveLast
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
RsMst_AuditType.MovePrevious
If RsMst_AuditType.BOF Then
   RsMst_AuditType.MoveFirst
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

If TxtAuditType = "" Then
   MsgBox "Blank Audit Type  not Allowed"
   TxtAuditType.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_AuditType.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_AuditType.MoveFirst
        RsMst_AuditType.Find "AuditType = '" & TxtAuditType & "'"
        If Not RsMst_AuditType.EOF Then
           MsgBox "Duplicate Audit Type not Allowed"
           Exit Sub
        End If
        RsMst_AuditType.MoveLast
        
        tmp = "select max(AuditTypeID) from Mst_AuditType"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_AuditType.AddNew
    RsMst_AuditType!AuditTypeID = tmp
    TxtAuditTypeID = tmp
    RsMst_AuditType!G_UID = GetGUID
Else
    RsMst_AuditType.MoveFirst
    RsMst_AuditType.Find "AuditType = '" & TxtAuditType & "'"
    If Not RsMst_AuditType.EOF Then
       If RsMst_AuditType!AuditTypeID <> TxtAuditTypeID Then
          MsgBox "Duplicate Audit Type not Allowed"
          Exit Sub
       End If
    End If
   RsMst_AuditType.MoveFirst
   RsMst_AuditType.Find "AuditTypeID = " & TxtAuditTypeID & ""
End If
RsMst_AuditType!AuditType = TxtAuditType


If IsNull(RsMst_AuditType!CreatedBy) = True Or RsMst_AuditType!CreatedBy = "" Then
   RsMst_AuditType!CreatedBy = Gen_UserLoginID
   RsMst_AuditType!CreatedDate = Now
Else
   RsMst_AuditType!UpdatedBy = Gen_UserLoginID
   RsMst_AuditType!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_AuditType.Update
RsMst_AuditType.Requery
RsMst_AuditType.MoveFirst
RsMst_AuditType.Find "AuditTypeID = " & TxtAuditTypeID & ""

Call GButtonLock(Me, True)
TxtAuditType.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_AuditType.EOF Then
      RsMst_AuditType.MoveFirst
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

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000

RsMst_AuditType.MoveFirst
MSHFlexGrid1.Rows = RsMst_AuditType.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "Audit Type Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Audit Type"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4500
For I = 1 To RsMst_AuditType.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_AuditType.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_AuditType.Fields(1)
    RsMst_AuditType.MoveNext
Next


End Sub

Private Sub TxtAuditType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAuditType")
End If
End Sub

Private Sub TxtAuditType_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAuditType)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAuditType.SetFocus
End If
End Sub


Private Sub TxtAuditTypeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAuditTypeID")
End If
End Sub


Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Mst_AuditType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\Mst_AuditType.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub



