VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_AuditSegment 
   Caption         =   "Audit Segments"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   8610
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   7080
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   3600
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   105
         TabIndex        =   24
         Top             =   195
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   23
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
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   11
      Top             =   2610
      Width           =   6840
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   0
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   17
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1425
         TabIndex        =   2
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1425
         TabIndex        =   16
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2745
         TabIndex        =   15
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   405
         Left            =   2745
         TabIndex        =   14
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4065
         TabIndex        =   13
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4065
         TabIndex        =   12
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5385
         TabIndex        =   4
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5385
         TabIndex        =   3
         Top             =   645
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6840
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
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1035
         Width           =   4680
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
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4680
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1995
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.TextBox TxtSegmentType 
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Width           =   4680
      End
      Begin VB.TextBox TxtSegmentTypeID 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   2040
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
         Left            =   255
         TabIndex        =   21
         Top             =   1080
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
         Left            =   255
         TabIndex        =   20
         Top             =   1515
         Width           =   1695
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
         TabIndex        =   10
         Top             =   2018
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label LblStateName 
         Caption         =   "Segment Type"
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
         TabIndex        =   9
         Top             =   675
         Width           =   1800
      End
      Begin VB.Label LblStateID 
         Caption         =   "Segment Type ID"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmMst_AuditSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtSegmentType.Locked = False
TxtSegmentType = ""
TxtSegmentTypeID = ""
TxtFlag = ""
TxtCreated = ""
TxtUpdated = ""

If RsMst_AuditSegment.RecordCount > 0 Then
   TxtSegmentType.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub


Private Sub DeleteCmd_Click()



Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   RsMst_AuditSegment.Delete
   RsMst_AuditSegment.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_AuditSegment.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_AuditSegment.MoveNext
   If RsMst_AuditSegment.EOF() Then
      RsMst_AuditSegment.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_AuditSegment.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtSegmentType.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_AuditSegment.MoveFirst
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


Call TableMst_AuditSegment("  Order by SegmentTypeID ")
If RsMst_AuditSegment.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtSegmentType.Locked = True
TxtSegmentTypeID = RsMst_AuditSegment!SegmentTypeID
TxtSegmentType = RsMst_AuditSegment!SegmentType
TxtCreated = IIf(IsNull(RsMst_AuditSegment!CreatedBy), "", RsMst_AuditSegment!CreatedBy) & " :- " & IIf(IsNull(RsMst_AuditSegment!CreatedDate), "", RsMst_AuditSegment!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_AuditSegment!UpdatedBy), "", RsMst_AuditSegment!UpdatedBy) & " :- " & IIf(IsNull(RsMst_AuditSegment!UpdatedDate), "", RsMst_AuditSegment!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_AuditSegment.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_AuditSegment.MoveFirst

If RsMst_AuditSegment.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_AuditSegment.MoveFirst
RsMst_AuditSegment.Find "SegmentTypeID = " & Val(mCaId) & ""
If RsMst_AuditSegment.EOF Then
   RsMst_AuditSegment.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_AuditSegment.MoveNext
If RsMst_AuditSegment.EOF Then
   RsMst_AuditSegment.MoveLast
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
RsMst_AuditSegment.MovePrevious
If RsMst_AuditSegment.BOF Then
   RsMst_AuditSegment.MoveFirst
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

If TxtSegmentType = "" Then
   MsgBox "Blank State Name  not Allowed"
   TxtSegmentType.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_AuditSegment.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_AuditSegment.MoveFirst
        RsMst_AuditSegment.Find "SegmentType = '" & TxtSegmentType & "'"
        If Not RsMst_AuditSegment.EOF Then
           MsgBox "Duplicate Segment Type not Allowed"
           Exit Sub
        End If
        RsMst_AuditSegment.MoveLast
        
        tmp = "select max(SegmentTypeID) from Mst_AuditSegment"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_AuditSegment.AddNew
    RsMst_AuditSegment!SegmentTypeID = tmp
    TxtSegmentTypeID = tmp
    RsMst_AuditSegment!G_UID = GetGUID
Else
    RsMst_AuditSegment.MoveFirst
    RsMst_AuditSegment.Find "SegmentType = '" & TxtSegmentType & "'"
    If Not RsMst_AuditSegment.EOF Then
       If RsMst_AuditSegment!SegmentTypeID <> TxtSegmentTypeID Then
          MsgBox "Duplicate Segment Type not Allowed"
          Exit Sub
       End If
    End If
   RsMst_AuditSegment.MoveFirst
   RsMst_AuditSegment.Find "SegmentTypeID = " & TxtSegmentTypeID & ""
End If
RsMst_AuditSegment!SegmentType = TxtSegmentType


If IsNull(RsMst_AuditSegment!CreatedBy) = True Or RsMst_AuditSegment!CreatedBy = "" Then
   RsMst_AuditSegment!CreatedBy = Gen_UserLoginID
   RsMst_AuditSegment!CreatedDate = Now
Else
   RsMst_AuditSegment!UpdatedBy = Gen_UserLoginID
   RsMst_AuditSegment!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_AuditSegment.Update
RsMst_AuditSegment.Requery
RsMst_AuditSegment.MoveFirst
RsMst_AuditSegment.Find "SegmentTypeID = " & TxtSegmentTypeID & ""

Call GButtonLock(Me, True)
TxtSegmentType.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_AuditSegment.EOF Then
      RsMst_AuditSegment.MoveFirst
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

RsMst_AuditSegment.MoveFirst
MSHFlexGrid1.Rows = RsMst_AuditSegment.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "Segment Type Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Segment Type"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4500
For I = 1 To RsMst_AuditSegment.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_AuditSegment.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_AuditSegment.Fields(1)
    RsMst_AuditSegment.MoveNext
Next


End Sub

Private Sub TxtSegmentType_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtSegmentType")
End If
End Sub

Private Sub TxtSegmentType_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtSegmentType)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtSegmentType.SetFocus
End If
End Sub


Private Sub TxtSegmentTypeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtSegmentTypeID")
End If
End Sub


Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Mst_AuditSegment.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\Mst_AuditSegment.xls")
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

