VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_AuditPoint 
   Caption         =   "Audit Points"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   9120
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   3600
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
         Height          =   350
         Left            =   90
         TabIndex        =   25
         Top             =   195
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   26
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
      Height          =   4335
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   8685
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3120
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3555
         Width           =   4680
      End
      Begin VB.ComboBox CmbAuditSegmentID 
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
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   2205
         Width           =   6315
      End
      Begin VB.CommandButton CmdAuditSegment 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8160
         TabIndex        =   29
         Top             =   2160
         Width           =   375
      End
      Begin VB.ComboBox CmbRequired 
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
         ItemData        =   "Frm_Mst_AuditPoint.frx":0000
         Left            =   1785
         List            =   "Frm_Mst_AuditPoint.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2640
         Width           =   6315
      End
      Begin VB.CommandButton CmdAuditTypeID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8160
         TabIndex        =   18
         Top             =   1770
         Width           =   375
      End
      Begin VB.ComboBox CmbAuditTypeID 
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
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1770
         Width           =   6315
      End
      Begin VB.TextBox TxtAuditPointID 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   2280
      End
      Begin VB.TextBox TxtAuditPoint 
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
         Height          =   1080
         Left            =   1800
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   645
         Width           =   6300
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4215
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.TextBox TxtAuditTypeID 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1770
         Width           =   6300
      End
      Begin VB.TextBox TxtRequired 
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
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2640
         Width           =   6300
      End
      Begin VB.TextBox TxtAuditSegmentID 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2205
         Width           =   6300
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
         TabIndex        =   35
         Top             =   3165
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
         TabIndex        =   34
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Audit Segment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Top             =   2273
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Required"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   27
         Top             =   2708
         Width           =   1800
      End
      Begin VB.Label LblAuditPointID 
         Caption         =   "Audit Point ID"
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
         TabIndex        =   23
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label LblAuditPoint 
         Caption         =   "Audit Point"
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
         TabIndex        =   22
         Top             =   1058
         Width           =   1800
      End
      Begin VB.Label LblAuditTypeID 
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
         Height          =   225
         Left            =   240
         TabIndex        =   21
         Top             =   1838
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   225
         Left            =   225
         TabIndex        =   20
         Top             =   4245
         Visible         =   0   'False
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   0
      Top             =   4500
      Width           =   8685
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   14
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1740
         TabIndex        =   13
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3375
         TabIndex        =   12
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5010
         TabIndex        =   11
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   6645
         TabIndex        =   6
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   10
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1740
         TabIndex        =   5
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3375
         TabIndex        =   9
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5010
         TabIndex        =   8
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   6645
         TabIndex        =   7
         Top             =   240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "FrmMst_AuditPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mAuditTypeID, mAuditSegmentID, mRequiredTypeID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtAuditPoint.Locked = False
TxtAuditPoint = ""
TxtAuditPointID = ""

TxtAuditTypeID.Visible = False
CmdAuditTypeID.Visible = True
CmbAuditTypeID.Visible = True
CmbAuditTypeID.Text = ""

TxtAuditSegmentID.Visible = False
CmdAuditSegment.Visible = True
CmbAuditSegmentID.Visible = True
CmbAuditSegmentID.Text = ""


TxtRequired.Visible = False
CmbRequired.Visible = True
CmbRequired.Text = ""

TxtCreated = ""
TxtUpdated = ""
'RsMst_AuditPoint
If RsMst_AuditPoint.RecordCount > 0 Then
   TxtAuditPoint.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbAuditTypeID_GotFocus()
tmp = CmbAuditTypeID.Text
Call TableMst_AuditType("")

CmbAuditTypeID.Clear
If RsMst_AuditType.RecordCount = 0 Then
   MsgBox "No AuditType found"
   CmdAuditTypeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditType.RecordCount
    CmbAuditTypeID.AddItem RsMst_AuditType!AuditType
    RsMst_AuditType.MoveNext
Next
CmbAuditTypeID.Text = tmp
End Sub

Private Sub CmbAuditTypeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAuditTypeID")
End If
End Sub

Private Sub CmbAuditTypeID_LostFocus()
If CmbAuditTypeID.Text = "" Then
   Exit Sub
End If
RsMst_AuditType.MoveFirst
RsMst_AuditType.Find "AuditType = '" & CmbAuditTypeID.Text & "'"

If RsMst_AuditType.EOF Then
   MsgBox "Invalid AuditType selection "
   CmbAuditTypeID.SetFocus
   Exit Sub
End If
mAuditTypeID = RsMst_AuditType.Fields(0)

End Sub

Private Sub CmdAuditTypeID_Click()
FrmMstAuditType.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   RsMst_AuditPoint.Delete
   RsMst_AuditPoint.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_AuditPoint.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_AuditPoint.MoveNext
   If RsMst_AuditPoint.EOF() Then
      RsMst_AuditPoint.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_AuditPoint.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"

TxtAuditTypeID.Visible = False
CmdAuditTypeID.Visible = True
CmbAuditTypeID.Visible = True
CmbAuditTypeID.Text = TxtAuditTypeID

TxtAuditSegmentID.Visible = False
CmdAuditSegment.Visible = True
CmbAuditSegmentID.Visible = True
CmbAuditSegmentID.Text = TxtAuditSegmentID

TxtRequired.Visible = False
CmbRequired.Visible = True
CmbRequired.Text = TxtRequired

Call GButtonLock(Me, False)
TxtAuditPoint.Locked = False

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_AuditPoint.MoveFirst
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
 
Call TableMst_AuditPoint(" Order by AuditPointID")
If RsMst_AuditPoint.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtAuditPoint.Locked = True
TxtAuditTypeID.Visible = True
CmdAuditTypeID.Visible = False
CmbAuditTypeID.Visible = False
TxtAuditTypeID.Locked = True
TxtAuditType = ""
TxtAuditPoint = ""
TxtAuditPointID = ""
mAuditTypeID = Null
mAuditSegmentID = Null
mRequiredTypeID = Null

TxtAuditSegmentID.Visible = True
CmdAuditSegment.Visible = False
CmbAuditSegmentID.Visible = False


TxtRequired.Visible = True
CmbRequired.Visible = False



TxtAuditPointID = IIf(IsNull(RsMst_AuditPoint!AuditPointID), "", RsMst_AuditPoint!AuditPointID)
TxtAuditPoint = IIf(IsNull(RsMst_AuditPoint!AuditPoint), "", RsMst_AuditPoint!AuditPoint) 'RsMst_AuditPoint!AuditPoint
mAuditTypeID = RsMst_AuditPoint!AuditTypeID
TxtAuditTypeID = GetAuditType(mAuditTypeID)
mAuditSegmentID = RsMst_AuditPoint!SegmentTypeID
TxtAuditSegmentID = GetAuditSegmentType(mAuditSegmentID)
mRequiredTypeID = RsMst_AuditPoint!RequiredType
If mRequiredTypeID = "D" Then
   TxtRequired = "Date"
ElseIf mRequiredTypeID = "N" Then
   TxtRequired = "Number"
ElseIf mRequiredTypeID = "V" Then
   TxtRequired = "Value"
End If

TxtCreated = IIf(IsNull(RsMst_AuditPoint!CreatedBy), "", RsMst_AuditPoint!CreatedBy) & " :- " & IIf(IsNull(RsMst_AuditPoint!CreatedDate), "", RsMst_AuditPoint!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_AuditPoint!UpdatedBy), "", RsMst_AuditPoint!UpdatedBy) & " :- " & IIf(IsNull(RsMst_AuditPoint!UpdatedDate), "", RsMst_AuditPoint!UpdatedDate)


If mAuditTypeID <> "" Then
    tmp = "select * from Mst_AuditType where AuditTypeID = " & mAuditTypeID & ""
    Call TmpTable
    
    If TmpRs.RecordCount > 0 Then
       TxtAuditTypeID = TmpRs!AuditType
    End If
End If
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_AuditPoint.MoveLast
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

RsMst_AuditPoint.MoveFirst

If RsMst_AuditPoint.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_AuditPoint.MoveFirst
RsMst_AuditPoint.Find "AuditPointID = " & Val(mCaId) & ""
If RsMst_AuditPoint.EOF Then
   RsMst_AuditPoint.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_AuditPoint.MoveNext
If RsMst_AuditPoint.EOF Then
   RsMst_AuditPoint.MoveLast
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

RsMst_AuditPoint.MovePrevious
If RsMst_AuditPoint.BOF Then
   RsMst_AuditPoint.MoveFirst
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
Private Function GetMaxAuditPointID() As Double
Dim Retval As Variant
Retval = 0
tmp = "select max(AuditPointID) from Mst_AuditPoint  "
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs4.Fields(0)), 0, TmpRs4.Fields(0))
End If
TmpRs.Close
Retval = Retval + 1

GetMaxAuditPointID = Retval
End Function

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtAuditPoint = "" Then
   MsgBox "Blank Audit Point  not Allowed"
   TxtAuditPoint.SetFocus
   Exit Sub
End If


If CmbAuditTypeID.Text = "" Then
   MsgBox "Blank Audit Type not Allowed"
   CmbAuditTypeID.SetFocus
   Exit Sub
End If

If CmbAuditSegmentID.Text = "" Then
   MsgBox "Blank Audit Segment not Allowed"
   CmbAuditSegmentID.SetFocus
   Exit Sub
End If

If CmbRequired.Text = "" Then
   MsgBox "Blank Required not Allowed"
   CmbRequired.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsMst_AuditPoint.RecordCount > 0 Then
      RsMst_AuditPoint.MoveFirst
      tmp = "AuditPoint = '" & TxtAuditPoint.Text & "' And AuditTypeID = " & mAuditTypeID & " And SegmentTypeID = " & mAuditSegmentID & ""
      'RsMst_AuditPoint.Find "ClientID = '" & TxtClientID.Text & "'"
      RsMst_AuditPoint.Filter = tmp
      If Not RsMst_AuditPoint.EOF Then
         MsgBox "Duplicate Audit Point Not Allowed !!!"
         TxtAuditPoint.SetFocus
         RsMst_AuditPoint.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_AuditPoint.Filter = ""
   RsMst_AuditPoint.AddNew
   RsMst_AuditPoint!AuditPointID = GetMaxAuditPointID
   TxtAuditPointID = RsMst_AuditPoint!AuditPointID
   RsMst_AuditPoint!G_UID = GetGUID
Else
   RsMst_AuditPoint.MoveFirst
   tmp = "AuditPoint = '" & TxtAuditPoint.Text & "' And AuditTypeID = " & mAuditTypeID & " And SegmentTypeID = " & mAuditSegmentID & ""
   'RsMst_AuditPoint.Find "ClientID = '" & TxtClientID.Text & "'"
   RsMst_AuditPoint.Filter = tmp
   If Not RsMst_AuditPoint.EOF Then
      If RsMst_AuditPoint!AuditPointID <> TxtAuditPointID.Text Then
         MsgBox "Duplicate Audit Point Not Allowed !!! "
         TxtAuditPoint.SetFocus
         Exit Sub
      End If
   End If
   RsMst_AuditPoint.Filter = ""
   RsMst_AuditPoint.MoveFirst
   RsMst_AuditPoint.Find "AuditPointID = " & TxtAuditPointID
   If RsMst_AuditPoint.EOF Then
      MsgBox "No Such Record Found!!! "
      Exit Sub
   End If
End If


RsMst_AuditPoint!AuditPoint = IIf(IsNull(TxtAuditPoint), "", TxtAuditPoint)
RsMst_AuditPoint!AuditTypeID = mAuditTypeID
RsMst_AuditPoint!SegmentTypeID = mAuditSegmentID
RsMst_AuditPoint!RequiredType = mRequiredTypeID

If IsNull(RsMst_AuditPoint!CreatedBy) = True Or RsMst_AuditPoint!CreatedBy = "" Then
   RsMst_AuditPoint!CreatedBy = Gen_UserLoginID
   RsMst_AuditPoint!CreatedDate = Now
Else
   RsMst_AuditPoint!UpdatedBy = Gen_UserLoginID
   RsMst_AuditPoint!UpdatedDate = Now
End If

RsMst_AuditPoint.Update
Call GButtonLock(Me, True)
TxtAuditPoint.Locked = True
RsMst_AuditPoint.Requery
RsMst_AuditPoint.MoveFirst
RsMst_AuditPoint.Find "AuditPointID = " & TxtAuditPointID & ""
If RsMst_AuditPoint.EOF Then
   RsMst_AuditPoint.MoveLast
End If
Call Indata

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_AuditPoint.EOF Then
      RsMst_AuditPoint.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Left = 120
Frame0.Top = 120

Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500
Frame0.Visible = True
MSHFlexGrid1.Clear

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 800

RsMst_AuditPoint.MoveFirst
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.Rows = RsMst_AuditPoint.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "Audit Point Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Audit Point "
MSHFlexGrid1.TextMatrix(0, 2) = "Audit Type "
MSHFlexGrid1.TextMatrix(0, 3) = "Audit Segment "
MSHFlexGrid1.TextMatrix(0, 4) = "Required"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 8000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400


For I = 1 To RsMst_AuditPoint.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_AuditPoint!AuditPointID
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_AuditPoint!AuditPoint
    MSHFlexGrid1.TextMatrix(I, 2) = GetAuditType(RsMst_AuditPoint!AuditTypeID)
    MSHFlexGrid1.TextMatrix(I, 3) = GetAuditSegmentType(RsMst_AuditPoint!SegmentTypeID)
    If RsMst_AuditPoint!RequiredType = "D" Then
       MSHFlexGrid1.TextMatrix(I, 4) = "Date" 'GetAuditTypeName(RsMst_AuditPoint!AuditTypeID)
    ElseIf RsMst_AuditPoint!RequiredType = "N" Then
       MSHFlexGrid1.TextMatrix(I, 4) = "Number"
    ElseIf RsMst_AuditPoint!RequiredType = "V" Then
       MSHFlexGrid1.TextMatrix(I, 4) = "Value"
    End If
    RsMst_AuditPoint.MoveNext
Next


End Sub

Private Sub TxtAuditPoint_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAuditPoint")
End If
End Sub

Private Sub TxtAuditPoint_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAuditPoint)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAuditPoint.SetFocus
End If
End Sub
Private Sub TxtAuditPointID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAuditPointID")
End If
End Sub

Private Sub TxtAuditType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAuditType")
End If
End Sub


Private Sub CmbAuditSegmentID_GotFocus()
tmp = CmbAuditSegmentID.Text
Call TableMst_AuditSegment("")

CmbAuditSegmentID.Clear
If RsMst_AuditSegment.RecordCount = 0 Then
   MsgBox "No Audit Segment found"
   CmbAuditSegmentID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditSegment.RecordCount
    CmbAuditSegmentID.AddItem RsMst_AuditSegment!SegmentType
    RsMst_AuditSegment.MoveNext
Next
CmbAuditSegmentID.Text = tmp
End Sub

Private Sub CmbAuditSegmentID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAuditSegmentID")
End If
End Sub

Private Sub CmbAuditSegmentID_LostFocus()
If CmbAuditSegmentID.Text = "" Then
   mAuditSegmentID = Null
   Exit Sub
End If
RsMst_AuditSegment.MoveFirst
RsMst_AuditSegment.Find "SegmentType = '" & CmbAuditSegmentID.Text & "'"

If RsMst_AuditSegment.EOF Then
   MsgBox "Invalid Segment selection "
   CmbAuditSegmentID.SetFocus
   Exit Sub
End If
mAuditSegmentID = RsMst_AuditSegment!SegmentTypeID

End Sub

Private Sub CmdAuditSegment_Click()
FrmMstAuditSegment.Show
End Sub

Private Sub CmbRequired_GotFocus()

tmp = CmbRequired.Text
CmbRequired.Text = tmp

End Sub

Private Sub CmbRequired_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbRequired")
End If
End Sub

Private Sub CmbRequired_LostFocus()
If CmbRequired = "" Then
   mRequiredTypeID = Null
   Exit Sub
End If

If CmbRequired.Text = "Date" Then
   mRequiredTypeID = "D"
ElseIf CmbRequired.Text = "Number" Then
   mRequiredTypeID = "N"
ElseIf CmbRequired.Text = "Value" Then
   mRequiredTypeID = "V"
Else
   MsgBox "Invalid Selection"
   mRequiredTypeID = Null
   CmbRequired.SetFocus
End If

End Sub


Private Sub CmdExcel_Click()

Call Gen_Create_Xls("Mst_AuditPoint.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\Mst_AuditPoint.xls")
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
