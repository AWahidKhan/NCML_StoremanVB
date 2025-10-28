VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_SecurityShifts 
   Caption         =   "Security Shifts"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   6660
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3495
      Left            =   6720
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   6450
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5100
         TabIndex        =   3
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3885
         TabIndex        =   19
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2550
         TabIndex        =   18
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1335
         TabIndex        =   17
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5100
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3885
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2550
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1335
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6450
      Begin VB.TextBox TxtShift 
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
      Begin VB.TextBox TxtShiftID 
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
         TabIndex        =   5
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Shift Type"
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
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Shift ID"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1695
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
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
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   3600
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMst_SecurityShifts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtShift.Locked = False
TxtShift = ""
TxtShiftID = ""

TxtCreated = ""
TxtUpdated = ""

If RsMst_SecurityShifts.RecordCount > 0 Then
   TxtShift.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   tmp = "Select * from Mst_SecurityDetails Where SecurityTypeID = " & TxtShiftID & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Child Record Present!!!"
      Exit Sub
   End If
   RsMst_SecurityShifts.Delete
   RsMst_SecurityShifts.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_SecurityShifts.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SecurityShifts.MoveNext
   If RsMst_SecurityShifts.EOF() Then
      RsMst_SecurityShifts.MoveLast
   End If
   Call Indata
   TmpRs.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_SecurityShifts.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtShift.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_SecurityShifts.MoveFirst
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

Call StatusBarData(Me)
Call TableMst_SecurityShifts(" Order by SecurityShiftsID ")

If RsMst_SecurityShifts.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtShift.Locked = True
TxtShiftID = RsMst_SecurityShifts!SecurityShiftsID
TxtShift = RsMst_SecurityShifts!SecurityShifts
TxtCreated = IIf(IsNull(RsMst_SecurityShifts!CreatedBy), "", RsMst_SecurityShifts!CreatedBy) & " :- " & IIf(IsNull(RsMst_SecurityShifts!CreatedDate), "", RsMst_SecurityShifts!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_SecurityShifts!UpdatedBy), "", RsMst_SecurityShifts!UpdatedBy) & " :- " & IIf(IsNull(RsMst_SecurityShifts!UpdatedDate), "", RsMst_SecurityShifts!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_SecurityShifts.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_SecurityShifts.MoveFirst

If RsMst_SecurityShifts.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_SecurityShifts.MoveFirst
RsMst_SecurityShifts.Find "SecurityShiftsID = " & Val(mCaId) & ""
If RsMst_SecurityShifts.EOF Then
   RsMst_SecurityShifts.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SecurityShifts.MoveNext
If RsMst_SecurityShifts.EOF Then
   RsMst_SecurityShifts.MoveLast
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
RsMst_SecurityShifts.MovePrevious
If RsMst_SecurityShifts.BOF Then
   RsMst_SecurityShifts.MoveFirst
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

If TxtShift = "" Then
   MsgBox "Blank Security Shifts  not Allowed"
   TxtShift.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_SecurityShifts.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_SecurityShifts.MoveFirst
        RsMst_SecurityShifts.Find "SecurityShifts = '" & TxtShift & "'"
        If Not RsMst_SecurityShifts.EOF Then
           MsgBox "Duplicate Security Shifts not Allowed"
           Exit Sub
        End If
        RsMst_SecurityShifts.MoveLast
        
        tmp = "select max(SecurityShiftsID) from Mst_SecurityShifts"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_SecurityShifts.AddNew
    RsMst_SecurityShifts!SecurityShiftsID = tmp
    TxtShiftID = tmp
Else
    RsMst_SecurityShifts.MoveFirst
    RsMst_SecurityShifts.Find "SecurityShifts = '" & TxtShift & "'"
    If Not RsMst_SecurityShifts.EOF Then
       If RsMst_SecurityShifts!SecurityShiftsID <> TxtShiftID Then
          MsgBox "Duplicate Security Shifts not Allowed"
          Exit Sub
       End If
    End If
   RsMst_SecurityShifts.MoveFirst
   RsMst_SecurityShifts.Find "SecurityShiftsID = " & TxtShiftID & ""
End If
RsMst_SecurityShifts!SecurityShifts = TxtShift


If IsNull(RsMst_SecurityShifts!CreatedBy) = True Or RsMst_SecurityShifts!CreatedBy = "" Then
   RsMst_SecurityShifts!CreatedBy = Gen_UserLoginID
   RsMst_SecurityShifts!CreatedDate = Now
Else
   RsMst_SecurityShifts!UpdatedBy = Gen_UserLoginID
   RsMst_SecurityShifts!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_SecurityShifts.Update
RsMst_SecurityShifts.Requery
RsMst_SecurityShifts.MoveFirst
RsMst_SecurityShifts.Find "SecurityShiftsID = " & TxtShiftID & ""

Call GButtonLock(Me, True)
TxtShift.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_SecurityShifts.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_SecurityShifts.EOF Then
      RsMst_SecurityShifts.MoveFirst
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
MSHFlexGrid1.Height = Frame0.Height - 500

RsMst_SecurityShifts.MoveFirst
MSHFlexGrid1.Rows = RsMst_SecurityShifts.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "Security Shift Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Security Shift Type"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4500
For I = 1 To RsMst_SecurityShifts.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_SecurityShifts.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_SecurityShifts.Fields(1)
    RsMst_SecurityShifts.MoveNext
Next

End Sub

Private Sub TxtShift_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtShift")
End If
End Sub

Private Sub TxtShift_LostFocus()
'Dim Ans As Boolean
'Ans = IsClean(TxtShift)
'If Ans = False Then
'   MsgBox "Special character not Allowed  !!!!!!!! "
'   TxtShift.SetFocus
'End If
End Sub



Private Sub TxtShiftID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtShiftID")
End If
End Sub
