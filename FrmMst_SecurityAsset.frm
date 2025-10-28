VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_SecurityAsset 
   Caption         =   "Security Assets"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   975
      Left            =   7320
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   1
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
      Left            =   0
      TabIndex        =   11
      Top             =   2400
      Width           =   6480
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5100
         TabIndex        =   21
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3885
         TabIndex        =   20
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2550
         TabIndex        =   19
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1335
         TabIndex        =   18
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5100
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3885
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2550
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1335
         TabIndex        =   13
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6465
      Begin VB.TextBox TxtSecurityAsset 
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
         TabIndex        =   6
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox TxtSecurityAssetID 
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
         TabIndex        =   4
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
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Security Asset"
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
         Caption         =   "Security Asset ID"
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
End
Attribute VB_Name = "FrmMst_SecurityAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtSecurityAsset.Locked = False
TxtSecurityAsset = ""
TxtSecurityAssetID = ""

TxtCreated = ""
TxtUpdated = ""

If RsMst_SecurityAssets.RecordCount > 0 Then
   TxtSecurityAsset.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then

   tmp = "Select * from Mst_SecurityAssetMap Where (SecurityAssetID = " & TxtSecurityAssetID & ")"
   Call Tmp1Table
   
   If TmpRs1.RecordCount > 0 Then
      MsgBox "Child Record Present!!!"
      Exit Sub
   End If
   RsMst_SecurityAssets.Delete
   RsMst_SecurityAssets.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_SecurityAssets.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SecurityAssets.MoveNext
   If RsMst_SecurityAssets.EOF() Then
      RsMst_SecurityAssets.MoveLast
   End If
   Call Indata
   TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_SecurityAssets.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtSecurityAsset.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_SecurityAssets.MoveFirst
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

Call TableMst_SecurityAssets(" Order by SecurityAsset ")

If RsMst_SecurityAssets.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()

TxtSecurityAsset.Locked = True
TxtSecurityAssetID = RsMst_SecurityAssets!SecurityAssetID
TxtSecurityAsset = RsMst_SecurityAssets!SecurityAsset
TxtCreated = IIf(IsNull(RsMst_SecurityAssets!CreatedBy), "", RsMst_SecurityAssets!CreatedBy) & " :- " & IIf(IsNull(RsMst_SecurityAssets!CreatedDate), "", RsMst_SecurityAssets!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_SecurityAssets!UpdatedBy), "", RsMst_SecurityAssets!UpdatedBy) & " :- " & IIf(IsNull(RsMst_SecurityAssets!UpdatedDate), "", RsMst_SecurityAssets!UpdatedDate)
Call GButtonLock(Me, True)

End Sub

Private Sub LastCmd_Click()

RsMst_SecurityAssets.MoveLast
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
RsMst_SecurityAssets.MoveFirst

If RsMst_SecurityAssets.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_SecurityAssets.MoveFirst
RsMst_SecurityAssets.Find "SecurityAssetID = " & Val(mCaId) & ""
If RsMst_SecurityAssets.EOF Then
   RsMst_SecurityAssets.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SecurityAssets.MoveNext
If RsMst_SecurityAssets.EOF Then
   RsMst_SecurityAssets.MoveLast
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
RsMst_SecurityAssets.MovePrevious
If RsMst_SecurityAssets.BOF Then
   RsMst_SecurityAssets.MoveFirst
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

If TxtSecurityAsset = "" Then
   MsgBox "Blank SecurityAsset  not Allowed"
   TxtSecurityAsset.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_SecurityAssets.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_SecurityAssets.MoveFirst
        RsMst_SecurityAssets.Find "SecurityAsset = '" & TxtSecurityAsset & "'"
        If Not RsMst_SecurityAssets.EOF Then
           MsgBox "Duplicate SecurityAsset not Allowed"
           Exit Sub
        End If
        RsMst_SecurityAssets.MoveLast
        
        tmp = "select max(SecurityAssetID) from Mst_SecurityAssets"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_SecurityAssets.AddNew
    RsMst_SecurityAssets!SecurityAssetID = tmp
    TxtSecurityAssetID = tmp
Else
    RsMst_SecurityAssets.MoveFirst
    RsMst_SecurityAssets.Find "SecurityAsset = '" & TxtSecurityAsset & "'"
    If Not RsMst_SecurityAssets.EOF Then
       If RsMst_SecurityAssets!SecurityAssetID <> TxtSecurityAssetID Then
          MsgBox "Duplicate SecurityAsset not Allowed"
          Exit Sub
       End If
    End If
   RsMst_SecurityAssets.MoveFirst
   RsMst_SecurityAssets.Find "SecurityAssetID = " & TxtSecurityAssetID & ""
End If
RsMst_SecurityAssets!SecurityAsset = TxtSecurityAsset


If IsNull(RsMst_SecurityAssets!CreatedBy) = True Or RsMst_SecurityAssets!CreatedBy = "" Then
   RsMst_SecurityAssets!CreatedBy = Gen_UserLoginID
   RsMst_SecurityAssets!CreatedDate = Now
Else
   RsMst_SecurityAssets!UpdatedBy = Gen_UserLoginID
   RsMst_SecurityAssets!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_SecurityAssets.Update
RsMst_SecurityAssets.Requery
RsMst_SecurityAssets.MoveFirst
RsMst_SecurityAssets.Find "SecurityAssetID = " & TxtSecurityAssetID & ""

Call GButtonLock(Me, True)
TxtSecurityAsset.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_SecurityAssets.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_SecurityAssets.EOF Then
      RsMst_SecurityAssets.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500

Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
'Frame0.Width = Frame1.Width
Frame0.Visible = True
MSHFlexGrid1.Clear

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 500

RsMst_SecurityAssets.MoveFirst
MSHFlexGrid1.Rows = RsMst_SecurityAssets.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "Security Asset Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Security Asset"

MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 4000
For I = 1 To RsMst_SecurityAssets.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_SecurityAssets.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_SecurityAssets.Fields(1)
    RsMst_SecurityAssets.MoveNext
Next

End Sub

Private Sub TxtSecurityAsset_KeyPress(KeyAscii As Integer)

If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSecurityAsset")
End If

End Sub

Private Sub TxtSecurityAsset_LostFocus()

Dim ans As Boolean
ans = IsClean(TxtSecurityAsset)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtSecurityAsset.SetFocus
End If

End Sub

Private Sub TxtSecurityAssetID_KeyPress(KeyAscii As Integer)

If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSecurityAssetID")
End If

End Sub
