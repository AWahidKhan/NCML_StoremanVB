VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_AssetStatus 
   Caption         =   "Asset Status"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   975
      Left            =   6600
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   6900
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1665
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox TxtAssetStatusID 
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtAssetStatus 
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
         TabIndex        =   21
         Top             =   1665
         Width           =   1695
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
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Asset Status ID"
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
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Asset Status"
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
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   6450
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1335
         TabIndex        =   2
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
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3885
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5100
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1335
         TabIndex        =   10
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2550
         TabIndex        =   9
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3885
         TabIndex        =   8
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5100
         TabIndex        =   4
         Top             =   615
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmMst_AssetStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtAssetStatus.Locked = False
TxtAssetStatus = ""
TxtAssetStatusID = ""
TxtCreated = ""
TxtUpdated = ""

If RsMst_AssetStatus.RecordCount > 0 Then
   TxtAssetStatus.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   'tmp = "Select * from Mst_EmpAssetMap Where (AssetStatusID = " & TxtAssetStatusID & ")"
   'Call Tmp1Table
   'If TmpRs1.RecordCount > 0 Then
   '   MsgBox "Child Record Present!!!"
   '   Exit Sub
   'End If
   RsMst_AssetStatus.Delete
   RsMst_AssetStatus.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_AssetStatus.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_AssetStatus.MoveNext
   If RsMst_AssetStatus.EOF() Then
      RsMst_AssetStatus.MoveLast
   End If
   Call Indata
   'TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_AssetStatus.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtAssetStatus.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_AssetStatus.MoveFirst
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


Call TableMst_AssetStatus(" Order by AssetStatus ")

If RsMst_AssetStatus.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtAssetStatus.Locked = True
TxtAssetStatusID = RsMst_AssetStatus!AssetsStatusID
TxtAssetStatus = RsMst_AssetStatus!AssetStatus
TxtCreated = IIf(IsNull(RsMst_AssetStatus!CreatedBy), "", RsMst_AssetStatus!CreatedBy) & " :- " & IIf(IsNull(RsMst_AssetStatus!CreatedDate), "", RsMst_AssetStatus!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_AssetStatus!UpdatedBy), "", RsMst_AssetStatus!UpdatedBy) & " :- " & IIf(IsNull(RsMst_AssetStatus!UpdatedDate), "", RsMst_AssetStatus!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_AssetStatus.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_AssetStatus.MoveFirst

If RsMst_AssetStatus.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_AssetStatus.MoveFirst
RsMst_AssetStatus.Find "AssetsStatusID = " & Val(mCaId) & ""
If RsMst_AssetStatus.EOF Then
   RsMst_AssetStatus.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_AssetStatus.MoveNext
If RsMst_AssetStatus.EOF Then
   RsMst_AssetStatus.MoveLast
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
RsMst_AssetStatus.MovePrevious
If RsMst_AssetStatus.BOF Then
   RsMst_AssetStatus.MoveFirst
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

If TxtAssetStatus = "" Then
   MsgBox "Blank AssetStatus  not Allowed"
   TxtAssetStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_AssetStatus.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_AssetStatus.MoveFirst
        RsMst_AssetStatus.Find "AssetStatus = '" & TxtAssetStatus & "'"
        If Not RsMst_AssetStatus.EOF Then
           MsgBox "Duplicate AssetStatus not Allowed"
           Exit Sub
        End If
        RsMst_AssetStatus.MoveLast
        
        tmp = "select max(AssetsStatusID) from Mst_AssetStatus"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_AssetStatus.AddNew
    RsMst_AssetStatus!AssetsStatusID = tmp
    TxtAssetStatusID = tmp
Else
    RsMst_AssetStatus.MoveFirst
    RsMst_AssetStatus.Find "AssetStatus = '" & TxtAssetStatus & "'"
    If Not RsMst_AssetStatus.EOF Then
       If RsMst_AssetStatus!AssetsStatusID <> TxtAssetStatusID Then
          MsgBox "Duplicate AssetStatus not Allowed"
          Exit Sub
       End If
    End If
   RsMst_AssetStatus.MoveFirst
   RsMst_AssetStatus.Find "AssetsStatusID = " & TxtAssetStatusID & ""
End If
RsMst_AssetStatus!AssetStatus = TxtAssetStatus


If IsNull(RsMst_AssetStatus!CreatedBy) = True Or RsMst_AssetStatus!CreatedBy = "" Then
   RsMst_AssetStatus!CreatedBy = Gen_UserLoginID
   RsMst_AssetStatus!CreatedDate = Now
Else
   RsMst_AssetStatus!UpdatedBy = Gen_UserLoginID
   RsMst_AssetStatus!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_AssetStatus.Update
RsMst_AssetStatus.Requery
RsMst_AssetStatus.MoveFirst
RsMst_AssetStatus.Find "AssetsStatusID = " & TxtAssetStatusID & ""

Call GButtonLock(Me, True)
TxtAssetStatus.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_AssetStatus.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_AssetStatus.EOF Then
      RsMst_AssetStatus.MoveFirst
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
MSHFlexGrid1.Height = Frame0.Height - 500

RsMst_AssetStatus.MoveFirst
MSHFlexGrid1.Rows = RsMst_AssetStatus.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "AssetStatus Id"
MSHFlexGrid1.TextMatrix(0, 1) = "AssetStatus"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4500
For I = 1 To RsMst_AssetStatus.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_AssetStatus.Fields(0)
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_AssetStatus.Fields(1)
    RsMst_AssetStatus.MoveNext
Next

End Sub

Private Sub TxtAssetStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetStatus")
End If
End Sub

Private Sub TxtAssetStatus_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAssetStatus)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAssetStatus.SetFocus
End If
End Sub



Private Sub TxtAssetStatusID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetStatusID")
End If
End Sub
