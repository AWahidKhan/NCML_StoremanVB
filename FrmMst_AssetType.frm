VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_AssetType 
   Caption         =   "Asset Type"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   8220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1095
      Left            =   6840
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   17
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
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6375
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
         Top             =   1680
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
      Begin VB.TextBox TxtAssetTypeID 
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtAssetType 
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
         TabIndex        =   21
         Top             =   1680
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
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Asset Type ID"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Asset Type"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   6405
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
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3885
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5100
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1335
         TabIndex        =   8
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2550
         TabIndex        =   7
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3885
         TabIndex        =   6
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5100
         TabIndex        =   3
         Top             =   615
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmMst_AssetType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtAssetType.Locked = False
TxtAssetType = ""
TxtAssetTypeID = ""
TxtCreated = ""
TxtUpdated = ""
If RsMst_AssetType.RecordCount > 0 Then
   TxtAssetType.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub DeleteCmd_Click()

Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If Ans = vbYes Then
'   tmp = "Select * from Mst_EmpAssetMap Where (AssetID = " & TxtAssetTypeID & ")"
'   Call Tmp1Table
'   If TmpRs1.RecordCount > 0 Then
'      MsgBox "Child Record Present!!!"
'      Exit Sub
'   End If
   RsMst_AssetType.Delete
   RsMst_AssetType.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_AssetType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_AssetType.MoveNext
   If RsMst_AssetType.EOF() Then
      RsMst_AssetType.MoveLast
   End If
   Call InData
   'TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_AssetType.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtAssetType.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_AssetType.MoveFirst
Call InData
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

Call TableMst_AssetType(" Order by AssetsID ")

If RsMst_AssetType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call InData
Call GButtonLock(Me, True)

End Sub

Public Sub InData()
TxtAssetType.Locked = True
TxtAssetTypeID = RsMst_AssetType!AssetsID
TxtAssetType = RsMst_AssetType!AssetType
TxtCreated = IIf(IsNull(RsMst_AssetType!CreatedBy), "", RsMst_AssetType!CreatedBy) & " :- " & IIf(IsNull(RsMst_AssetType!CreatedDate), "", RsMst_AssetType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_AssetType!UpdatedBy), "", RsMst_AssetType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_AssetType!UpdatedDate), "", RsMst_AssetType!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_AssetType.MoveLast
Call InData
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_AssetType.MoveFirst

If RsMst_AssetType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_AssetType.MoveFirst
RsMst_AssetType.Find "AssetsID = " & Val(mCaId) & ""
If RsMst_AssetType.EOF Then
   RsMst_AssetType.MoveLast
End If
Frame0.Visible = False
Call InData

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_AssetType.MoveNext
If RsMst_AssetType.EOF Then
   RsMst_AssetType.MoveLast
   LF_Flag = "Y"
End If
Call InData
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
RsMst_AssetType.MovePrevious
If RsMst_AssetType.BOF Then
   RsMst_AssetType.MoveFirst
   LF_Flag = "Y"
End If
Call InData
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

If TxtAssetType = "" Then
   MsgBox "Blank AssetsType  not Allowed"
   TxtAssetType.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_AssetType.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_AssetType.MoveFirst
        RsMst_AssetType.Find "AssetType = '" & TxtAssetType & "'"
        If Not RsMst_AssetType.EOF Then
           MsgBox "Duplicate Asset Type not Allowed"
           Exit Sub
        End If
        RsMst_AssetType.MoveLast
        
        tmp = "select max(AssetsID) from Mst_AssetsType"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_AssetType.AddNew
    RsMst_AssetType!AssetsID = tmp
    TxtAssetTypeID = tmp
Else
    RsMst_AssetType.MoveFirst
    RsMst_AssetType.Find "AssetType = '" & TxtAssetType & "'"
    If Not RsMst_AssetType.EOF Then
       If RsMst_AssetType!AssetsID <> TxtAssetTypeID Then
          MsgBox "Duplicate Asset Type not Allowed"
          Exit Sub
       End If
    End If
   RsMst_AssetType.MoveFirst
   RsMst_AssetType.Find "AssetsID = " & TxtAssetTypeID & ""
End If
RsMst_AssetType!AssetType = TxtAssetType


If IsNull(RsMst_AssetType!CreatedBy) = True Or RsMst_AssetType!CreatedBy = "" Then
   RsMst_AssetType!CreatedBy = Gen_UserLoginID
   RsMst_AssetType!CreatedDate = Now
Else
   RsMst_AssetType!UpdatedBy = Gen_UserLoginID
   RsMst_AssetType!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If


RsMst_AssetType.Update
RsMst_AssetType.Requery
RsMst_AssetType.MoveFirst
RsMst_AssetType.Find "AssetsID = " & TxtAssetTypeID & ""

Call GButtonLock(Me, True)
TxtAssetType.Locked = True
Call InData
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_AssetType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_AssetType.EOF Then
      RsMst_AssetType.MoveFirst
   End If
   Call InData
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

RsMst_AssetType.MoveFirst
MSHFlexGrid1.Rows = RsMst_AssetType.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "AssetsType Id"
MSHFlexGrid1.TextMatrix(0, 1) = "AssetsType"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4500
For i = 1 To RsMst_AssetType.RecordCount
    MSHFlexGrid1.TextMatrix(i, 0) = RsMst_AssetType.Fields(0)
    MSHFlexGrid1.TextMatrix(i, 1) = RsMst_AssetType.Fields(1)
    RsMst_AssetType.MoveNext
Next

End Sub

Private Sub TxtAssetType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetType")
End If
End Sub

Private Sub TxtAssetType_LostFocus()
Dim Ans As Boolean
Ans = IsClean(TxtAssetType)
If Ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAssetType.SetFocus
End If
End Sub

Private Sub TxtAssetTypeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAssetTypeID")
End If
End Sub
