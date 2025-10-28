VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_InsuranceBy 
   Caption         =   "Insurance By"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   10815
      Begin VB.ComboBox CmbSearchStatus 
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
         ItemData        =   "FrmMst_InsuranceBy.frx":0000
         Left            =   8520
         List            =   "FrmMst_InsuranceBy.frx":000A
         TabIndex        =   27
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxtSearchInsuranceBy 
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
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   8295
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8700
         TabIndex        =   24
         Top             =   880
         Width           =   1005
      End
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
         Left            =   9720
         TabIndex        =   25
         Top             =   880
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2100
         Left            =   60
         TabIndex        =   28
         Top             =   1275
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3704
         _Version        =   393216
         Rows            =   3
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   8280
         TabIndex        =   32
         Top             =   932
         Width           =   90
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
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
         Left            =   5880
         TabIndex        =   31
         Top             =   925
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Insurance By"
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
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   8295
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Status"
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
         Left            =   8520
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame ButtonFrm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   10800
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2100
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
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
         Left            =   2220
         TabIndex        =   14
         Top             =   240
         Width           =   2100
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
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
         Left            =   4320
         TabIndex        =   15
         Top             =   240
         Width           =   2100
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
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
         Left            =   6420
         TabIndex        =   16
         Top             =   240
         Width           =   2100
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
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
         Left            =   8540
         TabIndex        =   17
         Top             =   240
         Width           =   2100
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
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
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
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
         Left            =   2220
         TabIndex        =   19
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
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
         Left            =   4320
         TabIndex        =   20
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
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
         Left            =   6420
         TabIndex        =   21
         Top             =   600
         Width           =   2100
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
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
         Left            =   8540
         TabIndex        =   22
         Top             =   600
         Width           =   2100
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10815
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmMst_InsuranceBy.frx":0021
         Left            =   1920
         List            =   "FrmMst_InsuranceBy.frx":002B
         TabIndex        =   1
         Top             =   1200
         Width           =   8775
      End
      Begin VB.TextBox TxtStatus 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   8775
      End
      Begin VB.TextBox TxtInsuranceBy 
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
         Left            =   1920
         MaxLength       =   250
         TabIndex        =   0
         Top             =   720
         Width           =   8775
      End
      Begin VB.TextBox TxtInsuranceByID 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2160
         Width           =   8775
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1680
         Width           =   8775
      End
      Begin VB.Label Label1 
         Caption         =   "Insurance By ID"
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
         TabIndex        =   11
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Insurance By"
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
         TabIndex        =   10
         Top             =   780
         Width           =   2055
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
         TabIndex        =   9
         Top             =   2160
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
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Status"
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
         TabIndex        =   7
         Top             =   1260
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmMst_InsuranceBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtInsuranceBy.Locked = False
TxtInsuranceBy = ""
TxtInsuranceByID = ""
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = ""
TxtStatus.Text = ""
TxtCreated = ""
TxtUpdated = ""

If RsMst_InsuranceBy.RecordCount > 0 Then
   TxtInsuranceBy.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbSearchStatus_LostFocus()
If CmbSearchStatus.Text = "" Then Exit Sub
If LCase(CmbSearchStatus.Text) <> "active" And LCase(CmbSearchStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbSearchStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
wc = ""
Call Grid_Head
Label22.Caption = "-"

If TxtSearchInsuranceBy <> "" Then
   wc = " Where InsuranceByName like '%" & TxtSearchInsuranceBy & "%'"
End If
If CmbSearchStatus <> "" Then
    If wc = "" Then
        wc = " Where Flag = '" & Left(CmbSearchStatus, 1) & "'"
    Else
        wc = wc & " And Flag = '" & Left(CmbSearchStatus, 1) & "'"
    End If
End If
tmp = "select InsuranceByID,InsuranceByName,Flag from Mst_InsuranceBy"
tmp = tmp & wc
Call TmpTable

For iu = 1 To TmpRs.RecordCount
With MSHFlexGrid1
    .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!InsuranceByID), "", TmpRs!InsuranceByID)
    .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)
    .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
    If LCase(.TextMatrix(iu, 2)) = LCase("A") Then .TextMatrix(iu, 2) = "Active"
    If LCase(.TextMatrix(iu, 2)) = LCase("D") Then .TextMatrix(iu, 2) = "De-Active"
    .Rows = .Rows + 1
    TmpRs.MoveNext
End With
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   RsMst_InsuranceBy.Delete
   RsMst_InsuranceBy.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_InsuranceBy.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_InsuranceBy.MoveNext
   If RsMst_InsuranceBy.EOF() Then
      RsMst_InsuranceBy.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_InsuranceBy.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtInsuranceBy.Locked = False
TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus.Text
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_InsuranceBy.MoveFirst
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

CmbSearchStatus.Text = ""
TxtSearchInsuranceBy = ""
Call Grid_Head
Call TableMst_InsuranceBy
If RsMst_InsuranceBy.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub

Public Sub Indata()
TxtInsuranceBy = ""
TxtStatus = ""
CmbStatus = ""
TxtCreated = ""
TxtUpdated = ""

TxtInsuranceBy.Locked = True
TxtStatus.Visible = True
TxtStatus.Locked = True
CmbStatus.Visible = False
TxtInsuranceByID = RsMst_InsuranceBy!InsuranceByID
TxtInsuranceBy = RsMst_InsuranceBy!InsuranceByName
TxtCreated = IIf(IsNull(RsMst_InsuranceBy!CreatedBy), "", RsMst_InsuranceBy!CreatedBy) & " :- " & IIf(IsNull(RsMst_InsuranceBy!CreatedDate), "", RsMst_InsuranceBy!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_InsuranceBy!UpdatedBy), "", RsMst_InsuranceBy!UpdatedBy) & " :- " & IIf(IsNull(RsMst_InsuranceBy!UpdatedDate), "", RsMst_InsuranceBy!UpdatedDate)
If IIf(IsNull(RsMst_InsuranceBy!Flag), "", RsMst_InsuranceBy!Flag) = "A" Then
   TxtStatus.Text = "Active"
ElseIf IIf(IsNull(RsMst_InsuranceBy!Flag), "", RsMst_InsuranceBy!Flag) = "D" Then
   TxtStatus.Text = "De-Active"
Else
   TxtStatus.Text = ""
End If
Call GButtonLock(Me, True)
End Sub


Private Sub LastCmd_Click()
RsMst_InsuranceBy.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
With MSHFlexGrid1
RsMst_InsuranceBy.MoveFirst
If RsMst_InsuranceBy.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_InsuranceBy.MoveFirst
RsMst_InsuranceBy.Find "InsuranceByID = " & Val(.TextMatrix(.RowSel, 0)) & ""
If RsMst_InsuranceBy.EOF Then
   RsMst_InsuranceBy.MoveLast
End If
Frame0.Visible = False
Call Indata
End With
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_InsuranceBy.MoveNext
If RsMst_InsuranceBy.EOF Then
   RsMst_InsuranceBy.MoveLast
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
RsMst_InsuranceBy.MovePrevious
If RsMst_InsuranceBy.BOF Then
   RsMst_InsuranceBy.MoveFirst
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
If TxtInsuranceBy = "" Then
   MsgBox "Blank Insurance By not Allowed"
   TxtInsuranceBy.SetFocus
   Exit Sub
End If
If CmbStatus.Text = "" Then
   MsgBox "Blank Status not Allowed !!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_InsuranceBy.RecordCount > 0 Then
        RsMst_InsuranceBy.MoveFirst
        RsMst_InsuranceBy.Find "InsuranceByName = '" & TxtInsuranceBy & "'"
        If Not RsMst_InsuranceBy.EOF Then
           MsgBox "Duplicate Insurance By not Allowed"
           Exit Sub
        End If
   End If
   RsMst_InsuranceBy.AddNew
   RsMst_InsuranceBy!InsuranceByID = GetMaxID("InsuranceByID", "Mst_InsuranceBy")
   TxtInsuranceByID = RsMst_InsuranceBy!InsuranceByID
   RsMst_InsuranceBy!G_UID = GetGUID
ElseIf Edit_Flg = "E" Then
   RsMst_InsuranceBy.MoveFirst
   RsMst_InsuranceBy.Find "InsuranceByName = '" & TxtInsuranceBy & "'"
   If Not RsMst_InsuranceBy.EOF Then
      If RsMst_InsuranceBy!InsuranceByID <> TxtInsuranceByID Then
         MsgBox "Duplicate Insurance By not Allowed"
         Exit Sub
      End If
   End If
   RsMst_InsuranceBy.MoveFirst
   RsMst_InsuranceBy.Find "InsuranceByID = " & TxtInsuranceByID & ""
End If

RsMst_InsuranceBy!InsuranceByName = TxtInsuranceBy
RsMst_InsuranceBy!Flag = UCase(Left(CmbStatus.Text, 1))

If IsNull(RsMst_InsuranceBy!CreatedBy) = True Or RsMst_InsuranceBy!CreatedBy = "" Then
   RsMst_InsuranceBy!CreatedBy = Gen_UserLoginID
   RsMst_InsuranceBy!CreatedDate = Now
Else
   RsMst_InsuranceBy!UpdatedBy = Gen_UserLoginID
   RsMst_InsuranceBy!UpdatedDate = Now
End If

RsMst_InsuranceBy.Update
RsMst_InsuranceBy.Requery
RsMst_InsuranceBy.MoveFirst
RsMst_InsuranceBy.Find "InsuranceByID = " & TxtInsuranceByID & ""
Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 3
    .Rows = 2
    .WordWrap = True
    .TextMatrix(0, 0) = "Insurance By ID":    .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "Insurance By":       .ColWidth(1) = 8000
    .TextMatrix(0, 2) = "Status":             .ColWidth(2) = 1000
    .RowHeight(0) = 600
End With
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_InsuranceBy.EOF Then
      RsMst_InsuranceBy.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width
Frame0.Height = Me.Height - 500
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1500
Frame0.Visible = True

End Sub

Private Sub TxtInsuranceBy_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtInsuranceBy)
If ans = False Then
   MsgBox "Special character not Allowed."
   TxtInsuranceBy.SetFocus
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_InsuranceBy.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_InsuranceBy.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub



