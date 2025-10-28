VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_TestingAgency 
   Caption         =   "Testing Agency"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      Top             =   4035
      Visible         =   0   'False
      Width           =   13335
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
         Left            =   7965
         TabIndex        =   15
         Top             =   240
         Width           =   1000
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
         Height          =   350
         Left            =   6960
         TabIndex        =   14
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox TxtSAgencyName 
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
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   16
         Top             =   795
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         Caption         =   "Agency Name"
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
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1455
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
         Left            =   12120
         TabIndex        =   18
         Top             =   300
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
         Left            =   9600
         TabIndex        =   17
         Top             =   288
         Width           =   2370
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1140
      Left            =   120
      TabIndex        =   30
      Top             =   2760
      Width           =   7410
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
         Left            =   195
         TabIndex        =   0
         Top             =   240
         Width           =   1395
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
         Left            =   1590
         TabIndex        =   3
         Top             =   240
         Width           =   1395
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
         Left            =   2985
         TabIndex        =   4
         Top             =   240
         Width           =   1395
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
         Left            =   4380
         TabIndex        =   5
         Top             =   240
         Width           =   1395
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
         Left            =   5775
         TabIndex        =   6
         Top             =   240
         Width           =   1395
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
         Left            =   195
         TabIndex        =   11
         Top             =   615
         Width           =   1395
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
         Left            =   1590
         TabIndex        =   10
         Top             =   615
         Width           =   1395
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
         Left            =   2985
         TabIndex        =   9
         Top             =   615
         Width           =   1395
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
         Left            =   4380
         TabIndex        =   8
         Top             =   615
         Width           =   1395
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
         Left            =   5775
         TabIndex        =   7
         Top             =   615
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   7395
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5295
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5295
      End
      Begin VB.TextBox TxtAgencyID 
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtAgencyName 
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
         TabIndex        =   1
         Top             =   720
         Width           =   5295
      End
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
         ItemData        =   "FrmMst_TestingAgency.frx":0000
         Left            =   1905
         List            =   "FrmMst_TestingAgency.frx":000A
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
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
         Height          =   360
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Agency ID"
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
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Agency Name"
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
         TabIndex        =   26
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label Label7 
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
         TabIndex        =   25
         Top             =   1260
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_TestingAgency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtAgencyName.Locked = False
TxtAgencyName = ""
TxtAgencyID = ""
CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = ""

TxtCreated = ""
TxtUpdated = ""

If RsMst_TestingAgency.RecordCount > 0 Then
   TxtAgencyName.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "deactive" Then
   MsgBox "Invalid Selection."
   CmbStatus.SetFocus
End If
End Sub

Private Sub CmdGo_Click()
Call Grid_Head

Label22.Caption = ""

tmp = "Select AgencyID,AgencyName,Status from Mst_TestingAgency"

If TxtSAgencyName.Text <> "" Then
   tmp = tmp & "  Where AgencyName Like '%" & TxtSAgencyName.Text & "%'"
End If

tmp = tmp & " Order By AgencyName"
Call TmpTable
For I = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!TestingAgencyID
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!AgencyName
    If LCase(TmpRs!Status) = "a" Then
       MSHFlexGrid1.TextMatrix(I, 2) = "Active"
    ElseIf LCase(TmpRs!Status) = "d" Then
       MSHFlexGrid1.TextMatrix(I, 2) = "Deactive"
    End If
    TmpRs.MoveNext
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   
   RsMst_TestingAgency.Delete
   RsMst_TestingAgency.Update
   If RsMst_TestingAgency.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_TestingAgency.MoveNext
   If RsMst_TestingAgency.EOF() Then
      RsMst_TestingAgency.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_TestingAgency.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtAgencyName.Locked = False
CmbStatus.Visible = True
TxtStatus.Visible = False
CmbStatus.Text = TxtStatus.Text

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_TestingAgency.MoveFirst
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
   Frame1.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
TxtSAgencyName.Text = ""
Label22.Caption = ""
Call Grid_Head
Frame1.Enabled = True
Call TableMst_TestingAgency(" Order by AgencyID ")

If RsMst_TestingAgency.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtAgencyName.Locked = True
TxtStatus.Locked = True

CmbStatus.Visible = False
TxtStatus.Visible = True
TxtStatus.Text = ""

TxtAgencyID = RsMst_TestingAgency!TestingAgencyID
TxtAgencyName = RsMst_TestingAgency!AgencyName
If LCase(RsMst_TestingAgency!Status) = "a" Then
   TxtStatus.Text = "Active"
ElseIf LCase(RsMst_TestingAgency!Status) = "d" Then
   TxtStatus.Text = "Deactive"
Else
   TxtStatus.Text = ""
End If
TxtCreated = IIf(IsNull(RsMst_TestingAgency!CreatedBy), "", RsMst_TestingAgency!CreatedBy) & " :- " & IIf(IsNull(RsMst_TestingAgency!CreatedDate), "", RsMst_TestingAgency!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_TestingAgency!UpdatedBy), "", RsMst_TestingAgency!UpdatedBy) & " :- " & IIf(IsNull(RsMst_TestingAgency!UpdatedDate), "", RsMst_TestingAgency!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_TestingAgency.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
If mCaId = "" Then Exit Sub
RsMst_TestingAgency.MoveFirst

If RsMst_TestingAgency.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_TestingAgency.MoveFirst
RsMst_TestingAgency.Find "AgencyID = " & Val(mCaId) & ""
If RsMst_TestingAgency.EOF Then
   RsMst_TestingAgency.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_TestingAgency.MoveNext
If RsMst_TestingAgency.EOF Then
   RsMst_TestingAgency.MoveLast
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
RsMst_TestingAgency.MovePrevious
If RsMst_TestingAgency.BOF Then
   RsMst_TestingAgency.MoveFirst
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

If TxtAgencyName = "" Then
   MsgBox "Blank Agency Name Not Allowed"
   TxtAgencyName.SetFocus
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Blank Status Not Allowed !!"
   CmbStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_TestingAgency.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_TestingAgency.MoveFirst
        RsMst_TestingAgency.Find "AgencyName = '" & TxtAgencyName & "'"
        If Not RsMst_TestingAgency.EOF Then
           MsgBox "Duplicate Testing Agency not Allowed"
           Exit Sub
        End If
        RsMst_TestingAgency.MoveLast
        
        tmp = "select max(AgencyID) from Mst_TestingAgency"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_TestingAgency.AddNew
    RsMst_TestingAgency!TestingAgencyID = tmp
    TxtAgencyID = tmp
Else
    RsMst_TestingAgency.MoveFirst
    RsMst_TestingAgency.Find "AgencyName = '" & TxtAgencyName & "'"
    If Not RsMst_TestingAgency.EOF Then
       If RsMst_TestingAgency!TestingAgencyID <> TxtAgencyID Then
          MsgBox "Duplicate Agency Name not Allowed"
          Exit Sub
       End If
    End If
   RsMst_TestingAgency.MoveFirst
   RsMst_TestingAgency.Find "AgencyID = " & TxtAgencyID & ""
End If
RsMst_TestingAgency!AgencyName = TxtAgencyName
RsMst_TestingAgency!Status = UCase(Left(CmbStatus.Text, 1))

If IsNull(RsMst_TestingAgency!CreatedBy) = True Or RsMst_TestingAgency!CreatedBy = "" Then
   RsMst_TestingAgency!CreatedBy = Gen_UserLoginID
   RsMst_TestingAgency!CreatedDate = Now
Else
   RsMst_TestingAgency!UpdatedBy = Gen_UserLoginID
   RsMst_TestingAgency!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_TestingAgency.Update

Call GButtonLock(Me, True)
TxtAgencyName.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_TestingAgency.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_TestingAgency.EOF Then
      RsMst_TestingAgency.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 600
Frame0.Width = Me.Width - 300
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top

Frame0.Visible = True
RsMst_TestingAgency.MoveFirst

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 900
''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'TxtSAgencyName.Text = ""
'Label22.Caption = ""
'Call Grid_Head
''===========comment on 7 july 2012 for not clearing search grid=============
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Font.Size = 10
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 500
MSHFlexGrid1.TextMatrix(0, 0) = "Agency ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Agency Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Status"
MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 9500
MSHFlexGrid1.ColWidth(2) = 2300
End Sub
Private Sub TxtAgencyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAgencyName")
End If
End Sub

Private Sub TxtAgencyName_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAgencyName)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAgencyName.SetFocus
End If
End Sub



Private Sub TxtAgencyID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAgencyID")
End If
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Mst_TestingAgency.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Mst_TestingAgency.xls")
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


