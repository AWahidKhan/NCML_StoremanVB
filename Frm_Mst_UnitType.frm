VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_UnitType 
   Caption         =   "Unit Type Master"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   10575
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   15
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   14
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
      Height          =   2625
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7245
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
         ItemData        =   "Frm_Mst_UnitType.frx":0000
         Left            =   4755
         List            =   "Frm_Mst_UnitType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1147
         Width           =   2325
      End
      Begin VB.TextBox TxtUnitTypeID 
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1320
      End
      Begin VB.TextBox TxtUnitType 
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
         MaxLength       =   75
         TabIndex        =   1
         Top             =   690
         Width           =   5160
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
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1590
         Width           =   5175
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
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2040
         Width           =   5175
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
         Left            =   4755
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1147
         Width           =   2325
      End
      Begin VB.ComboBox CmbFixVrb 
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
         ItemData        =   "Frm_Mst_UnitType.frx":0021
         Left            =   1920
         List            =   "Frm_Mst_UnitType.frx":002B
         TabIndex        =   2
         Top             =   1147
         Width           =   1935
      End
      Begin VB.TextBox TxtFixVrb 
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
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1920
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   26
         Top             =   2107
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   1657
         Width           =   855
      End
      Begin VB.Label LblUnitTypeID 
         AutoSize        =   -1  'True
         Caption         =   "Unit Type ID"
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
         TabIndex        =   17
         Top             =   307
         Width           =   1110
      End
      Begin VB.Label LblUnitType 
         AutoSize        =   -1  'True
         Caption         =   "Unit Type"
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
         TabIndex        =   19
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fixed-Variable"
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
         TabIndex        =   22
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3960
         TabIndex        =   28
         Top             =   1207
         Width           =   660
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1320
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   7245
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   13
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   8
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   12
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   7
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2895
         TabIndex        =   11
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   6
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   10
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   5
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   9
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   4
         Top             =   240
         Width           =   1400
      End
   End
End
Attribute VB_Name = "FrmMst_UnitType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub CmbFixVrb_LostFocus()
If CmbFixVrb.Text <> "Fixed" And CmbFixVrb.Text <> "Variable" Then
   MsgBox "Enter Fixed or Variable !!!"
   CmbFixVrb.SetFocus
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_UnitType
If RsMst_UnitType.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_UnitType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_UnitType.xls")
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
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_UnitType.MoveNext
If RsMst_UnitType.EOF Then
   RsMst_UnitType.MoveLast
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
RsMst_UnitType.MovePrevious
If RsMst_UnitType.BOF Then
   RsMst_UnitType.MoveFirst
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
Private Sub FirstCmd_Click()
RsMst_UnitType.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_UnitType.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_UnitType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 1000
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_UnitType.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_UnitType.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For I = 1 To RsMst_UnitType.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(RsMst_UnitType!UnitTypeID), "", RsMst_UnitType!UnitTypeID)
    MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(RsMst_UnitType!UnitType), "", RsMst_UnitType!UnitType)
    MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(RsMst_UnitType!Flag), "", RsMst_UnitType!Flag)
    MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(RsMst_UnitType!Status), "", RsMst_UnitType!Status)
    If LCase(MSHFlexGrid1.TextMatrix(I, 3)) = "a" Then
       MSHFlexGrid1.TextMatrix(I, 3) = "Active"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 3)) = "d" Then
       MSHFlexGrid1.TextMatrix(I, 3) = "De-Active"
    End If
    RsMst_UnitType.MoveNext
Next
End Sub
Public Sub Indata()

TxtUnitTypeID.Locked = True
TxtUnitType.Locked = True
CmbFixVrb.Visible = False
TxtFixVrb.Visible = True
CmbStatus.Visible = False
TxtStatus.Visible = True

TxtUnitTypeID = IIf(IsNull(RsMst_UnitType!UnitTypeID), "", RsMst_UnitType!UnitTypeID)
TxtUnitType = IIf(IsNull(RsMst_UnitType!UnitType), "", RsMst_UnitType!UnitType)
TxtFixVrb = IIf(IsNull(RsMst_UnitType!Flag), "", RsMst_UnitType!Flag)
TxtStatus = IIf(IsNull(RsMst_UnitType!Status), "A", RsMst_UnitType!Status)

If LCase(TxtStatus.Text) = "a" Then
   TxtStatus.Text = "Active"
ElseIf LCase(TxtStatus.Text) = "d" Then
   TxtStatus.Text = "De-Active"
End If

TxtCreated = IIf(IsNull(RsMst_UnitType!CreatedBy), "", RsMst_UnitType!CreatedBy) & " :- " & IIf(IsNull(RsMst_UnitType!CreatedDate), "", RsMst_UnitType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_UnitType!UpdatedBy), "", RsMst_UnitType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_UnitType!UpdatedDate), "", RsMst_UnitType!UpdatedDate)

CmbFixVrb.Text = TxtFixVrb.Text
CmbStatus.Text = TxtStatus.Text
If RsMst_UnitType.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtUnitType_LostFocus()
TxtUnitType = Trim(TxtUnitType)
End Sub

Private Sub TxtUnitTypeID_LostFocus()
If TxtUnitTypeID = "" Then Exit Sub
If IsNumeric(TxtUnitTypeID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtUnitTypeID.SetFocus
   Exit Sub
End If
TxtUnitTypeID = Val(TxtUnitTypeID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtUnitTypeID.Locked = True
TxtUnitType.Locked = False
CmbFixVrb.Visible = True
TxtFixVrb.Visible = False
CmbStatus.Visible = True
TxtStatus.Visible = False

TxtUnitTypeID = ""
TxtUnitType = ""
CmbStatus.Text = ""
TxtStatus.Text = ""
CmbFixVrb.Text = ""
TxtFixVrb.Text = ""
TxtCreated = ""
TxtUpdated = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtUnitTypeID.Locked = True
TxtUnitType.Locked = False
CmbFixVrb.Visible = True
TxtFixVrb.Visible = False
CmbStatus.Visible = True
TxtStatus.Visible = False

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_UnitType.Delete
   RsMst_UnitType.Update
   If RsMst_UnitType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_UnitType.MoveNext
   If RsMst_UnitType.EOF() Then
      RsMst_UnitType.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_UnitType.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If TxtUnitType.Text = "" Then
   MsgBox "Blank Unit Type Not Allowed!!!"
   TxtUnitType.SetFocus
   Exit Sub
End If
If CmbFixVrb.Text = "" Then
   MsgBox "Blank Fixed-Variable Not Allowed!!!"
   CmbFixVrb.SetFocus
   Exit Sub
End If
If CmbStatus.Text = "" Then
   MsgBox "Blank Status Not Allowed!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_UnitType.RecordCount > 0 Then
      RsMst_UnitType.MoveFirst
      RsMst_UnitType.Find "UnitType = '" & TxtUnitType.Text & "'"
      If Not RsMst_UnitType.EOF Then
        MsgBox "Duplicate Unit Name Not Allowed !!! "
         TxtUnitType.SetFocus
         Exit Sub
      End If
   End If
      RsMst_UnitType.AddNew
      RsMst_UnitType!UnitTypeID = GetUnitTypeID
      RsMst_UnitType!G_UID = GetGUID
      TxtUnitTypeID.Text = RsMst_UnitType!UnitTypeID
Else
    RsMst_UnitType.MoveFirst
    RsMst_UnitType.Find "UnitType = '" & TxtUnitType.Text & "'"
    If Not RsMst_UnitType.EOF Then
       If RsMst_UnitType!UnitTypeID <> TxtUnitTypeID.Text Then
          MsgBox "Duplicate Unit Name Not Allowed !!! "
          Exit Sub
       End If
    End If
    RsMst_UnitType.MoveFirst
    RsMst_UnitType.Find " UnitTypeID= " & TxtUnitTypeID.Text
End If

RsMst_UnitType!UnitTypeID = TxtUnitTypeID.Text
RsMst_UnitType!UnitType = TxtUnitType.Text
RsMst_UnitType!Flag = CmbFixVrb.Text
RsMst_UnitType!Status = Left$(CmbStatus.Text, 1)
If IsNull(RsMst_UnitType!CreatedBy) = True Or RsMst_UnitType!CreatedBy = "" Then
   RsMst_UnitType!CreatedBy = Gen_UserLoginID
   RsMst_UnitType!CreatedDate = Now
Else
   RsMst_UnitType!UpdatedBy = Gen_UserLoginID
   RsMst_UnitType!UpdatedDate = Now
End If
RsMst_UnitType.Update
Call Indata
End Sub
Private Function GetUnitTypeID() As Double
Dim Retval As Double
tmp = "Select max(UnitTypeID) from Mst_UnitType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetUnitTypeID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4
MSHFlexGrid1.TextMatrix(0, 0) = "UnitType ID"
MSHFlexGrid1.TextMatrix(0, 1) = "UnitType"
MSHFlexGrid1.TextMatrix(0, 2) = "Fixed-Variable"
MSHFlexGrid1.TextMatrix(0, 3) = "Status"

MSHFlexGrid1.ColWidth(1) = 6000
MSHFlexGrid1.ColWidth(2) = 2000
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitTypeID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_UnitType.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
