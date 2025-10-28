VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_Mst_LicenseType 
   Caption         =   "License Type Master"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2925
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   14880
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   8775
         TabIndex        =   26
         Top             =   492
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchState 
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
         TabIndex        =   25
         Top             =   480
         Width           =   7605
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   7770
         TabIndex        =   24
         Top             =   495
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   27
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "License Type Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2842
         TabIndex        =   30
         Top             =   210
         Width           =   2160
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
         Left            =   9885
         TabIndex        =   29
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
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
         Left            =   12285
         TabIndex        =   28
         Top             =   540
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   7200
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   11
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   6
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   10
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   5
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   405
         Left            =   2895
         TabIndex        =   9
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   4
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   8
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   7
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   3
         Top             =   240
         Width           =   1400
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   8400
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
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1530
         Width           =   4650
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
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1995
         Width           =   4635
      End
      Begin VB.TextBox TxtLicenseID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox TxtLicenseType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2325
         MaxLength       =   100
         TabIndex        =   1
         Top             =   684
         Width           =   4635
      End
      Begin VB.ComboBox CmbFlag 
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
         ItemData        =   "Frm_Mst_LicenseType.frx":0000
         Left            =   2325
         List            =   "Frm_Mst_LicenseType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1095
         Width           =   4635
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2325
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   16
         Top             =   1095
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   1125
         Width           =   480
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
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2055
         Width           =   975
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
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1605
         Width           =   1695
      End
      Begin VB.Label LblStateID 
         Caption         =   "License Type ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   308
         Width           =   1800
      End
      Begin VB.Label LblStateName 
         Caption         =   "License Type Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   750
         Width           =   2280
      End
   End
End
Attribute VB_Name = "Frm_Mst_LicenseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mStateIndiaID As Variant

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub

If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmdGo_Click()

tmp = "Select MST.LicenseTypeID,MST.LicenseTypeName,MST.Flag from Mst_LicenseType MST"

Dim wc As Variant
wc = ""
If TxtSearchState <> "" Then
   If wc = "" Then
      wc = " Where MST.LicenseTypeName Like '%" & TxtSearchState & "%' "
   Else
      wc = wc & " and MST.LicenseTypeName Like '%" & TxtSearchState & "%' "
   End If
End If
tmp = tmp & wc & " Order By MST.LicenseTypeID,MST.LicenseTypeName "

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    MSHFlexGrid1.TextMatrix(iu, 2) = IIf(MSHFlexGrid1.TextMatrix(iu, 2) = "A", "Active", "DeActive")
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
TxtSearchState = ""
Label22.Caption = ""

Call TableMst_LicenseType
If RsMst_LicenseType.RecordCount = 0 Then
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
Call InData
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_LicenseType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_LicenseType.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
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
RsMst_LicenseType.MoveNext
If RsMst_LicenseType.EOF Then
   RsMst_LicenseType.MoveLast
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
RsMst_LicenseType.MovePrevious
If RsMst_LicenseType.BOF Then
   RsMst_LicenseType.MoveFirst
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
Private Sub FirstCmd_Click()
RsMst_LicenseType.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_LicenseType.MoveLast
Call InData
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
   If RsMst_LicenseType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_LicenseType.EOF Then
      RsMst_LicenseType.MoveFirst
   End If
   Call InData
   Exit Sub
End If

Call Grid_Head
TxtSearchState.Text = ""

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 600
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_LicenseType.MoveFirst

End Sub
Public Sub InData()

TxtLicenseID.Locked = True
TxtLicenseType.Locked = True
CmbFlag.Visible = False
TxtFlag.Visible = True

TxtLicenseID = IIf(IsNull(RsMst_LicenseType!LicenseTypeID), "", RsMst_LicenseType!LicenseTypeID)
TxtLicenseType = IIf(IsNull(RsMst_LicenseType!LicenseTypeName), "", RsMst_LicenseType!LicenseTypeName)

TxtCreated = IIf(IsNull(RsMst_LicenseType!CreatedBy), "", RsMst_LicenseType!CreatedBy) & " :- " & IIf(IsNull(RsMst_LicenseType!CreatedDate), "", RsMst_LicenseType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_LicenseType!UpdatedBy), "", RsMst_LicenseType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_LicenseType!UpdatedDate), "", RsMst_LicenseType!UpdatedDate)

If LCase(RsMst_LicenseType!Flag) = "a" Then
   TxtFlag.Text = "Active"
ElseIf LCase(RsMst_LicenseType!Flag) = "d" Then
   TxtFlag.Text = "De-Active"
Else
   TxtFlag.Text = ""
End If

If RsMst_LicenseType.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtFlag")
End If
End Sub

Private Sub TxtLicenseID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtLicenseID")
End If
End Sub

Private Sub TxtLicenseID_LostFocus()
If TxtLicenseID = "" Then Exit Sub
If IsNumeric(TxtLicenseID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtLicenseID.SetFocus
   Exit Sub
End If
TxtLicenseID = Val(TxtLicenseID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtLicenseID.Locked = True
TxtLicenseType.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True

TxtLicenseID = ""
TxtLicenseType = ""

CmbFlag.Text = ""

TxtCreated = ""
TxtUpdated = ""
If RsMst_LicenseType.RecordCount > 0 Then
   TxtLicenseType.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtLicenseID.Locked = True
TxtLicenseType.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = TxtFlag.Text
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If Ans = vbYes Then
   RsMst_LicenseType.Delete
   RsMst_LicenseType.Update
   If RsMst_LicenseType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_LicenseType.MoveNext
   If RsMst_LicenseType.EOF() Then
      RsMst_LicenseType.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_LicenseType.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If TxtLicenseType.Text = "" Then
   MsgBox "Blank LicenseTypeName Not Allowed !!! "
   TxtLicenseType.SetFocus
   Exit Sub
End If
If CmbFlag.Text = "" Then
   MsgBox "Blank Flag Not Allowed !!! "
   CmbFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_LicenseType.RecordCount > 0 Then
      RsMst_LicenseType.MoveFirst
      RsMst_LicenseType.Find "LicenseTypeName= '" & TxtLicenseType.Text & "'"
      If Not RsMst_LicenseType.EOF Then
        MsgBox "Duplicate License Type Name Not Allowed !!! "
         TxtLicenseType.SetFocus
         Exit Sub
      End If
   End If
      RsMst_LicenseType.AddNew
      RsMst_LicenseType!LicenseTypeID = GetLicenseTypeID
      RsMst_LicenseType!G_UID = GetGUID
      TxtLicenseID.Text = RsMst_LicenseType!LicenseTypeID
Else
    RsMst_LicenseType.MoveFirst
    RsMst_LicenseType.Find "LicenseTypeName= '" & TxtLicenseType.Text & "'"
    If Not RsMst_LicenseType.EOF Then
       If RsMst_LicenseType!LicenseTypeID <> TxtLicenseID.Text Then
          MsgBox "Duplicate License Type Name Not Allowed !!! "
          TxtLicenseType.SetFocus
          Exit Sub
       End If
    End If
    RsMst_LicenseType.MoveFirst
    RsMst_LicenseType.Find "LicenseTypeID= " & TxtLicenseID.Text
End If

RsMst_LicenseType!LicenseTypeID = TxtLicenseID.Text
RsMst_LicenseType!LicenseTypeName = TxtLicenseType.Text
RsMst_LicenseType!Flag = Left(CmbFlag.Text, 1)

If IsNull(RsMst_LicenseType!CreatedBy) = True Or RsMst_LicenseType!CreatedBy = "" Then
   RsMst_LicenseType!CreatedBy = Gen_UserLoginID
   RsMst_LicenseType!CreatedDate = Now
Else
   RsMst_LicenseType!UpdatedBy = Gen_UserLoginID
   RsMst_LicenseType!UpdatedDate = Now
End If

RsMst_LicenseType.Update
RsMst_LicenseType.Requery

RsMst_LicenseType.MoveFirst
RsMst_LicenseType.Find "LicenseTypeID = " & TxtLicenseID

Call InData

End Sub
Private Function GetLicenseTypeID() As Double
Dim Retval As Double
tmp = "Select max(LicenseTypeID) from Mst_LicenseType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetLicenseTypeID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0, 0) = "LicenseID"
MSHFlexGrid1.TextMatrix(0, 1) = "License Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 1500

End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_LicenseType.MoveFirst
RsMst_LicenseType.Find "LicenseTypeID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_LicenseType.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub

Private Sub TxtLicenseType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtLicenseType")
End If
End Sub



