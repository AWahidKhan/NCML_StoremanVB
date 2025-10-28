VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_WHRUOM 
   Caption         =   "WHR UOM"
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
      Height          =   885
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   1095
         TabIndex        =   14
         Top             =   165
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   90
         TabIndex        =   13
         Top             =   165
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   15
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
         Left            =   9405
         TabIndex        =   17
         Top             =   210
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
         Left            =   11805
         TabIndex        =   16
         Top             =   210
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   6600
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   255
         TabIndex        =   3
         Top             =   225
         Width           =   1200
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   255
         TabIndex        =   8
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1470
         TabIndex        =   4
         Top             =   225
         Width           =   1200
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1470
         TabIndex        =   9
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2670
         TabIndex        =   5
         Top             =   225
         Width           =   1200
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   405
         Left            =   2670
         TabIndex        =   10
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   3870
         TabIndex        =   6
         Top             =   225
         Width           =   1200
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   3870
         TabIndex        =   11
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5070
         TabIndex        =   7
         Top             =   225
         Width           =   1200
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5070
         TabIndex        =   12
         Top             =   630
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   6600
      Begin VB.ComboBox CmbApplicableOn 
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
         ItemData        =   "FrmMst_WHRUOM.frx":0000
         Left            =   1560
         List            =   "FrmMst_WHRUOM.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1162
         Width           =   4635
      End
      Begin VB.TextBox TxtDesc 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   701
         Width           =   4635
      End
      Begin VB.TextBox TxtWhrUnitID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1680
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2085
         Width           =   4635
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1623
         Width           =   4635
      End
      Begin VB.TextBox TxtApplicableOn 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.Label Label1 
         Caption         =   "Applicable On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   90
         TabIndex        =   27
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label LblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Descripntion"
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
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label LblWHRUnitID 
         Caption         =   "WHR UnitID"
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
         TabIndex        =   24
         Top             =   315
         Width           =   1800
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
         TabIndex        =   23
         Top             =   2145
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
         TabIndex        =   22
         Top             =   1740
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_WHRUOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub CmbApplicableOn_LostFocus()
If CmbApplicableOn.Text = "" Then
   Exit Sub
End If
If CmbApplicableOn.Text <> "Quantity" And CmbApplicableOn.Text <> "Bags" And CmbApplicableOn.Text <> "Direct Amount" And CmbApplicableOn.Text <> "Other" Then
   MsgBox "Invalid selection "
   CmbApplicableOn.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdGo_Click()
tmp = "Select WHRUnitID,Description,ApplicableOn from Mst_WHRUOM "

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
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

Call TableMst_WHRUOM
If RsMst_WHRUOM.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_WHRUOM.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_WHRUOM.xls")
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
RsMst_WHRUOM.MoveNext
If RsMst_WHRUOM.EOF Then
   RsMst_WHRUOM.MoveLast
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
RsMst_WHRUOM.MovePrevious
If RsMst_WHRUOM.BOF Then
   RsMst_WHRUOM.MoveFirst
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
RsMst_WHRUOM.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_WHRUOM.MoveLast
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
   If RsMst_WHRUOM.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 600
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_WHRUOM.MoveFirst
Call Grid_Head
Label22.Caption = ""
'MSHFlexGrid1.Rows = RsMst_WHRUOM.RecordCount + 1
''MSHFlexGrid1.ColWidth(0) = 1000
'For i = 1 To RsMst_WHRUOM.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_WHRUOM.Fields(c)), "", RsMst_WHRUOM.Fields(c))
'    Next
'    RsMst_WHRUOM.MoveNext
'Next
End Sub
Public Sub Indata()

TxtWhrUnitID.Locked = True
TxtDesc.Locked = True
CmbApplicableOn.Visible = False
TxtApplicableOn.Visible = True
TxtApplicableOn.Locked = True

TxtWhrUnitID = IIf(IsNull(RsMst_WHRUOM!WHRUnitID), "", RsMst_WHRUOM!WHRUnitID)
TxtDesc = IIf(IsNull(RsMst_WHRUOM!Description), "", RsMst_WHRUOM!Description)
TxtApplicableOn = IIf(IsNull(RsMst_WHRUOM!ApplicableOn), "", RsMst_WHRUOM!ApplicableOn)

TxtCreated = IIf(IsNull(RsMst_WHRUOM!CreatedBy), "", RsMst_WHRUOM!CreatedBy) & " :- " & IIf(IsNull(RsMst_WHRUOM!CreatedDate), "", RsMst_WHRUOM!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_WHRUOM!UpdatedBy), "", RsMst_WHRUOM!UpdatedBy) & " :- " & IIf(IsNull(RsMst_WHRUOM!UpdatedDate), "", RsMst_WHRUOM!UpdatedDate)


If RsMst_WHRUOM.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtWhrUnitID.Locked = True
TxtDesc.Locked = False
TxtApplicableOn.Visible = False
CmbApplicableOn.Visible = True
TxtWhrUnitID = ""
TxtDesc = ""
CmbApplicableOn.Text = ""
TxtCreated = ""
TxtUpdated = ""
TxtFlag = ""
If RsMst_WHRUOM.RecordCount > 0 Then
   TxtDesc.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtWhrUnitID.Locked = True
TxtDesc.Locked = False
TxtApplicableOn.Visible = False
CmbApplicableOn.Visible = True
CmbApplicableOn.Text = TxtApplicableOn.Text
End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_WHRUOM.Delete
   RsMst_WHRUOM.Update
   If RsMst_WHRUOM.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_WHRUOM.MoveNext
   If RsMst_WHRUOM.EOF() Then
      RsMst_WHRUOM.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_WHRUOM.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If TxtDesc.Text = "" Then
   MsgBox "Blank Description Not Allowed !!! "
   TxtDesc.SetFocus
   Exit Sub
End If

If CmbApplicableOn.Text = "" Then
   MsgBox "Select Applicable On!!! "
   CmbApplicableOn.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_WHRUOM.RecordCount > 0 Then
      RsMst_WHRUOM.MoveFirst
      'RsMst_WHRUOM.Find "ApplicableOn= '" & CmbApplicableOn.Text & "'"
      RsMst_WHRUOM.Find "Description= '" & TxtDesc & "'"
      If Not RsMst_WHRUOM.EOF Then
        MsgBox "Duplicate Description Not Allowed !!! "
        TxtDesc.SetFocus
        Exit Sub
      End If
   End If
      RsMst_WHRUOM.AddNew
      RsMst_WHRUOM!WHRUnitID = GetWHRUnitID
      RsMst_WHRUOM!G_UID = GetGUID
      TxtWhrUnitID.Text = RsMst_WHRUOM!WHRUnitID
Else
    RsMst_WHRUOM.MoveFirst
    'RsMst_WHRUOM.Find "ApplicableOn= '" & CmbApplicableOn.Text & "'"
    RsMst_WHRUOM.Find "Description= '" & TxtDesc & "'"
    If Not RsMst_WHRUOM.EOF Then
       If RsMst_WHRUOM!WHRUnitID <> TxtWhrUnitID.Text Then
          MsgBox "Duplicate Description Not Allowed !!! "
          TxtDesc.SetFocus
          Exit Sub
       End If
    End If
    RsMst_WHRUOM.MoveFirst
    RsMst_WHRUOM.Find " WHRUnitID= " & TxtWhrUnitID.Text
End If

RsMst_WHRUOM!WHRUnitID = TxtWhrUnitID.Text
RsMst_WHRUOM!Description = TxtDesc.Text
RsMst_WHRUOM!ApplicableOn = CmbApplicableOn.Text

If IsNull(RsMst_WHRUOM!CreatedBy) = True Or RsMst_WHRUOM!CreatedBy = "" Then
   RsMst_WHRUOM!CreatedBy = Gen_UserLoginID
   RsMst_WHRUOM!CreatedDate = Now
Else
   RsMst_WHRUOM!UpdatedBy = Gen_UserLoginID
   RsMst_WHRUOM!UpdatedDate = Now
End If

RsMst_WHRUOM.Update
Call Indata
End Sub
Private Function GetWHRUnitID() As Double
Dim Retval As Double
tmp = "Select max(WHRUnitID) from Mst_WHRUOM"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetWHRUnitID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0, 0) = "WHR UnitID"
MSHFlexGrid1.TextMatrix(0, 1) = "Description"
MSHFlexGrid1.TextMatrix(0, 2) = "Applicable On"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 1800

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_WHRUOM.MoveFirst
RsMst_WHRUOM.Find "WHRUnitID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_WHRUOM.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtDesc")
End If
End Sub

