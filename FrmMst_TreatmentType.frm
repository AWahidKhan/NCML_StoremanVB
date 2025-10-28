VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_TreatmentType 
   Caption         =   "Treatment Type"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2625
      Left            =   60
      TabIndex        =   0
      Top             =   3735
      Visible         =   0   'False
      Width           =   14895
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmMst_TreatmentType.frx":0000
         Left            =   5295
         List            =   "FrmMst_TreatmentType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   450
         Width           =   2220
      End
      Begin VB.TextBox TxtSTreatmentType 
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
         Left            =   75
         MaxLength       =   50
         TabIndex        =   13
         Top             =   450
         Width           =   5175
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
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
         Left            =   12735
         TabIndex        =   17
         Top             =   460
         Width           =   1020
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
         Left            =   13770
         TabIndex        =   16
         Top             =   460
         Width           =   1020
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   885
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Treatment Type"
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
         Left            =   1995
         TabIndex        =   20
         Top             =   165
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   6135
         TabIndex        =   19
         Top             =   165
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   60
      TabIndex        =   22
      Top             =   45
      Width           =   9360
      Begin VB.TextBox TxtTreatmentType 
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
         Left            =   2400
         MaxLength       =   200
         TabIndex        =   1
         Top             =   720
         Width           =   6645
      End
      Begin VB.TextBox TxtTreatmentTypeID 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   2715
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
         ItemData        =   "FrmMst_TreatmentType.frx":0021
         Left            =   2400
         List            =   "FrmMst_TreatmentType.frx":002B
         TabIndex        =   2
         Top             =   1200
         Width           =   2835
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
         Left            =   2400
         MaxLength       =   25
         TabIndex        =   26
         Top             =   1200
         Width           =   2835
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1680
         Width           =   6645
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2160
         Width           =   6645
      End
      Begin VB.Label LblVendorName 
         AutoSize        =   -1  'True
         Caption         =   "Treatment Type"
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
         TabIndex        =   31
         Top             =   780
         Width           =   1665
      End
      Begin VB.Label LblVendorID 
         AutoSize        =   -1  'True
         Caption         =   "Treatment Type ID"
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
         TabIndex        =   30
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label Label8 
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
         TabIndex        =   29
         Top             =   2220
         Width           =   885
      End
      Begin VB.Label Label9 
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
         TabIndex        =   28
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   27
         Top             =   1260
         Width           =   660
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   60
      TabIndex        =   21
      Top             =   2685
      Width           =   9360
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1935
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3750
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5565
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7380
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1935
         TabIndex        =   9
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3750
         TabIndex        =   10
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   5565
         TabIndex        =   11
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7380
         TabIndex        =   12
         Top             =   615
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmMst_TreatmentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As String
Dim mTreatmentTypeID As Double

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbSStatus.SetFocus
   Exit Sub
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

Private Sub CmdSearch_Click()
Dim x As Long
Dim wc As String

wc = ""
RsMst_TreatmentType.MoveFirst
Call Grid_Head

If Trim(TxtSTreatmentType) <> "" Then
   wc = " where TreatmentType like '%" & Trim(TxtSTreatmentType) & "%' "
End If

If CmbSStatus.Text <> "" Then
   If wc = "" Then
      wc = " where Status = '" & Left(CmbSStatus.Text, 1) & "'"
   Else
      wc = wc & " and Status ='" & Left(CmbSStatus.Text, 1) & "'"
   End If
End If

tmp = "Select TreatmentTypeID, TreatmentType, Status, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy  FROM Mst_TreatmentType"
tmp = tmp & wc
tmp = tmp & " order by TreatmentTypeID"
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs4.RecordCount + MSHFlexGrid1.Rows
   ProgressBar1.Max = MSHFlexGrid1.Rows
   For I = 1 To TmpRs4.RecordCount
      MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs4!TreatmentTypeID), "", TmpRs4!TreatmentTypeID)
      MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs4!TreatmentType), "", TmpRs4!TreatmentType)
      MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs4!Status), "", TmpRs4!Status)
      MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs4!CreatedBy), "", TmpRs4!CreatedBy & " : " & TmpRs4!CreatedDate)
      MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs4!UpdatedBy), "", TmpRs4!UpdatedBy & " : " & TmpRs4!UpdatedDate)
      If LCase(MSHFlexGrid1.TextMatrix(I, 2)) = "a" Then
         MSHFlexGrid1.TextMatrix(I, 2) = "Active"
      ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 2)) = "d" Then
         MSHFlexGrid1.TextMatrix(I, 2) = "De-Active"
      End If
      TmpRs4.MoveNext
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
Else
   MsgBox "Record not found!!!"
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call Grid_Head

TxtSTreatmentType.Text = ""
CmbSStatus.Text = ""

Call TableMst_TreatmentType
If RsMst_TreatmentType.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_TreatmentType.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Long
Dim mCol As Long
MsgBox ("Wait till next message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_TreatmentType.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For c = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
   Next
Next
oWB.Save
MsgBox ("Excel file created!!!")
oXL.Visible = True
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_TreatmentType.MoveNext
If RsMst_TreatmentType.EOF Then
   RsMst_TreatmentType.MoveLast
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
RsMst_TreatmentType.MovePrevious
If RsMst_TreatmentType.BOF Then
   RsMst_TreatmentType.MoveFirst
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
RsMst_TreatmentType.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsMst_TreatmentType.MoveLast
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
   If RsMst_TreatmentType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 600
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 150
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200

ProgressBar1.Width = MSHFlexGrid1.Width
ProgressBar1.Left = MSHFlexGrid1.Left
ProgressBar1.Top = MSHFlexGrid1.Top + MSHFlexGrid1.Height
End Sub

Public Sub Indata()
Frame1.Enabled = False

CmbStatus.Visible = False
TxtStatus.Visible = True

TxtTreatmentTypeID.Text = IIf(IsNull(RsMst_TreatmentType!TreatmentTypeID), "", RsMst_TreatmentType!TreatmentTypeID)
TxtTreatmentType.Text = IIf(IsNull(RsMst_TreatmentType!TreatmentType), "", RsMst_TreatmentType!TreatmentType)
TxtStatus.Text = IIf(IsNull(RsMst_TreatmentType!Status), "", RsMst_TreatmentType!Status)
TxtCreated.Text = IIf(IsNull(RsMst_TreatmentType!CreatedBy), "", RsMst_TreatmentType!CreatedBy) & " :- " & IIf(IsNull(RsMst_TreatmentType!CreatedDate), "", RsMst_TreatmentType!CreatedDate)
TxtUpdated.Text = IIf(IsNull(RsMst_TreatmentType!UpdatedBy), "", RsMst_TreatmentType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_TreatmentType!UpdatedDate), "", RsMst_TreatmentType!UpdatedDate)
If LCase(TxtStatus.Text) = "a" Then
   TxtStatus.Text = "Active"
ElseIf LCase(TxtStatus.Text) = "d" Then
   TxtStatus.Text = "De-Active"
End If
CmbStatus.Text = TxtStatus.Text

Call GButtonLock(Me, True)
End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmbStatus.Visible = True
TxtStatus.Visible = False

TxtTreatmentTypeID.Text = ""
TxtTreatmentType.Text = ""
TxtStatus.Text = ""
TxtCreated.Text = ""
TxtUpdated.Text = ""
CmbStatus.Text = ""
If RsMst_TreatmentType.RecordCount > 0 Then
   TxtTreatmentType.SetFocus
End If

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmbStatus.Visible = True
TxtStatus.Visible = False

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_TreatmentType.Delete
   RsMst_TreatmentType.Update
   If RsMst_TreatmentType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_TreatmentType.MoveNext
   If RsMst_TreatmentType.EOF() Then
      RsMst_TreatmentType.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
   MsgBox "Child record Present "
   RsMst_TreatmentType.CancelUpdate
End Sub

Private Sub SaveCmd_Click()
If TxtTreatmentType.Text = "" Then
   MsgBox "Blank Treatment Type Not Allowed!!!"
   TxtTreatmentType.SetFocus
   Exit Sub
End If
If CmbStatus.Text = "" Then
   MsgBox "Blank Status Not Allowed!!!"
   CmbStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_TreatmentType.RecordCount > 0 Then
      RsMst_TreatmentType.MoveFirst
      RsMst_TreatmentType.Find "TreatmentType= '" & Trim(TxtTreatmentType.Text) & "'"
      If Not RsMst_TreatmentType.EOF Then
         MsgBox "Duplicate Treatment Type Not Allowed!!!"
         TxtTreatmentType.SetFocus
         Exit Sub
      End If
   End If
   RsMst_TreatmentType.AddNew
   RsMst_TreatmentType!TreatmentTypeID = GetTreatmentTypeID
   RsMst_TreatmentType!G_UID = GetGUID
   TxtTreatmentTypeID.Text = RsMst_TreatmentType!TreatmentTypeID
Else
   RsMst_TreatmentType.MoveFirst
   RsMst_TreatmentType.Find "TreatmentType= '" & Trim(TxtTreatmentType.Text) & "'"
   If Not RsMst_TreatmentType.EOF Then
      If RsMst_TreatmentType!TreatmentTypeID <> TxtTreatmentTypeID.Text Then
         MsgBox "Duplicate Treatment Type Not Allowed!!!"
         TxtTreatmentType.SetFocus
         Exit Sub
      End If
   End If
   RsMst_TreatmentType.MoveFirst
   RsMst_TreatmentType.Find " TreatmentTypeID= " & TxtTreatmentTypeID.Text
End If
RsMst_TreatmentType!TreatmentType = TxtTreatmentType.Text
RsMst_TreatmentType!Status = Left$(CmbStatus.Text, 1)
If IsNull(RsMst_TreatmentType!CreatedBy) = True Or RsMst_TreatmentType!CreatedBy = "" Then
   RsMst_TreatmentType!CreatedBy = Gen_UserLoginID
   RsMst_TreatmentType!CreatedDate = Now
Else
   RsMst_TreatmentType!UpdatedBy = Gen_UserLoginID
   RsMst_TreatmentType!UpdatedDate = Now
End If
RsMst_TreatmentType.Update

Call Indata
End Sub

Private Function GetTreatmentTypeID() As Double
Dim Retval As Double
tmp = "Select max(TreatmentTypeID) from Mst_TreatmentType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetTreatmentTypeID = Retval
End Function

Private Sub Grid_Head()
ProgressBar1.Value = 0
CmdExcel.Enabled = False

With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 5
   
   .TextMatrix(0, 0) = "Treatment TypeID"
   .TextMatrix(0, 1) = "Treatment Type"
   .TextMatrix(0, 2) = "Status"
   .TextMatrix(0, 3) = "CreatedBy"
   .TextMatrix(0, 4) = "UpdatedBy"
   
   .ColWidth(0) = 1200
   .ColWidth(1) = 3900
   .ColWidth(2) = 1500
   .ColWidth(3) = 2200
   .ColWidth(4) = 2200
   
   .RowHeight(0) = 600
   
End With
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_TreatmentType.MoveFirst
RsMst_TreatmentType.Find "TreatmentTypeID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_TreatmentType.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

