VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMstInsuranceCo 
   Caption         =   "Insurance Co"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   7095
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   4230
      Left            =   135
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   14925
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmMstInsuranceCo.frx":0000
         Left            =   6870
         List            =   "FrmMstInsuranceCo.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   150
         Width           =   2535
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   10560
         TabIndex        =   24
         Top             =   150
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchInsuranceCo 
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
         Left            =   2295
         TabIndex        =   23
         Top             =   120
         Width           =   3645
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   9555
         TabIndex        =   22
         Top             =   150
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   60
         TabIndex        =   3
         Top             =   540
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   6015
         TabIndex        =   32
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Insurance Co"
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
         Left            =   120
         TabIndex        =   27
         Top             =   165
         Width           =   2895
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
         Left            =   11670
         TabIndex        =   26
         Top             =   165
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
         Left            =   14070
         TabIndex        =   25
         Top             =   165
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   6405
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5100
         TabIndex        =   17
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3885
         TabIndex        =   16
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2550
         TabIndex        =   15
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1335
         TabIndex        =   14
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5100
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3885
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2550
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmMstInsuranceCo.frx":0020
         Left            =   2385
         List            =   "FrmMstInsuranceCo.frx":002A
         Sorted          =   -1  'True
         TabIndex        =   28
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2160
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
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox TxtInsuranceCo 
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
      Begin VB.TextBox TxtInsuranceCoID 
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
         Left            =   2385
         TabIndex        =   30
         Top             =   1215
         Width           =   3735
      End
      Begin VB.Label Label3 
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
         Left            =   270
         TabIndex        =   29
         Top             =   1253
         Width           =   1695
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
         Top             =   2160
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
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Insurance Co Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Insurance Co ID"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMstInsuranceCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtInsuranceCo.Locked = False
TxtCreated = ""
TxtUpdated = ""

TxtInsuranceCo = ""
TxtInsuranceCoID = ""
TxtStatus.Visible = False
TxtStatus.Text = ""
CmbStatus.Text = ""
CmbStatus.Visible = True
If RsMst_InsuranceCo.RecordCount > 0 Then
   TxtInsuranceCo.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If CmbSStatus.Text <> "Active" And CmbSStatus.Text <> "Deactive" Then
   MsgBox "Invalid Selection"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If CmbStatus.Text <> "Active" And CmbStatus.Text <> "Deactive" Then
   MsgBox "Invalid Selection"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdGo_Click()
Dim mFlag As String
tmp = "Select InsuranceCoID,InsuranceCoName,Flag From Mst_InsuranceCoMaster "

Dim wc As Variant
wc = ""
If TxtSearchInsuranceCo <> "" Then
   wc = " Where InsuranceCoName Like '%" & TxtSearchInsuranceCo & "%' "
End If
If CmbSStatus.Text <> "" Then
   If CmbSStatus.Text = "Active" Then
      mFlag = "A"
   Else
      mFlag = "D"
   End If
   If wc = "" Then
      wc = " Where Flag='" & mFlag & "'"
   Else
      wc = wc & " And Flag='" & mFlag & "'"
   End If
End If

tmp = tmp & wc

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If MSHFlexGrid1.TextMatrix(iu, 2) = "D" Then
       MSHFlexGrid1.TextMatrix(iu, 2) = "Deactive"
    Else
       MSHFlexGrid1.TextMatrix(iu, 2) = "Active"
    End If
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   RsMst_InsuranceCo.Delete
   RsMst_InsuranceCo.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_InsuranceCo.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_InsuranceCo.MoveNext
   If RsMst_InsuranceCo.EOF() Then
      RsMst_InsuranceCo.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_InsuranceCo.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"

TxtStatus.Visible = False
CmbStatus.Visible = True
CmbStatus.Text = TxtStatus.Text
TxtStatus.Text = ""
TxtInsuranceCo.SetFocus

Call GButtonLock(Me, False)
TxtInsuranceCo.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_InsuranceCo.MoveFirst
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
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

'Call StatusBarData(Me)
Call Grid_Head
TxtSearchInsuranceCo = ""
Label22.Caption = ""

Call TableMst_InsuranceCo("  Order by InsuranceCoName ")
If RsMst_InsuranceCo.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtInsuranceCo.Locked = True
TxtStatus.Locked = True
TxtStatus.Visible = True
CmbStatus.Visible = False


TxtInsuranceCoID = RsMst_InsuranceCo!InsuranceCoID
TxtInsuranceCo = RsMst_InsuranceCo!InsuranceCoName


TxtStatus.Text = IIf(IIf(IsNull(RsMst_InsuranceCo!Flag), "A", RsMst_InsuranceCo!Flag) = "A", "Active", "Deactive")

TxtCreated = IIf(IsNull(RsMst_InsuranceCo!CreatedBy), "", RsMst_InsuranceCo!CreatedBy) & " :- " & IIf(IsNull(RsMst_InsuranceCo!CreatedDate), "", RsMst_InsuranceCo!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_InsuranceCo!UpdatedBy), "", RsMst_InsuranceCo!UpdatedBy) & " :- " & IIf(IsNull(RsMst_InsuranceCo!UpdatedDate), "", RsMst_InsuranceCo!UpdatedDate)


Call GButtonLock(Me, True)
End Sub


Private Sub LastCmd_Click()
RsMst_InsuranceCo.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_InsuranceCo.MoveFirst

If RsMst_InsuranceCo.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoID = " & Val(mCaId) & ""
If RsMst_InsuranceCo.EOF Then
   RsMst_InsuranceCo.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_InsuranceCo.MoveNext
If RsMst_InsuranceCo.EOF Then
   RsMst_InsuranceCo.MoveLast
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
RsMst_InsuranceCo.MovePrevious
If RsMst_InsuranceCo.BOF Then
   RsMst_InsuranceCo.MoveFirst
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

If TxtInsuranceCo = "" Then
   MsgBox "Blank Insurance Co Name not Allowed"
   TxtInsuranceCo.SetFocus
   Exit Sub
End If

If CmbStatus.Text = "" Then
   MsgBox "Select Status."
   CmbStatus.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_InsuranceCo.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_InsuranceCo.MoveFirst
        RsMst_InsuranceCo.Find "InsuranceCoName = '" & TxtInsuranceCo & "'"
        If Not RsMst_InsuranceCo.EOF Then
           MsgBox "Duplicate Insurance Co Name not Allowed"
           Exit Sub
        End If
        RsMst_InsuranceCo.MoveLast
        
        tmp = "select max(InsuranceCoID) from Mst_InsuranceCoMaster"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_InsuranceCo.AddNew
    RsMst_InsuranceCo!InsuranceCoID = tmp
    RsMst_InsuranceCo!G_UID = GetGUID
    TxtInsuranceCoID = tmp
Else
    RsMst_InsuranceCo.MoveFirst
    RsMst_InsuranceCo.Find "InsuranceCoName = '" & TxtInsuranceCo & "'"
    If Not RsMst_InsuranceCo.EOF Then
       If RsMst_InsuranceCo!InsuranceCoID <> TxtInsuranceCoID Then
          MsgBox "Duplicate Insurance Co Name not Allowed"
          Exit Sub
       End If
    End If
   RsMst_InsuranceCo.MoveFirst
   RsMst_InsuranceCo.Find "InsuranceCoID = " & TxtInsuranceCoID & ""
End If
RsMst_InsuranceCo!InsuranceCoName = TxtInsuranceCo

If CmbStatus.Text = "Active" Then
   RsMst_InsuranceCo!Flag = "A"
Else
   RsMst_InsuranceCo!Flag = "D"
End If

If IsNull(RsMst_InsuranceCo!CreatedBy) = True Or RsMst_InsuranceCo!CreatedBy = "" Then
   RsMst_InsuranceCo!CreatedBy = Gen_UserLoginID
   RsMst_InsuranceCo!CreatedDate = Now
Else
   RsMst_InsuranceCo!UpdatedBy = Gen_UserLoginID
   RsMst_InsuranceCo!UpdatedDate = Now
End If

RsMst_InsuranceCo.Update

RsMst_InsuranceCo.Requery
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoID = " & TxtInsuranceCoID & ""
If RsMst_InsuranceCo.EOF Then
   RsMst_InsuranceCo.MoveFirst
End If

Call GButtonLock(Me, True)
TxtInsuranceCo.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_InsuranceCo.EOF Then
      RsMst_InsuranceCo.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 400
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 800
RsMst_InsuranceCo.Requery
RsMst_InsuranceCo.MoveFirst
''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'TxtSearchInsuranceCo = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============


End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3

MSHFlexGrid1.TextMatrix(0, 0) = "InsuranceCo Id"
MSHFlexGrid1.TextMatrix(0, 1) = "InsuranceCo"
MSHFlexGrid1.TextMatrix(0, 2) = "Status"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 9500
MSHFlexGrid1.ColWidth(2) = 1500

End Sub

Private Sub TxtInsuranceCo_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtInsuranceCo)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtInsuranceCo.SetFocus
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_InsuranceCo.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_InsuranceCo.xls")
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
