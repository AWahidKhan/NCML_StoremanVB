VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_BillingCycle 
   Caption         =   "Billing Cycle"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   9390
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   14880
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   465
         Left            =   2895
         TabIndex        =   23
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   465
         Left            =   4020
         TabIndex        =   25
         Top             =   180
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   1296
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label LblTotalRecord 
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
         Height          =   255
         Left            =   8160
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label20 
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
         Left            =   5760
         TabIndex        =   24
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   8880
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   225
         TabIndex        =   6
         Top             =   240
         Width           =   1680
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
         Height          =   400
         Left            =   225
         TabIndex        =   9
         Top             =   645
         Width           =   1680
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
         Height          =   400
         Left            =   1905
         TabIndex        =   3
         Top             =   240
         Width           =   1680
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
         Height          =   400
         Left            =   1905
         TabIndex        =   10
         Top             =   645
         Width           =   1680
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
         Height          =   400
         Left            =   3585
         TabIndex        =   7
         Top             =   240
         Width           =   1680
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
         Height          =   405
         Left            =   3585
         TabIndex        =   11
         Top             =   645
         Width           =   1680
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
         Height          =   400
         Left            =   5265
         TabIndex        =   8
         Top             =   240
         Width           =   1680
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
         Height          =   400
         Left            =   5265
         TabIndex        =   12
         Top             =   645
         Width           =   1680
      End
      Begin VB.CommandButton SearchCmd 
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
         Height          =   400
         Left            =   6945
         TabIndex        =   5
         Top             =   240
         Width           =   1680
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
         Height          =   400
         Left            =   6945
         TabIndex        =   4
         Top             =   645
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox TxtNoofDays 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtBillingCycle 
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
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   2175
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5775
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5775
      End
      Begin VB.TextBox TxtBillingCycleID 
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
         Left            =   1560
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No of Days"
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
         TabIndex        =   27
         Top             =   1260
         Width           =   1170
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
         Left            =   120
         TabIndex        =   20
         Top             =   1680
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
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Billing Cycle"
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
         TabIndex        =   15
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Billing Cycle ID"
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
         TabIndex        =   14
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmMst_BillingCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtCreated = ""
TxtUpdated = ""

TxtBillingCycleID.Text = ""
TxtBillingCycle.Text = ""
TxtNoofDays.Text = ""

Frame1.Enabled = True

If RsMst_BillingCycle.RecordCount > 0 Then
   TxtBillingCycle.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub



Private Sub CmdGo_Click()
tmp = "Select * From Mst_BillingCycle"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For mRow = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For mCol = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(mRow, mCol) = IIf(IsNull(TmpRs.Fields(mCol)), "", TmpRs.Fields(mCol))
       Next
       TmpRs.MoveNext
   Next
End If
LblTotalRecord.Caption = TmpRs.RecordCount
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_BillingCycle.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_BillingCycle.xls")
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

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_BillingCycle.MoveFirst

If RsMst_BillingCycle.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycleID = " & Val(mCaId) & ""
If RsMst_BillingCycle.EOF Then
   RsMst_BillingCycle.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
Call GButtonLock(Me, False)
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_BillingCycle.MoveNext
If RsMst_BillingCycle.EOF Then
   RsMst_BillingCycle.MoveLast
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
RsMst_BillingCycle.MovePrevious
If RsMst_BillingCycle.BOF Then
   RsMst_BillingCycle.MoveFirst
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
RsMst_BillingCycle.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_BillingCycle.MoveLast
Call Indata
End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   RsMst_BillingCycle.Delete
   RsMst_BillingCycle.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_BillingCycle.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_BillingCycle.MoveNext
   If RsMst_BillingCycle.EOF() Then
      RsMst_BillingCycle.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_BillingCycle.CancelUpdate
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_BillingCycle.EOF Then
      RsMst_BillingCycle.MoveFirst
   End If
   Call Indata

   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000

RsMst_BillingCycle.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head
'LblTotalRecord.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============
End Sub
Private Sub SaveCmd_Click()

If TxtBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle not Allowed"
   TxtBillingCycle.SetFocus
   Exit Sub
End If

If TxtNoofDays.Text = "" Then
   MsgBox "Blank Number of Days not Allowed"
   TxtNoofDays.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_BillingCycle.RecordCount > 0 Then
      RsMst_BillingCycle.MoveFirst
      RsMst_BillingCycle.Find "BillingCycle = '" & TxtBillingCycle.Text & "'"
      If Not RsMst_BillingCycle.EOF Then
         MsgBox "Duplicate Billing Cycle Name not Allowed"
         TxtBillingCycle.SetFocus
         Exit Sub
      End If
      RsMst_BillingCycle.MoveLast
   End If
   RsMst_BillingCycle.AddNew
   RsMst_BillingCycle!BillingCycleID = GetBillingCycleID
   RsMst_BillingCycle!G_UID = GetGUID
   TxtBillingCycleID.Text = RsMst_BillingCycle!BillingCycleID
    
Else
   RsMst_BillingCycle.MoveFirst
   RsMst_BillingCycle.Find "BillingCycle = '" & TxtBillingCycle.Text & "'"
   If Not RsMst_BillingCycle.EOF Then
       If RsMst_BillingCycle!BillingCycleID <> TxtBillingCycleID.Text Then
          MsgBox "Duplicate Billing Cycle Name not Allowed"
          TxtBillingCycle.SetFocus
          Exit Sub
       End If
   End If
   RsMst_BillingCycle.MoveFirst
   RsMst_BillingCycle.Find "BillingCycleID = " & TxtBillingCycleID.Text & ""
End If

RsMst_BillingCycle!BillingCycle = TxtBillingCycle.Text
RsMst_BillingCycle!NoofDays = Val(TxtNoofDays)


If IsNull(RsMst_BillingCycle!CreatedBy) = True Or RsMst_BillingCycle!CreatedBy = "" Then
   RsMst_BillingCycle!CreatedBy = Gen_UserLoginID
   RsMst_BillingCycle!CreatedDate = Now
Else
   RsMst_BillingCycle!UpdatedBy = Gen_UserLoginID
   RsMst_BillingCycle!UpdatedDate = Now
End If

RsMst_BillingCycle.Update
RsMst_BillingCycle.Requery
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycleID = " & TxtBillingCycleID.Text & ""
If RsMst_BillingCycle.EOF Then
   RsMst_BillingCycle.MoveFirst
End If
Call GButtonLock(Me, True)
Call Indata
End Sub
Public Sub Indata()

TxtBillingCycleID.Text = RsMst_BillingCycle!BillingCycleID
TxtBillingCycle.Text = RsMst_BillingCycle!BillingCycle
TxtNoofDays.Text = IIf(IsNull(RsMst_BillingCycle!NoofDays), 0, RsMst_BillingCycle!NoofDays)
TxtCreated = IIf(IsNull(RsMst_BillingCycle!CreatedBy), "", RsMst_BillingCycle!CreatedBy) & " :- " & IIf(IsNull(RsMst_BillingCycle!CreatedDate), "", RsMst_BillingCycle!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_BillingCycle!UpdatedBy), "", RsMst_BillingCycle!UpdatedBy) & " :- " & IIf(IsNull(RsMst_BillingCycle!UpdatedDate), "", RsMst_BillingCycle!UpdatedDate)

Frame1.Enabled = False
Call GButtonLock(Me, True)
End Sub
Private Function GetBillingCycleID()
Dim Retval As Integer
tmp = "Select max(BillingCycleID) from Mst_BillingCycle"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetBillingCycleID = Retval
End Function
Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
LblTotalRecord.Caption = ""
Frame1.Enabled = True
Call TableMst_BillingCycle
If RsMst_BillingCycle.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 3
    .Font.Size = 10
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Billing Cycle ID"
    .TextMatrix(0, 1) = "Billing Cycle Name"
    .TextMatrix(0, 2) = "No of Days"
    .ColWidth(0) = 2000
    .ColWidth(1) = 4000
    .ColWidth(2) = 4000
    .RowHeight(0) = 400
End With
End Sub
Private Sub TxtNoofDays_LostFocus()
If TxtNoofDays.Text = "" Then Exit Sub
If IsNumeric(TxtNoofDays.Text) = False Or Val(TxtNoofDays.Text) < 0 Then
   MsgBox "Invalid Numeric Value !!"
   TxtNoofDays.SetFocus
   Exit Sub
End If
End Sub
