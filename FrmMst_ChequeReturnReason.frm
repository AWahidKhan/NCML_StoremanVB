VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ChequeReturnReason 
   Caption         =   "Cheque dishonour Reason Master"
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
      Height          =   3645
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   15000
      Begin VB.TextBox TxtSReason 
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
         Height          =   390
         Left            =   2505
         MaxLength       =   500
         TabIndex        =   26
         Top             =   195
         Width           =   7260
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
         Left            =   12630
         TabIndex        =   24
         Top             =   215
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
         Height          =   345
         Left            =   11625
         TabIndex        =   23
         Top             =   215
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   25
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label1 
         Caption         =   "Cheque dishonour reason"
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
         Left            =   105
         TabIndex        =   27
         Top             =   263
         Width           =   2385
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   21
      Top             =   3420
      Width           =   8685
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   6645
         TabIndex        =   6
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5010
         TabIndex        =   9
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3375
         TabIndex        =   8
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1740
         TabIndex        =   4
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   0
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   6645
         TabIndex        =   5
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5010
         TabIndex        =   13
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3375
         TabIndex        =   12
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1740
         TabIndex        =   11
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   10
         Top             =   660
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8685
      Begin VB.TextBox TxtCharges 
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
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1740
         Width           =   2280
      End
      Begin VB.CheckBox ChkGenerateAdHoc 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox TxtReason 
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   765
         Width           =   6570
      End
      Begin VB.TextBox TxtCRRID 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   2280
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2835
         Width           =   4680
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2280
         Width           =   4680
      End
      Begin VB.Label Label3 
         Caption         =   "Cheque dishonour Charges"
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
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Collect Charges"
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
         TabIndex        =   28
         Top             =   1320
         Width           =   1860
      End
      Begin VB.Label LblAuditPoint 
         Caption         =   "Cheque dishonour reason"
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
         Left            =   120
         TabIndex        =   20
         Top             =   698
         Width           =   1860
      End
      Begin VB.Label LblAuditPointID 
         Caption         =   "Cheque Dishonour Reason ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Width           =   1920
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   18
         Top             =   2895
         Width           =   885
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   17
         Top             =   2340
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmMst_ChequeReturnReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mCRRID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

'TxtCharges.Locked = False
TxtReason.Locked = False
TxtCharges = ""
TxtReason = ""
TxtCRRID = ""
TxtCreated = ""
TxtUpdated = ""
ChkGenerateAdHoc.Enabled = True
ChkGenerateAdHoc.Value = vbUnchecked

If RsMst_ChequeReturnReason.RecordCount > 0 Then
   TxtReason.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub


Private Sub ChkGenerateAdHoc_Click()
TxtCharges.Text = 0
If ChkGenerateAdHoc.Value = vbUnchecked Then
    TxtCharges.Locked = True
Else
    TxtCharges.Locked = False
End If
End Sub

Private Sub CmdGo_Click()

Dim wc As Variant

Call Grid_Head

tmp = "Select * from Mst_ChequeReturnReason "

If TxtSReason <> "" Then
   wc = " Where Reason like '%" & TxtSReason & "%' "
   
End If

tmp = tmp & wc

Call TmpTable


MSHFlexGrid1.Rows = TmpRs.RecordCount + 2

For I = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!CRRID), "", TmpRs!CRRID)
    MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Reason), "", TmpRs!Reason) 'TmpRs!Reason
    MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!GenerateAdHoc), "", TmpRs!GenerateAdHoc) 'TmpRs!GenerateAdHoc
    If MSHFlexGrid1.TextMatrix(I, 2) = "Y" Then
        MSHFlexGrid1.TextMatrix(I, 2) = "Yes"
    ElseIf MSHFlexGrid1.TextMatrix(I, 2) = "N" Then
        MSHFlexGrid1.TextMatrix(I, 2) = "No"
    End If
    MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ReturnCharges), 0, TmpRs!ReturnCharges)
    TmpRs.MoveNext
Next

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   RsMst_ChequeReturnReason.Delete
   RsMst_ChequeReturnReason.Update
   If RsMst_ChequeReturnReason.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ChequeReturnReason.MoveNext
   If RsMst_ChequeReturnReason.EOF() Then
      RsMst_ChequeReturnReason.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_ChequeReturnReason.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"

Call GButtonLock(Me, False)

TxtReason.Locked = False
'TxtCharges.Locked = False
ChkGenerateAdHoc.Enabled = True
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_ChequeReturnReason.MoveFirst
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
   Frame1.Enabled = False
   ButtonFrm.Enabled = False
   Exit Sub
End If
MSHFlexGrid1.Clear
Call Grid_Head
TxtSReason = ""

Frame1.Enabled = True
ButtonFrm.Enabled = True
Call TableMst_ChequeReturnReason(" Order by CRRID")
If RsMst_ChequeReturnReason.RecordCount = 0 Then
'   If UserType = "Auditor" Then
'      Call GButtonLock(Me, False)
'      SaveCmd.Enabled = False
'      SearchCmd.Enabled = False
'      Exit Sub
'   End If
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()
TxtReason.Locked = True
'TxtCharges.Locked = True
ChkGenerateAdHoc.Enabled = False

TxtCRRID = IIf(IsNull(RsMst_ChequeReturnReason!CRRID), "", RsMst_ChequeReturnReason!CRRID)
TxtReason = IIf(IsNull(RsMst_ChequeReturnReason!Reason), "", RsMst_ChequeReturnReason!Reason)
If RsMst_ChequeReturnReason!GenerateAdHoc = "Y" Then
    ChkGenerateAdHoc.Value = vbChecked
ElseIf RsMst_ChequeReturnReason!GenerateAdHoc = "N" Then
    ChkGenerateAdHoc.Value = vbUnchecked
End If
TxtCharges = IIf(IsNull(RsMst_ChequeReturnReason!ReturnCharges), "", RsMst_ChequeReturnReason!ReturnCharges)
TxtCreated = IIf(IsNull(RsMst_ChequeReturnReason!CreatedBy), "", RsMst_ChequeReturnReason!CreatedBy) & " :- " & IIf(IsNull(RsMst_ChequeReturnReason!CreatedDate), "", RsMst_ChequeReturnReason!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ChequeReturnReason!UpdatedBy), "", RsMst_ChequeReturnReason!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ChequeReturnReason!UpdatedDate), "", RsMst_ChequeReturnReason!UpdatedDate)

Call GButtonLock(Me, True)
End Sub



Private Sub LastCmd_Click()
RsMst_ChequeReturnReason.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False

End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_ChequeReturnReason.MoveFirst

If RsMst_ChequeReturnReason.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_ChequeReturnReason.MoveFirst
RsMst_ChequeReturnReason.Find "CRRID = " & Val(mCaId) & ""
If RsMst_ChequeReturnReason.EOF Then
   RsMst_ChequeReturnReason.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ChequeReturnReason.MoveNext
If RsMst_ChequeReturnReason.EOF Then
   RsMst_ChequeReturnReason.MoveLast
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

RsMst_ChequeReturnReason.MovePrevious
If RsMst_ChequeReturnReason.BOF Then
   RsMst_ChequeReturnReason.MoveFirst
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
Private Function GetMaxCRRID() As Double
Dim Retval As Variant
Retval = 0
tmp = "select max(CRRID) from Mst_ChequeReturnReason "
Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs4.Fields(0)), 0, TmpRs4.Fields(0))
End If
TmpRs.Close
Retval = Retval + 1

GetMaxCRRID = Retval
End Function

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtReason = "" Then
   MsgBox "Blank Reason not Allowed"
   TxtReason.SetFocus
   Exit Sub
End If

If ChkGenerateAdHoc.Value = vbChecked Then
    If TxtCharges = "" Then
       MsgBox "Blank Charges not Allowed"
       TxtCharges.SetFocus
       Exit Sub
    End If
End If
If Edit_Flg = "A" Then
   If RsMst_ChequeReturnReason.RecordCount > 0 Then
      RsMst_ChequeReturnReason.MoveFirst
      tmp = "Reason = '" & TxtReason.Text & "'"
      RsMst_ChequeReturnReason.Filter = tmp
      If Not RsMst_ChequeReturnReason.EOF Then
         MsgBox "Duplicate Reason Not Allowed !!!"
         TxtReason.SetFocus
         RsMst_ChequeReturnReason.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_ChequeReturnReason.Filter = ""
   RsMst_ChequeReturnReason.AddNew
   RsMst_ChequeReturnReason!CRRID = GetMaxCRRID
   TxtCRRID = RsMst_ChequeReturnReason!CRRID
   RsMst_ChequeReturnReason!G_UID = GetGUID
Else
   RsMst_ChequeReturnReason.MoveFirst
   tmp = "Reason = '" & TxtReason.Text & "'"
   'RsMst_ChequeReturnReason.Find "ClientID = '" & TxtClientID.Text & "'"
   RsMst_ChequeReturnReason.Filter = tmp
   If Not RsMst_ChequeReturnReason.EOF Then
      If RsMst_ChequeReturnReason!CRRID <> TxtCRRID.Text Then
         MsgBox "Duplicate Reason Not Allowed !!! "
         TxtReason.SetFocus
         Exit Sub
      End If
   End If
   RsMst_ChequeReturnReason.Filter = ""
   RsMst_ChequeReturnReason.MoveFirst
   RsMst_ChequeReturnReason.Find "CRRID = " & TxtCRRID
   If RsMst_ChequeReturnReason.EOF Then
      MsgBox "No Such Record Found!!! "
      Exit Sub
   End If
End If

RsMst_ChequeReturnReason!Reason = IIf(IsNull(TxtReason), "", TxtReason)
If ChkGenerateAdHoc.Value = vbChecked Then
    RsMst_ChequeReturnReason!GenerateAdHoc = "Y"
ElseIf ChkGenerateAdHoc.Value = vbUnchecked Then
    RsMst_ChequeReturnReason!GenerateAdHoc = "N"
End If
RsMst_ChequeReturnReason!ReturnCharges = Val(TxtCharges) '(IsNull(TxtReason), "", TxtReason)
If IsNull(RsMst_ChequeReturnReason!CreatedBy) = True Or RsMst_ChequeReturnReason!CreatedBy = "" Then
   RsMst_ChequeReturnReason!CreatedBy = Gen_UserLoginID
   RsMst_ChequeReturnReason!CreatedDate = Now
Else
   RsMst_ChequeReturnReason!UpdatedBy = Gen_UserLoginID
   RsMst_ChequeReturnReason!UpdatedDate = Now
End If

RsMst_ChequeReturnReason.Update
Call GButtonLock(Me, True)

RsMst_ChequeReturnReason.Requery
RsMst_ChequeReturnReason.MoveFirst
RsMst_ChequeReturnReason.Find "CRRID = " & TxtCRRID & ""
If RsMst_ChequeReturnReason.EOF Then
   RsMst_ChequeReturnReason.MoveLast
End If
Call Indata

End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_ChequeReturnReason.EOF Then
      RsMst_ChequeReturnReason.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Left = 120
Frame0.Top = 120

Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 800

RsMst_ChequeReturnReason.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head
'TxtSReason = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4

MSHFlexGrid1.TextMatrix(0, 0) = "Reason Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Reason"
MSHFlexGrid1.TextMatrix(0, 2) = "Collect Charges"
MSHFlexGrid1.TextMatrix(0, 3) = "Charges"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 8000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000


End Sub

Private Sub CmdExcel_Click()
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ChequeReturnReason.xls") '("Mst_ChequeReturnReason.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ChequeReturnReason.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 11.29
oWS.Range("b:b").ColumnWidth = 100
oWS.Range("c:c").ColumnWidth = 10.43
oWS.Range("d:d").ColumnWidth = 11.71

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "Cheque Return Reason"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True

For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        If I = 0 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, c + 1).WrapText = True
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub TxtCharges_LostFocus()
If TxtCharges = "" Then Exit Sub
If IsNumeric(TxtCharges) = False Or Val(TxtCharges) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtCharges.SetFocus
   Exit Sub
End If
End Sub
