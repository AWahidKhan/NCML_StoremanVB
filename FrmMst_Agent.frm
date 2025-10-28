VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Agent 
   Caption         =   "Agent Master"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   12195
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   14655
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   9000
         TabIndex        =   34
         Top             =   180
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchAgent 
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
         Left            =   2265
         TabIndex        =   33
         Top             =   150
         Width           =   5565
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   7995
         TabIndex        =   32
         Top             =   180
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Agent Name"
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
         TabIndex        =   37
         Top             =   195
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
         Left            =   10350
         TabIndex        =   36
         Top             =   195
         Width           =   2370
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         Left            =   12870
         TabIndex        =   35
         Top             =   195
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   9495
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4920
         Width           =   6615
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4440
         Width           =   6615
      End
      Begin VB.TextBox TxtContactNo 
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
         MaxLength       =   25
         TabIndex        =   3
         Top             =   2520
         Width           =   6615
      End
      Begin VB.TextBox TxtFax 
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
         MaxLength       =   15
         TabIndex        =   5
         Top             =   3480
         Width           =   6615
      End
      Begin VB.TextBox TxtEmailId 
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
         MaxLength       =   100
         TabIndex        =   6
         Top             =   3960
         Width           =   6615
      End
      Begin VB.TextBox TxtMobileNo 
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
         MaxLength       =   25
         TabIndex        =   4
         Top             =   3000
         Width           =   6615
      End
      Begin VB.TextBox TxtAddress 
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
         Height          =   1215
         Left            =   2415
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1185
         Width           =   6855
      End
      Begin VB.TextBox TxtAgentName 
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
         MaxLength       =   200
         TabIndex        =   1
         Top             =   720
         Width           =   6615
      End
      Begin VB.TextBox TxtAgentID 
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Updated"
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
         TabIndex        =   31
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Created  "
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
         TabIndex        =   30
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label LblContactNo 
         Caption         =   "Contact No"
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
         Top             =   2587
         Width           =   1695
      End
      Begin VB.Label LblFax 
         Caption         =   "Fax"
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
         TabIndex        =   26
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label LblEmailId 
         Caption         =   "EmailId"
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
         TabIndex        =   25
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label LblMobileNo 
         Caption         =   "Mobile No"
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
         TabIndex        =   24
         Top             =   3060
         Width           =   1695
      End
      Begin VB.Label LblAdess 
         Caption         =   "Address"
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
         Left            =   255
         TabIndex        =   23
         Top             =   1672
         Width           =   1695
      End
      Begin VB.Label LblAgentID 
         Caption         =   "Agent ID"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LblAgentName 
         Caption         =   "Agent Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   765
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   9510
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7485
         TabIndex        =   9
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   5670
         TabIndex        =   16
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7485
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5670
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3735
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmMst_Agent"
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

TxtAgentID = ""
TxtAgentName = ""
TxtAddress = ""
TxtContactNo = ""
TxtMobileNo = ""
TxtEmailID = ""
TxtFax = ""
Frame1.Enabled = True
If RsMst_Agent.RecordCount > 0 Then
   TxtAgentName.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmdGo_Click()
tmp = "Select AgentID,AgentName,Address,ContactNo,MobileNo,EmailID,Fax From Mst_Agent "

Dim Wc As Variant
Wc = ""
If TxtSearchAgent <> "" Then
   Wc = " Where AgentName Like '%" & TxtSearchAgent & "%' "
End If

tmp = tmp & Wc & " Order by AgentName"

Call TmpTable
Call GridHead
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

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsMst_Agent.Delete
   RsMst_Agent.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_Agent.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Agent.MoveNext
   If RsMst_Agent.EOF() Then
      RsMst_Agent.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Agent.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
Call GButtonLock(Me, False)
TxtAgentID.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_Agent.MoveFirst
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
Call GridHead
TxtSearchAgent = ""
Label22.Caption = ""
Call TableMst_Agent("  Order by AgentName ")
If RsMst_Agent.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()
TxtAgentID.Locked = True
TxtAgentID = RsMst_Agent!AgentID
TxtAgentName = RsMst_Agent!AgentName

TxtAddress = IIf(IsNull(RsMst_Agent!Address), "", RsMst_Agent!Address)
TxtContactNo = IIf(IsNull(RsMst_Agent!ContactNo), "", RsMst_Agent!ContactNo)
TxtMobileNo = IIf(IsNull(RsMst_Agent!MobileNo), "", RsMst_Agent!MobileNo)
TxtEmailID = IIf(IsNull(RsMst_Agent!EmailID), "", RsMst_Agent!EmailID)
TxtFax = IIf(IsNull(RsMst_Agent!Fax), "", RsMst_Agent!Fax)

TxtCreated = IIf(IsNull(RsMst_Agent!CreatedBy), "", RsMst_Agent!CreatedBy) & " :- " & IIf(IsNull(RsMst_Agent!CreatedDate), "", RsMst_Agent!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Agent!UpdatedBy), "", RsMst_Agent!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Agent!UpdatedDate), "", RsMst_Agent!UpdatedDate)



Frame1.Enabled = False
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_Agent.MoveLast
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

RsMst_Agent.MoveFirst

If RsMst_Agent.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "AgentID = " & Val(mCaId) & ""
If RsMst_Agent.EOF Then
   RsMst_Agent.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Agent.MoveNext
If RsMst_Agent.EOF Then
   RsMst_Agent.MoveLast
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
RsMst_Agent.MovePrevious
If RsMst_Agent.BOF Then
   RsMst_Agent.MoveFirst
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

If TxtAgentName = "" Then
   MsgBox "Blank Agent Name not Allowed!!!"
   TxtAgentName.SetFocus
   Exit Sub
End If

If Trim(TxtAddress) = "" Then
   MsgBox "Blank Agent Address not Allowed!!!"
   TxtAddress.SetFocus
   Exit Sub
End If

If Trim(TxtContactNo) = "" Then
   MsgBox "Blank Contact No not Allowed!!!"
   TxtContactNo.SetFocus
   Exit Sub
End If

If Trim(TxtMobileNo) = "" Then
   MsgBox "Blank Mobile No not Allowed!!!"
   TxtMobileNo.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Agent.RecordCount > 0 Then
      RsMst_Agent.MoveFirst
      RsMst_Agent.Find "AgentName = '" & TxtAgentName & "'"
      If Not RsMst_Agent.EOF Then
         MsgBox "Duplicate Agent Name not Allowed"
         Exit Sub
      End If
      RsMst_Agent.MoveLast
   End If
   RsMst_Agent.AddNew
   RsMst_Agent!AgentID = GetAgentID
   RsMst_Agent!G_UID = GetGUID
   TxtAgentID = RsMst_Agent!AgentID
    
Else
   RsMst_Agent.MoveFirst
   RsMst_Agent.Find "AgentName = '" & TxtAgentName & "'"
   If Not RsMst_Agent.EOF Then
       If RsMst_Agent!AgentID <> TxtAgentID Then
          MsgBox "Duplicate Agent Name not Allowed"
          Exit Sub
       End If
   End If
   RsMst_Agent.MoveFirst
   RsMst_Agent.Find "AgentID = " & TxtAgentID & ""
End If
RsMst_Agent!AgentName = TxtAgentName
RsMst_Agent!Address = Trim(TxtAddress)
RsMst_Agent!ContactNo = Trim(TxtContactNo)
RsMst_Agent!MobileNo = Trim(TxtMobileNo)
RsMst_Agent!EmailID = TxtEmailID
RsMst_Agent!Fax = TxtFax

If IsNull(RsMst_Agent!CreatedBy) = True Or RsMst_Agent!CreatedBy = "" Then
   RsMst_Agent!CreatedBy = Gen_UserLoginID
   RsMst_Agent!CreatedDate = Now
Else
   RsMst_Agent!UpdatedBy = Gen_UserLoginID
   RsMst_Agent!UpdatedDate = Now
End If

RsMst_Agent.Update
RsMst_Agent.Requery
RsMst_Agent.MoveFirst
RsMst_Agent.Find "AgentID = " & TxtAgentID & ""
If RsMst_Agent.EOF Then
   RsMst_Agent.MoveFirst
End If
Call GButtonLock(Me, True)
TxtAgentID.Locked = True
Call Indata
End Sub
Private Function GetAgentID() As Double
Dim Retval As Double
tmp = "Select max(AgentID) from Mst_Agent"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetAgentID = Retval
End Function
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_Agent.EOF Then
      RsMst_Agent.MoveFirst
   End If
   Call Indata

   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 120
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 120
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 700

RsMst_Agent.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call GridHead
'TxtSearchAgent = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============


End Sub
Private Sub GridHead()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 7
MSHFlexGrid1.TextMatrix(0, 0) = "Agent Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Agent Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Address"
MSHFlexGrid1.TextMatrix(0, 3) = "Contact No"
MSHFlexGrid1.TextMatrix(0, 4) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 5) = "Email ID"
MSHFlexGrid1.TextMatrix(0, 6) = "Fax"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4200
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 1400

End Sub
Private Sub TxtAgentID_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAgentID)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAgentID.SetFocus
End If
End Sub

Private Sub TxtAgentName_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAgentName)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAgentName.SetFocus
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Agent.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Agent.xls")
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
