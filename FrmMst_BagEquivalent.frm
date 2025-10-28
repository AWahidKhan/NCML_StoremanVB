VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_BagEquivalent 
   Caption         =   "Bag Equivalent Master"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11445
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   4215
      Left            =   120
      TabIndex        =   28
      Top             =   4560
      Visible         =   0   'False
      Width           =   15015
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   6075
         TabIndex        =   31
         Top             =   270
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchCommodity 
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
         TabIndex        =   30
         Top             =   240
         Width           =   3645
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   7080
         TabIndex        =   29
         Top             =   270
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   75
         TabIndex        =   32
         Top             =   660
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3916
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         Left            =   11790
         TabIndex        =   35
         Top             =   285
         Width           =   75
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
         Left            =   9390
         TabIndex        =   34
         Top             =   285
         Width           =   2355
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Commodity"
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
         TabIndex        =   33
         Top             =   285
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   8880
      Begin VB.TextBox TxtPEquivalent 
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
         Left            =   6720
         MaxLength       =   100
         TabIndex        =   36
         Top             =   1575
         Width           =   1560
      End
      Begin VB.TextBox TxtBagEquivalentID 
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
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   2880
      End
      Begin VB.TextBox TxtEquivalent 
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
         Left            =   5100
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1575
         Width           =   1560
      End
      Begin VB.CommandButton CmdCommodityId 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TxtBagToSize 
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
         Left            =   3480
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1575
         Width           =   1560
      End
      Begin VB.TextBox TxtBagFromSize 
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
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1575
         Width           =   1440
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
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2565
         Width           =   6360
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
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2145
         Width           =   6360
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   690
         Width           =   6375
      End
      Begin VB.TextBox TxtCommodityID 
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   690
         Width           =   6360
      End
      Begin VB.Label Label4 
         Caption         =   "Project Equivalent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6750
         TabIndex        =   37
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label Label3 
         Caption         =   "Bag Equivalent ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   285
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Non Project Equivalent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5130
         TabIndex        =   26
         Top             =   1035
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Bag To Size"
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
         Left            =   3600
         TabIndex        =   25
         Top             =   1155
         Width           =   1320
      End
      Begin VB.Label LblCommodityID 
         Caption         =   "Commodity "
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
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label LblStateName 
         Caption         =   "Bag From Size"
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
         Left            =   1920
         TabIndex        =   23
         Top             =   1155
         Width           =   1560
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
         Height          =   270
         Left            =   255
         TabIndex        =   22
         Top             =   2610
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
         Height          =   285
         Left            =   255
         TabIndex        =   21
         Top             =   2190
         Width           =   1335
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   60
      TabIndex        =   7
      Top             =   3210
      Width           =   8880
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   8
         Top             =   240
         Width           =   1680
      End
   End
End
Attribute VB_Name = "FrmMst_BagEquivalent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mCommodityID As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtBagEquivalentID.Text = ""

TxtCommodityID.Locked = False

CmdCommodityId.Visible = True
Frame1.Enabled = True

TxtCommodityID.Text = ""
TxtBagFromSize.Text = ""
TxtBagToSize.Text = ""
TxtEquivalent.Text = ""
TxtPEquivalent = ""

TxtCreated = ""
TxtUpdated = ""

TxtCommodityID.Visible = False
CmbCommodityID.Visible = True
CmbCommodityID.Text = ""
CmdCommodityId.Visible = True
If RsMst_BagEquivalent.RecordCount > 0 Then
   CmbCommodityID.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
mCommodityID = Null
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
End Sub


Private Sub CmdCommodityID_Click()
FrmMst_Commodity.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsMst_BagEquivalent.Delete
   RsMst_BagEquivalent.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_BagEquivalent.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_BagEquivalent.MoveNext
   If RsMst_BagEquivalent.EOF() Then
      RsMst_BagEquivalent.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_BagEquivalent.CancelUpdate



End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
TxtCommodityID.Visible = False
CmbCommodityID.Visible = True
CmbCommodityID.Text = TxtCommodityID.Text
Call GButtonLock(Me, False)
TxtBagEquivalentID.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_BagEquivalent.MoveFirst
Call Indata
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

Call Grid_Head
TxtSearchCommodity = ""
Label22.Caption = ""

Call TableMst_BagEquivalent
If RsMst_BagEquivalent.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub

Public Sub Indata()
Frame1.Enabled = False

mCommodityID = RsMst_BagEquivalent!CommodityID
TxtBagEquivalentID.Text = RsMst_BagEquivalent!BagEquivalentID
TxtCommodityID.Text = GetCommodityName(RsMst_BagEquivalent!CommodityID)
TxtBagFromSize.Text = RsMst_BagEquivalent!BagFromSize
TxtBagToSize.Text = RsMst_BagEquivalent!BagToSize
TxtEquivalent.Text = RsMst_BagEquivalent!Equivalent
TxtPEquivalent.Text = IIf(IsNull(RsMst_BagEquivalent!ProjectEquivalent), 0, RsMst_BagEquivalent!ProjectEquivalent)

TxtCreated = IIf(IsNull(RsMst_BagEquivalent!CreatedBy), "", RsMst_BagEquivalent!CreatedBy) & " :- " & IIf(IsNull(RsMst_BagEquivalent!CreatedDate), "", RsMst_BagEquivalent!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_BagEquivalent!UpdatedBy), "", RsMst_BagEquivalent!UpdatedBy) & " :- " & IIf(IsNull(RsMst_BagEquivalent!UpdatedDate), "", RsMst_BagEquivalent!UpdatedDate)

TxtCommodityID.Visible = True
CmbCommodityID.Visible = False

CmdCommodityId.Visible = False
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_BagEquivalent.MoveLast
Call Indata
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_BagEquivalent.MoveNext
If RsMst_BagEquivalent.EOF Then
   RsMst_BagEquivalent.MoveLast
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
RsMst_BagEquivalent.MovePrevious
If RsMst_BagEquivalent.BOF Then
   RsMst_BagEquivalent.MoveFirst
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

If CmbCommodityID.Text = "" Then
   MsgBox "Blank Commodity Name Not Allowed"
   CmbCommodityID.SetFocus
   Exit Sub
End If

If TxtBagFromSize = "" Then
   MsgBox "Blank Bag From Size Not Allowed"
   TxtBagFromSize.SetFocus
   Exit Sub
End If

If TxtBagToSize = "" Then
   MsgBox "Blank Bag To Size Not Allowed"
   TxtBagToSize.SetFocus
   Exit Sub
End If

If TxtEquivalent = "" Then
   MsgBox "Blank Non Project Equivalent Not Allowed"
   TxtEquivalent.SetFocus
   Exit Sub
End If

If TxtPEquivalent = "" Then
   MsgBox "Blank Project Equivalent Not Allowed"
   TxtPEquivalent.SetFocus
   Exit Sub
End If



If (Val(TxtBagFromSize.Text) > Val(TxtBagToSize.Text)) Then
   MsgBox "'Bag From Size' must be less than 'Bag To Size' !!!"
   TxtBagFromSize.SetFocus
   Exit Sub
End If

If CheckSelection() = False Then
   MsgBox "Bag Size values are not Allowed"
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsMst_BagEquivalent.Filter = ""
   RsMst_BagEquivalent.AddNew
   RsMst_BagEquivalent!BagEquivalentID = GetMaxBagEquivalentID
   RsMst_BagEquivalent!CommodityID = mCommodityID 'tmp
   RsMst_BagEquivalent!G_UID = GetGUID
   TxtBagEquivalentID = RsMst_BagEquivalent!BagEquivalentID
Else
   RsMst_BagEquivalent.MoveFirst
   RsMst_BagEquivalent.Find "BagEquivalentID = " & TxtBagEquivalentID & ""
End If

RsMst_BagEquivalent!BagFromSize = TxtBagFromSize.Text
RsMst_BagEquivalent!BagToSize = TxtBagToSize.Text

RsMst_BagEquivalent!Equivalent = TxtEquivalent.Text
RsMst_BagEquivalent!ProjectEquivalent = TxtPEquivalent.Text

If IsNull(RsMst_BagEquivalent!CreatedBy) = True Or RsMst_BagEquivalent!CreatedBy = "" Then
   RsMst_BagEquivalent!CreatedBy = Gen_UserLoginID
   RsMst_BagEquivalent!CreatedDate = Now
Else
   RsMst_BagEquivalent!UpdatedBy = Gen_UserLoginID
   RsMst_BagEquivalent!UpdatedDate = Now
End If


RsMst_BagEquivalent.Update
RsMst_BagEquivalent.Requery

RsMst_BagEquivalent.MoveFirst
RsMst_BagEquivalent.Find "BagEquivalentID = " & TxtBagEquivalentID & ""
If RsMst_BagEquivalent.EOF Then
   RsMst_BagEquivalent.MoveFirst
End If
Call GButtonLock(Me, True)
TxtBagEquivalentID.Locked = True
Call Indata
End Sub
Function CheckSelection() As Boolean
Dim Retval As Boolean
Retval = True

If RsMst_BagEquivalent.RecordCount = 0 Then
   RsMst_BagEquivalent.Filter = ""
   CheckSelection = Retval
   Exit Function
End If

RsMst_BagEquivalent.MoveFirst
If Edit_Flg = "A" Then
   RsMst_BagEquivalent.Filter = "CommodityID = " & mCommodityID & ""
Else
   RsMst_BagEquivalent.Filter = "CommodityID = " & mCommodityID & " and BagEquivalentID <> " & TxtBagEquivalentID & ""
End If

If RsMst_BagEquivalent.RecordCount > 0 Then
   For I = 1 To RsMst_BagEquivalent.RecordCount
       BagFSize = Val(TxtBagFromSize.Text)
       BagTSize = Val(TxtBagToSize.Text)
       DBFSize = RsMst_BagEquivalent!BagFromSize
       DBTSize = RsMst_BagEquivalent!BagToSize
       
       If (BagFSize > DBFSize And BagFSize < DBTSize) Or (BagTSize > DBFSize And BagTSize < DBTSize) Then
          RsMst_BagEquivalent.Filter = ""
          Retval = False
          CheckSelection = Retval
          Exit Function
       End If
   Next
End If

RsMst_BagEquivalent.Filter = ""
CheckSelection = Retval

End Function

Private Function GetMaxBagEquivalentID() As Double
Dim Retval As Double
tmp = "Select max(BagEquivalentID) from Mst_BagEquivalent"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxBagEquivalentID = Retval
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_BagEquivalent.EOF Then
      RsMst_BagEquivalent.MoveFirst
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

RsMst_BagEquivalent.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head
'TxtSearchCommodity = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .AllowUserResizing = flexResizeColumns
    .WordWrap = True
    .Cols = 6
    .Rows = 2
    .TextMatrix(0, 0) = "Equivalent ID"
    .TextMatrix(0, 1) = "Commodity"
    .TextMatrix(0, 2) = "Bag From Size"
    .TextMatrix(0, 3) = "Bag to Size"
    .TextMatrix(0, 4) = "Non Project Equivalent"
    .TextMatrix(0, 5) = "Project Equivalent"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .RowHeight(0) = 500
End With
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_BagEquivalent.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_BagEquivalent.xls")
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

Private Sub CmdGo_Click()
tmp = "Select * From Mst_BagEquivalent "

Dim Wc As Variant
Wc = ""
If TxtSearchCommodity <> "" Then
   Wc = " Where CommodityID in (Select CommodityID from Mst_Commodity where Commodity Like '%" & TxtSearchCommodity & "%') "
End If

tmp = tmp & Wc & " Order by CommodityID,BagFromSize,BagToSize"

Call TmpTable
Call Grid_Head

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        If c = 1 Then
           MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", GetCommodityName(TmpRs.Fields(c)))
        Else
           MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
        End If
        MSHFlexGrid1.TextMatrix(iu, 5) = IIf(IsNull(TmpRs!ProjectEquivalent), "", TmpRs!ProjectEquivalent)
    Next
    TmpRs.MoveNext
Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub
Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_BagEquivalent.MoveFirst

If RsMst_BagEquivalent.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_BagEquivalent.MoveFirst
RsMst_BagEquivalent.Find "BagEquivalentID = " & Val(mCaId) & ""
If RsMst_BagEquivalent.EOF Then
   RsMst_BagEquivalent.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub


Private Sub TxtEquivalent_LostFocus()
If TxtEquivalent.Text = "" Then Exit Sub
If IsNumeric(TxtEquivalent.Text) = False Or Val(TxtEquivalent.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtEquivalent.SetFocus
   Exit Sub
End If

TxtEquivalent = Round(Val(TxtEquivalent), 3)

End Sub

Private Sub TxtPEquivalent_LostFocus()
If TxtPEquivalent.Text = "" Then Exit Sub
If IsNumeric(TxtPEquivalent.Text) = False Or Val(TxtPEquivalent.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtPEquivalent.SetFocus
   Exit Sub
End If

TxtPEquivalent = Round(Val(TxtPEquivalent), 3)

End Sub

Private Sub TxtBagFromSize_LostFocus()
If TxtBagFromSize.Text = "" Then Exit Sub
If IsNumeric(TxtBagFromSize.Text) = False Or Val(TxtBagFromSize.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtBagFromSize.SetFocus
   Exit Sub
End If
End Sub


Private Sub TxtBagToSize_LostFocus()
If TxtBagToSize.Text = "" Then Exit Sub
If IsNumeric(TxtBagToSize.Text) = False Or Val(TxtBagToSize.Text) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtBagToSize.SetFocus
   Exit Sub
End If
End Sub
