VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_CommodityGroup 
   Caption         =   "Commodity Group"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1005
      Left            =   60
      TabIndex        =   0
      Top             =   3990
      Visible         =   0   'False
      Width           =   15000
      Begin VB.TextBox TxtSearchCommodityGroup 
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
         Left            =   2700
         TabIndex        =   28
         Top             =   150
         Width           =   3885
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   6990
         TabIndex        =   27
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   7995
         TabIndex        =   1
         Top             =   180
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   2
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
      Begin VB.Label Label19 
         Caption         =   "Search By Commodity Group"
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
         Left            =   75
         TabIndex        =   31
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
         Left            =   9345
         TabIndex        =   30
         Top             =   195
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
         Left            =   11745
         TabIndex        =   29
         Top             =   195
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   60
      TabIndex        =   12
      Top             =   2400
      Width           =   8880
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   22
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1740
         TabIndex        =   21
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3375
         TabIndex        =   20
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3375
         TabIndex        =   19
         Top             =   660
         Width           =   1755
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5130
         TabIndex        =   18
         Top             =   240
         Width           =   1755
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5130
         TabIndex        =   17
         Top             =   660
         Width           =   1755
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   6885
         TabIndex        =   16
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   6885
         TabIndex        =   15
         Top             =   660
         Width           =   1875
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1740
         TabIndex        =   14
         Top             =   660
         Width           =   1635
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   13
         Top             =   660
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   8880
      Begin VB.ComboBox CmbHRCommodityID 
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
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   1080
         Width           =   6870
      End
      Begin VB.TextBox TxtHRCommodityID 
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1080
         Width           =   6855
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
         Top             =   1950
         Width           =   6840
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
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1530
         Width           =   6840
      End
      Begin VB.TextBox TxtCommodityGroup 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   7
         Top             =   645
         Width           =   6840
      End
      Begin VB.TextBox TxtCommodityGroupID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtMaximum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   7200
         MaxLength       =   75
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox TxtMinimum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   4
         Top             =   2490
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label20 
         Caption         =   "HR Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1140
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2010
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   1590
         Width           =   1695
      End
      Begin VB.Label LblCommodity 
         Caption         =   "Group Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   705
         Width           =   1800
      End
      Begin VB.Label LblCommodityID 
         Caption         =   "Group Commodity ID"
         Height          =   225
         Left            =   240
         TabIndex        =   10
         Top             =   308
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Maximum bag size in Kg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   9
         Top             =   2550
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label Label2 
         Caption         =   "Minimum bag size in Kg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   8
         Top             =   2505
         Visible         =   0   'False
         Width           =   1320
      End
   End
End
Attribute VB_Name = "FrmMst_CommodityGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mExchangeTypeID, mHRCommodityID As Variant

Private Sub CmdGo_Click()

tmp = "Select CommodityGroupID,CommodityGroup,Isnull(HRCommodityID,0)'HRCommodityID',G_UID From Mst_CommodityGroup"

Dim Wc As Variant
Wc = ""
If TxtSearchCommodityGroup <> "" Then
   Wc = " Where CommodityGroup Like '%" & TxtSearchCommodityGroup & "%' "
End If

tmp = tmp & Wc & " Order By CommodityGroup"

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    MSHFlexGrid1.TextMatrix(iu, 2) = GetHRCommodity(Val(MSHFlexGrid1.TextMatrix(iu, 2)))
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
TxtSearchCommodityGroup = ""
Label22.Caption = ""

Call TableMst_CommodityGroup(" Order By CommodityGroup")

If RsMst_CommodityGroup.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_CommodityGroup.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_CommodityGroup.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 2
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
RsMst_CommodityGroup.MoveNext
If RsMst_CommodityGroup.EOF Then
   RsMst_CommodityGroup.MoveLast
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
RsMst_CommodityGroup.MovePrevious
If RsMst_CommodityGroup.BOF Then
   RsMst_CommodityGroup.MoveFirst
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
RsMst_CommodityGroup.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_CommodityGroup.MoveLast
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
   If RsMst_CommodityGroup.RecordCount = 0 Then Exit Sub
   
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 800
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_CommodityGroup.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchCommodityGroup = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

'MSHFlexGrid1.Rows = RsMst_CommodityGroup.RecordCount + 1
'
'For i = 1 To RsMst_CommodityGroup.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_CommodityGroup.Fields(c)), "", RsMst_CommodityGroup.Fields(c))
'    Next
'    RsMst_CommodityGroup.MoveNext
'Next
End Sub
Public Sub Indata()



TxtCommodityGroupID.Locked = True
TxtCommodityGroup.Locked = True
TxtMaximum.Locked = True
TxtMinimum.Locked = True
CmbHRCommodityID.Visible = False
TxtHRCommodityID.Visible = True

TxtGunnyWt = ""
TxtStdDeduction = ""
TxtLotSize = ""
TxtVariation = ""
TxtMaximum = ""
TxtMinimum = ""

mHRCommodityID = IIf(IsNull(RsMst_CommodityGroup!HRCommodityID), 0, RsMst_CommodityGroup!HRCommodityID)
TxtCommodityGroupID = IIf(IsNull(RsMst_CommodityGroup!CommodityGroupID), "", RsMst_CommodityGroup!CommodityGroupID)
TxtCommodityGroup = IIf(IsNull(RsMst_CommodityGroup!COmmodityGroup), "", RsMst_CommodityGroup!COmmodityGroup)
TxtHRCommodityID = GetHRCommodity(mHRCommodityID)
TxtCreated = IIf(IsNull(RsMst_CommodityGroup!CreatedBy), "", RsMst_CommodityGroup!CreatedBy) & " :- " & IIf(IsNull(RsMst_CommodityGroup!CreatedDate), "", RsMst_CommodityGroup!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_CommodityGroup!UpdatedBy), "", RsMst_CommodityGroup!UpdatedBy) & " :- " & IIf(IsNull(RsMst_CommodityGroup!UpdatedDate), "", RsMst_CommodityGroup!UpdatedDate)

If RsMst_CommodityGroup.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub TxtCommodityGroup_LostFocus()
TxtCommodityGroup = Trim(TxtCommodityGroup)
End Sub

Private Sub TxtCommodityGroupID_LostFocus()
If TxtCommodityGroupID = "" Then Exit Sub
If IsNumeric(TxtCommodityGroupID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtCommodityGroupID.SetFocus
   Exit Sub
End If
TxtCommodityGroupID = Val(TxtCommodityGroupID)
End Sub


Private Sub TxtMaximum_LostFocus()
If TxtMaximum = "" Then Exit Sub
If Not IsNumeric(TxtMaximum) Then
   MsgBox "Invalid numeric format !!!! "
   TxtMaximum.SetFocus
End If

End Sub


Private Sub TxtMinimum_LostFocus()
If TxtMinimum = "" Then Exit Sub
If Not IsNumeric(TxtMinimum) Then
   MsgBox "Invalid numeric format !!!! "
   TxtMinimum.SetFocus
End If
End Sub



Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
'Frame2.Enabled = True

TxtCommodityGroupID.Locked = True
TxtCommodityGroup.Locked = False
CmbExchangeTypeID = ""
TxtMaximum.Locked = False
TxtMinimum.Locked = False

TxtCommodityGroupID = ""
TxtCommodityGroup = ""
TxtGunnyWt = ""
TxtStdDeduction = ""
TxtLotSize = ""
TxtVariation = ""
TxtMinimum = ""
TxtMaximum = ""
'TxtCommodityGroup.SetFocus
TxtCreated = ""
TxtUpdated = ""

CmbHRCommodityID.Visible = True
TxtHRCommodityID.Visible = False
CmbHRCommodityID.Text = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
'Frame2.Enabled = True
TxtMaximum.Locked = False
TxtMinimum.Locked = False
TxtCommodityGroupID.Locked = True
TxtCommodityGroup.Locked = False
'TxtGunnyWt.Locked = False
'TxtStdDeduction.Locked = False
'TxtLotSize.Locked = False
'TxtVariation.Locked = False
TxtCommodityGroup.SetFocus

CmbHRCommodityID.Visible = True
TxtHRCommodityID.Visible = False
CmbHRCommodityID.Text = TxtHRCommodityID.Text
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   
   RsMst_CommodityGroup.Requery
   RsMst_CommodityGroup.Find "CommodityGroupID = " & TxtCommodityGroupID
   If RsMst_CommodityGroup.EOF Then Exit Sub
   
   RsMst_CommodityGroup.Delete
   RsMst_CommodityGroup.Update
   If RsMst_CommodityGroup.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_CommodityGroup.MoveNext
   If RsMst_CommodityGroup.EOF() Then
      RsMst_CommodityGroup.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_CommodityGroup.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtCommodityGroup.Text = "" Then
   MsgBox "Blank CommodityGroup Not Allowed !!! "
   TxtCommodityGroup.SetFocus
   Exit Sub
End If


If CmbHRCommodityID.Text = "" Then
   MsgBox "Blank HR Commodity not allowed !!! "
   CmbHRCommodityID.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_CommodityGroup.RecordCount > 0 Then
      RsMst_CommodityGroup.MoveFirst
      RsMst_CommodityGroup.Find "CommodityGroup= '" & TxtCommodityGroup.Text & "'"
      If Not RsMst_CommodityGroup.EOF Then
        MsgBox "Duplicate CommodityGroup Name Not Allowed !!! "
         TxtCommodityGroup.SetFocus
         Exit Sub
      End If
   End If
      RsMst_CommodityGroup.AddNew
      RsMst_CommodityGroup!CommodityGroupID = GetCommodityGroupID
      TxtCommodityGroupID.Text = RsMst_CommodityGroup!CommodityGroupID
      RsMst_CommodityGroup!G_UID = GetGUID
Else
    RsMst_CommodityGroup.MoveFirst
    RsMst_CommodityGroup.Find "CommodityGroup= '" & TxtCommodityGroup.Text & "'"
    If Not RsMst_CommodityGroup.EOF Then
       If RsMst_CommodityGroup!CommodityGroupID <> TxtCommodityGroupID.Text Then
          MsgBox "Duplicate CommodityGroup Name Not Allowed !!! "
          TxtCommodityGroup.SetFocus
          Exit Sub
       End If
    End If
    RsMst_CommodityGroup.MoveFirst
    RsMst_CommodityGroup.Find " CommodityGroupID= " & TxtCommodityGroupID.Text
End If

RsMst_CommodityGroup!CommodityGroupID = TxtCommodityGroupID.Text
RsMst_CommodityGroup!COmmodityGroup = TxtCommodityGroup.Text
RsMst_CommodityGroup!HRCommodityID = mHRCommodityID

If IsNull(RsMst_CommodityGroup!CreatedBy) = True Or RsMst_CommodityGroup!CreatedBy = "" Then
   RsMst_CommodityGroup!CreatedBy = Gen_UserLoginID
   RsMst_CommodityGroup!CreatedDate = Now
Else
   RsMst_CommodityGroup!UpdatedBy = Gen_UserLoginID
   RsMst_CommodityGroup!UpdatedDate = Now
End If


RsMst_CommodityGroup.Update
Call Indata
End Sub


Private Function GetCommodityGroupID() As Double
Dim Retval As Double
tmp = "Select max(CommodityGroupID) from Mst_CommodityGroup"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCommodityGroupID = Retval
End Function

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4
MSHFlexGrid1.TextMatrix(0, 0) = "Commodity Group ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Commodity Group"
MSHFlexGrid1.TextMatrix(0, 2) = "HR Commodity"
MSHFlexGrid1.TextMatrix(0, 3) = "G_UID"
'MSHFlexGrid1.TextMatrix(0, 2) = "Maximum bag size in Kg "
'MSHFlexGrid1.TextMatrix(0, 3) = "Minimum bag size in Kg "
'MSHFlexGrid1.TextMatrix(0, 4) = "Lot Size in MT"
'MSHFlexGrid1.TextMatrix(0, 5) = "Variation"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 4600
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 0
'MSHFlexGrid1.ColWidth(2) = 2000
'MSHFlexGrid1.ColWidth(3) = 2000
'MSHFlexGrid1.ColWidth(4) = 1200
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroupID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_CommodityGroup.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub CmbHRCommodityID_GotFocus()

tmp1 = CmbHRCommodityID.Text
tmp = "select * from Mst_Commodity "

Call TmpTableHRPro

CmbHRCommodityID.Clear
If TmpRsHRPro.RecordCount = 0 Then
   MsgBox "No HR Commodity found!!!"
   CmbHRCommodityID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRsHRPro.RecordCount
    CmbHRCommodityID.AddItem TmpRsHRPro!CommodityName
    TmpRsHRPro.MoveNext
Next
CmbHRCommodityID.Text = tmp1
End Sub

Private Sub CmbHRCommodityID_LostFocus()
If CmbHRCommodityID.Text = "" Then
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "CommodityName = '" & CmbHRCommodityID.Text & "'"

If TmpRsHRPro.EOF Then
   MsgBox "Invalid HR Commodity selection "
   CmbHRCommodityID.SetFocus
   Exit Sub
End If
mHRCommodityID = TmpRsHRPro!CommodityID
End Sub

