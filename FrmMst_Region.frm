VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Region 
   Caption         =   "Region Master"
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
      Height          =   1305
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   14055
      Begin VB.ComboBox CmbSFlag 
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
         ItemData        =   "FrmMst_Region.frx":0000
         Left            =   5640
         List            =   "FrmMst_Region.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   4
         Text            =   "Active"
         Top             =   480
         Width           =   1620
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
         Height          =   350
         Left            =   7350
         TabIndex        =   3
         Top             =   450
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
         Left            =   8385
         TabIndex        =   2
         Top             =   450
         Width           =   1020
      End
      Begin VB.TextBox TxtSearchRegion 
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
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   465
         Width           =   5415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   5
         Top             =   885
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         Left            =   6120
         TabIndex        =   9
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Region"
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
         Left            =   2400
         TabIndex        =   8
         Top             =   210
         Width           =   660
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
         Left            =   9585
         TabIndex        =   7
         Top             =   465
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
         Left            =   11985
         TabIndex        =   6
         Top             =   465
         Width           =   75
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   7380
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
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2145
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
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1665
         Width           =   4680
      End
      Begin VB.ComboBox CmbStatusFlag 
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
         ItemData        =   "FrmMst_Region.frx":0020
         Left            =   2310
         List            =   "FrmMst_Region.frx":002A
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   1200
         Width           =   1905
      End
      Begin VB.TextBox TxtRegionID 
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
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtRegion 
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
         Left            =   2310
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox TxtStatusFlag 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2310
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   1905
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
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Top             =   2220
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
         Height          =   300
         Left            =   240
         TabIndex        =   31
         Top             =   1755
         Width           =   1695
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   1275
         Width           =   1800
      End
      Begin VB.Label Lable1 
         Caption         =   "Region ID"
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
         Top             =   300
         Width           =   1920
      End
      Begin VB.Label LblFlag 
         Caption         =   "Region "
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
         TabIndex        =   28
         Top             =   754
         Width           =   2280
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   7365
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
         Left            =   100
         TabIndex        =   20
         Top             =   240
         Width           =   1425
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
         Left            =   100
         TabIndex        =   19
         Top             =   660
         Width           =   1425
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
         Left            =   1523
         TabIndex        =   18
         Top             =   240
         Width           =   1425
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
         Left            =   1523
         TabIndex        =   17
         Top             =   660
         Width           =   1425
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
         Left            =   2955
         TabIndex        =   16
         Top             =   240
         Width           =   1425
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
         Height          =   400
         Left            =   2955
         TabIndex        =   15
         Top             =   660
         Width           =   1425
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
         Left            =   4380
         TabIndex        =   14
         Top             =   240
         Width           =   1425
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
         Left            =   4380
         TabIndex        =   13
         Top             =   660
         Width           =   1425
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
         Left            =   5805
         TabIndex        =   12
         Top             =   240
         Width           =   1425
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
         Left            =   5805
         TabIndex        =   11
         Top             =   660
         Width           =   1425
      End
   End
End
Attribute VB_Name = "FrmMst_Region"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant


Private Sub CmdGo_Click()

tmp = "Select RegionID,Region,Flag from Mst_Region "

Dim wc As Variant
wc = ""
If TxtSearchRegion <> "" Then
   wc = " Where Region Like '%" & TxtSearchRegion & "%' "
End If

If CmbSFlag.Text <> "" Then
   If wc = "" Then
      wc = " Where Flag='" & Left$(CmbSFlag.Text, 1) & "' " 'Left$(CmbStatusFlag.Text, 1)
   Else
      wc = wc & " And Flag='" & Left$(CmbSFlag.Text, 1) & "' " 'Left$(CmbStatusFlag.Text, 1)
   End If
End If

tmp = tmp & wc & " Order By Region "

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(iu, 2) = IIf(MSHFlexGrid1.TextMatrix(iu, 2) = "A", "Active", "Deactive")
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
   
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call TableMst_Region
If RsMst_Region.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Region.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Region.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_PerLimit As Variant
LF_PerLimit = "N"
RsMst_Region.MoveNext
If RsMst_Region.EOF Then
   RsMst_Region.MoveLast
   LF_PerLimit = "Y"
End If
Call Indata
If LF_PerLimit = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_PerLimit As Variant
LF_PerLimit = "N"
RsMst_Region.MovePrevious
If RsMst_Region.BOF Then
   RsMst_Region.MoveFirst
   LF_PerLimit = "Y"
End If
Call Indata
If LF_PerLimit = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_Region.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Region.MoveLast
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
   If RsMst_Region.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Exit Sub
End If

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 600
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_Region.MoveFirst
Call Grid_Head
TxtSearchRegion = ""
Label22.Caption = ""

End Sub
Public Sub Indata()

TxtStatusFlag.Locked = True
TxtStatusFlag.Visible = True
CmbStatusFlag.Visible = False
TxtRegionID.Locked = True
TxtRegion.Locked = True
TxtRegionID = IIf(IsNull(RsMst_Region!RegionID), "", RsMst_Region!RegionID)
TxtRegion = IIf(IsNull(RsMst_Region!Region), "", RsMst_Region!Region)
If LCase(RsMst_Region!Flag) = "a" Then
   TxtStatusFlag.Text = "Active"
ElseIf LCase(RsMst_Region!Flag) = "d" Then
   TxtStatusFlag.Text = "DeActive"
Else
   TxtStatusFlag.Text = ""
End If

TxtCreated = IIf(IsNull(RsMst_Region!CreatedBy), "", RsMst_Region!CreatedBy) & " :- " & IIf(IsNull(RsMst_Region!CreatedDate), "", RsMst_Region!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Region!UpdatedBy), "", RsMst_Region!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Region!UpdatedDate), "", RsMst_Region!UpdatedDate)


If RsMst_Region.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtRegionID.Locked = True
TxtRegion.Locked = False
TxtRegionID = ""
TxtRegion = ""

TxtCreated = ""
TxtUpdated = ""

CmbStatusFlag.Visible = True
TxtStatusFlag.Visible = False
CmbStatusFlag.Text = ""

If RsMst_Region.RecordCount > 0 Then
   TxtRegion.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtRegionID.Locked = True
TxtRegion.Locked = False
CmbStatusFlag.Visible = True
TxtStatusFlag.Visible = False
CmbStatusFlag.Text = TxtStatusFlag.Text

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_Region.Delete
   RsMst_Region.Update
   If RsMst_Region.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Region.MoveNext
   If RsMst_Region.EOF() Then
      RsMst_Region.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Region.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If TxtRegion.Text = "" Then
   MsgBox "Blank Region Not Allowed !!! "
   TxtRegion.SetFocus
   Exit Sub
End If

If CmbStatusFlag.Text = "" Then
   MsgBox "Please Select Status !!!"
   CmbStatusFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Region.RecordCount > 0 Then
      RsMst_Region.MoveFirst
      RsMst_Region.Find "Region= '" & TxtRegion.Text & "'"
      If Not RsMst_Region.EOF Then
        MsgBox "Duplicate Region Not Allowed !!! "
         TxtRegion.SetFocus
         Exit Sub
      End If
   End If
      RsMst_Region.AddNew
      RsMst_Region!RegionID = GetRegionID
      RsMst_Region!G_UID = GetGUID
      TxtRegionID.Text = RsMst_Region!RegionID
Else
    RsMst_Region.MoveFirst
    RsMst_Region.Find "Region= '" & TxtRegion.Text & "'"
    If Not RsMst_Region.EOF Then
       If RsMst_Region!RegionID <> TxtRegionID.Text Then
          MsgBox "Duplicate Region Not Allowed !!! "
          TxtRegion.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Region.MoveFirst
    RsMst_Region.Find " RegionID= " & TxtRegionID.Text
End If

RsMst_Region!RegionID = TxtRegionID.Text
RsMst_Region!Region = TxtRegion.Text
RsMst_Region!Flag = Left$(CmbStatusFlag.Text, 1)

If IsNull(RsMst_Region!CreatedBy) = True Or RsMst_Region!CreatedBy = "" Then
   RsMst_Region!CreatedBy = Gen_UserLoginID
   RsMst_Region!CreatedDate = Now
Else
   RsMst_Region!UpdatedBy = Gen_UserLoginID
   RsMst_Region!UpdatedDate = Now
End If


RsMst_Region.Update
Call Indata
End Sub
Private Function GetRegionID() As Double
Dim Retval As Double
tmp = "Select max(RegionID) from Mst_Region"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetRegionID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0, 0) = "Region ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Region"
MSHFlexGrid1.TextMatrix(0, 2) = "Status"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 4000
MSHFlexGrid1.ColWidth(2) = 2000


End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_Region.MoveFirst
RsMst_Region.Find "RegionID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_Region.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub CmbStatusFlag_LostFocus()
If CmbStatusFlag.Text = "" Then Exit Sub

If LCase(CmbStatusFlag.Text) = "active" Or LCase(CmbStatusFlag.Text) = "deactive" Then
Else
   MsgBox "Invalid selection!!!"
   CmbStatusFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSFlag_LostFocus()
If CmbSFlag.Text = "" Then Exit Sub

If LCase(CmbSFlag.Text) = "active" Or LCase(CmbSFlag.Text) = "deactive" Then
Else
   MsgBox "Invalid selection!!!"
   CmbSFlag.SetFocus
   Exit Sub
End If

End Sub



