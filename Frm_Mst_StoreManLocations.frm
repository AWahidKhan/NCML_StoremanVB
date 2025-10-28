VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_StoreManLocations 
   Caption         =   "StoreManLocations Master"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   7455
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   9480
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   15
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   7
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
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7245
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1800
         Width           =   1560
      End
      Begin VB.TextBox TxtStoreManLocationsID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtStoreManLocations 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   1
         Top             =   615
         Width           =   5160
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   1920
         MaxLength       =   255
         TabIndex        =   2
         Top             =   990
         Width           =   5160
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1725
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Flag"
         Height          =   300
         Left            =   4800
         TabIndex        =   24
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label LblStoreManLocationsID 
         Caption         =   "StoreManLocationsID"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblStoreManLocations 
         Caption         =   "Store Man Locations"
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblAddress 
         Caption         =   "Address"
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   1110
         Width           =   1800
      End
      Begin VB.Label LblDRC_Name 
         Caption         =   "PreFix Code"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1725
         Width           =   1320
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   7245
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   6
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   23
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   8
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   9
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2895
         TabIndex        =   10
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   11
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   12
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   5
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   13
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   14
         Top             =   240
         Width           =   1400
      End
   End
End
Attribute VB_Name = "FrmMst_StoreManLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_StoreManLocations
If RsMst_StoreManLocations.RecordCount = 0 Then
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
Call Gen_Create_Xls("Mst_StoreManLocations.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & "Mst_StoreManLocations.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
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
RsMst_StoreManLocations.MoveNext
If RsMst_StoreManLocations.EOF Then
   RsMst_StoreManLocations.MoveLast
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
RsMst_StoreManLocations.MovePrevious
If RsMst_StoreManLocations.BOF Then
   RsMst_StoreManLocations.MoveFirst
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
RsMst_StoreManLocations.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_StoreManLocations.MoveLast
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
   Call GButtonLock(Me, True)
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_StoreManLocations.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_StoreManLocations.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For I = 1 To RsMst_StoreManLocations.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(RsMst_StoreManLocations.Fields(C)), "", RsMst_StoreManLocations.Fields(C))
    Next
    RsMst_StoreManLocations.MoveNext
Next
End Sub
Public Sub Indata()

TxtStoreManLocationsID.Locked = True
TxtStoreManLocations.Locked = True
TxtAddress.Locked = True
TxtPreFix.Locked = True
TxtFlag.Locked = True

TxtStoreManLocationsID = IIf(IsNull(RsMst_StoreManLocations!StoreManLocationsID), "", RsMst_StoreManLocations!StoreManLocationsID)
TxtStoreManLocations = IIf(IsNull(RsMst_StoreManLocations!StoreManLocations), "", RsMst_StoreManLocations!StoreManLocations)
TxtAddress = IIf(IsNull(RsMst_StoreManLocations!Address), "", RsMst_StoreManLocations!Address)
TxtPreFix = IIf(IsNull(RsMst_StoreManLocations!SM_Prefix), "", RsMst_StoreManLocations!SM_Prefix)
TxtFlag = IIf(IsNull(RsMst_StoreManLocations!Flag), "", RsMst_StoreManLocations!Flag)

If RsMst_StoreManLocations.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtFlag_LostFocus()
If TxtFlag <> "D" And TxtFlag <> "C" Then
   MsgBox "Enter D or C !!!"
   TxtFlag.SetFocus
End If
End Sub

Private Sub TxtStoreManLocationsID_LostFocus()
If TxtStoreManLocationsID = "" Then Exit Sub
If IsNumeric(TxtStoreManLocationsID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtStoreManLocationsID.SetFocus
   Exit Sub
End If
TxtStoreManLocationsID = Val(TxtStoreManLocationsID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtStoreManLocationsID.Locked = True
TxtStoreManLocations.Locked = False
TxtAddress.Locked = False
TxtPreFix.Locked = False
TxtFlag.Locked = False

TxtStoreManLocationsID = ""
TxtStoreManLocations = ""
TxtAddress = ""
TxtPreFix = ""
TxtArea_Manager = ""
TxtFlag = "C"
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtStoreManLocationsID.Locked = True
TxtStoreManLocations.Locked = False
TxtAddress.Locked = False
TxtPreFix.Locked = False
TxtFlag.Locked = False

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_StoreManLocations.Delete
   RsMst_StoreManLocations.Update
   If RsMst_StoreManLocations.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_StoreManLocations.MoveNext
   If RsMst_StoreManLocations.EOF() Then
      RsMst_StoreManLocations.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_StoreManLocations.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtStoreManLocations.Text = "" Then
    MsgBox "Blank StoreManLocation Not Allowed!!!"
    TxtStoreManLocations.SetFocus
    Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_StoreManLocations.RecordCount > 0 Then
      RsMst_StoreManLocations.MoveFirst
      RsMst_StoreManLocations.Find "StoreManLocations= '" & TxtStoreManLocations.Text & "'"
      If Not RsMst_StoreManLocations.EOF Then
        MsgBox "Duplicate StoreManLocation Name Not Allowed !!! "
         TxtStoreManLocations.SetFocus
         Exit Sub
      End If
   End If
      RsMst_StoreManLocations.AddNew
      RsMst_StoreManLocations!StoreManLocationsID = GetStoreManLocationsID
      RsMst_StoreManLocations!G_UID = GetGUID
      TxtStoreManLocationsID.Text = RsMst_StoreManLocations!StoreManLocationsID
Else
    RsMst_StoreManLocations.MoveFirst
    RsMst_StoreManLocations.Find "StoreManLocations= '" & TxtStoreManLocations.Text & "'"
    If Not RsMst_StoreManLocations.EOF Then
       If RsMst_StoreManLocations!StoreManLocationsID <> TxtStoreManLocationsID.Text Then
          MsgBox "Duplicate StoreManLocation Name Not Allowed !!! "
          TxtStoreManLocations.SetFocus
          Exit Sub
       End If
    End If
    RsMst_StoreManLocations.MoveFirst
    RsMst_StoreManLocations.Find " StoreManLocationsID= " & TxtStoreManLocationsID.Text
End If

RsMst_StoreManLocations!StoreManLocationsID = TxtStoreManLocationsID.Text
RsMst_StoreManLocations!StoreManLocations = TxtStoreManLocations.Text
RsMst_StoreManLocations!Address = TxtAddress.Text
RsMst_StoreManLocations!SM_Prefix = TxtPreFix.Text
RsMst_StoreManLocations!Flag = TxtFlag.Text

RsMst_StoreManLocations.Update
Call Indata
End Sub
Private Function GetStoreManLocationsID() As Double
Dim Retval As Double
tmp = "Select max(StoreManLocationsID) from Mst_StoreManLocations"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetStoreManLocationsID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.TextMatrix(0, 0) = "Store Man PreFix "
MSHFlexGrid1.TextMatrix(0, 1) = "Store Man Locations ID "
MSHFlexGrid1.TextMatrix(0, 2) = "Store Man Locations Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Address"
MSHFlexGrid1.TextMatrix(0, 4) = "Flag"

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "StoreManLocationsID = " & MSHFlexGrid1.TextMatrix(I, 1)
If Not RsMst_StoreManLocations.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
