VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Assayer 
   Caption         =   "Assayer Master"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   9450
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   9480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   18
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   1
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
      Height          =   3285
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   9285
      Begin VB.TextBox TxtAssayerID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtAssayerName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   4
         Top             =   615
         Width           =   6900
      End
      Begin VB.TextBox TxtAddress 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   6
         Top             =   990
         Width           =   6900
      End
      Begin VB.TextBox TxtTelephoneNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1365
         Width           =   6900
      End
      Begin VB.TextBox TxtMobile 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1740
         Width           =   6900
      End
      Begin VB.TextBox TxtEmailID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   12
         Top             =   2115
         Width           =   6900
      End
      Begin VB.TextBox TxtOrganization 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   14
         Top             =   2490
         Width           =   6900
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   16
         Top             =   2865
         Visible         =   0   'False
         Width           =   6900
      End
      Begin VB.Label LblAssayerID 
         Caption         =   "Assayer ID"
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblAssayerName 
         Caption         =   "Assayer Name"
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblAddress 
         Caption         =   "Address"
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label LblTelephoneNo 
         Caption         =   "Telephone No"
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   1365
         Width           =   1800
      End
      Begin VB.Label LblMobile 
         Caption         =   "Mobile"
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label LblEmailID 
         Caption         =   "EmailID"
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   2115
         Width           =   1800
      End
      Begin VB.Label LblOrganization 
         Caption         =   "Organization"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   2490
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   2865
         Visible         =   0   'False
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1140
      Left            =   120
      TabIndex        =   17
      Top             =   3405
      Width           =   9285
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   7320
         TabIndex        =   21
         Top             =   630
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   7320
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5520
         TabIndex        =   23
         Top             =   630
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5520
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3720
         TabIndex        =   25
         Top             =   630
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3720
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1920
         TabIndex        =   27
         Top             =   630
         Width           =   1815
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1920
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   29
         Top             =   630
         Width           =   1815
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmMst_Assayer"
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

Call TableMst_Assayer
If RsMst_Assayer.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Assayer.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Assayer.xls")
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
RsMst_Assayer.MoveNext
If RsMst_Assayer.EOF Then
   RsMst_Assayer.MoveLast
   LF_Flag = "Y"
End If
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Assayer.MovePrevious
If RsMst_Assayer.BOF Then
   RsMst_Assayer.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_Assayer.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Assayer.MoveLast
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
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 1000
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_Assayer.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_Assayer.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For I = 1 To RsMst_Assayer.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(RsMst_Assayer.Fields(c)), "", RsMst_Assayer.Fields(c))
    Next
    RsMst_Assayer.MoveNext
Next
End Sub
Public Sub Indata()

TxtAssayerID.Locked = True
TxtAssayerName.Locked = True
TxtAddress.Locked = True
TxtTelephoneNo.Locked = True
TxtMobile.Locked = True
TxtEmailID.Locked = True
TxtOrganization.Locked = True
TxtFlag.Locked = True

TxtAssayerID = IIf(IsNull(RsMst_Assayer!AssayerID), "", RsMst_Assayer!AssayerID)
TxtAssayerName = IIf(IsNull(RsMst_Assayer!AssayerName), "", RsMst_Assayer!AssayerName)
TxtAddress = IIf(IsNull(RsMst_Assayer!Address), "", RsMst_Assayer!Address)
TxtTelephoneNo = IIf(IsNull(RsMst_Assayer!TelephoneNo), "", RsMst_Assayer!TelephoneNo)
TxtMobile = IIf(IsNull(RsMst_Assayer!Mobile), "", RsMst_Assayer!Mobile)
TxtEmailID = IIf(IsNull(RsMst_Assayer!EmailID), "", RsMst_Assayer!EmailID)
TxtOrganization = IIf(IsNull(RsMst_Assayer!Organization), "", RsMst_Assayer!Organization)
TxtFlag = IIf(IsNull(RsMst_Assayer!Flag), "", RsMst_Assayer!Flag)
If RsMst_Assayer.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtAssayerID_LostFocus()
If TxtAssayerID = "" Then Exit Sub
If IsNumeric(TxtAssayerID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtAssayerID.SetFocus
   Exit Sub
End If
TxtAssayerID = Val(TxtAssayerID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtAssayerID.Locked = True
TxtAssayerName.Locked = False
TxtAddress.Locked = False
TxtTelephoneNo.Locked = False
TxtMobile.Locked = False
TxtEmailID.Locked = False
TxtOrganization.Locked = False
TxtFlag.Locked = False

TxtAssayerID = ""
TxtAssayerName = ""
TxtAddress = ""
TxtTelephoneNo = ""
TxtMobile = ""
TxtEmailID = ""
TxtOrganization = ""
TxtFlag = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtAssayerID.Locked = True
TxtAssayerName.Locked = False
TxtAddress.Locked = False
TxtTelephoneNo.Locked = False
TxtMobile.Locked = False
TxtEmailID.Locked = False
TxtOrganization.Locked = False
TxtFlag.Locked = False

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_Assayer.Delete
   RsMst_Assayer.Update
   If RsMst_Assayer.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Assayer.MoveNext
   If RsMst_Assayer.EOF() Then
      RsMst_Assayer.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Assayer.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtAssayerName.Text = "" Then
   MsgBox "Blank AssayerName Not Allowed !!! "
   TxtAssayerName.SetFocus
   Exit Sub
End If


If TxtAddress.Text = "" Then
   MsgBox "Blank Address Not Allowed !!! "
   TxtAddress.SetFocus
   Exit Sub
End If


If TxtTelephoneNo.Text = "" Then
   MsgBox "Blank TelephoneNo Not Allowed !!! "
   TxtTelephoneNo.SetFocus
   Exit Sub
End If


If TxtMobile.Text = "" Then
   MsgBox "Blank Mobile Not Allowed !!! "
   TxtMobile.SetFocus
   Exit Sub
End If


If TxtEmailID.Text = "" Then
   MsgBox "Blank EmailID Not Allowed !!! "
   TxtEmailID.SetFocus
   Exit Sub
End If


If TxtOrganization.Text = "" Then
   MsgBox "Blank Organization Not Allowed !!! "
   TxtOrganization.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Assayer.RecordCount > 0 Then
      RsMst_Assayer.MoveFirst
      RsMst_Assayer.Find "AssayerName= '" & TxtAssayerName.Text & "'"
      If Not RsMst_Assayer.EOF Then
        MsgBox "Duplicate Assayer Name Not Allowed !!! "
         TxtAssayerName.SetFocus
         Exit Sub
      End If
   End If
      RsMst_Assayer.AddNew
      RsMst_Assayer!AssayerID = GetAssayerID
      RsMst_Assayer!G_UID = GetGUID
      TxtAssayerID.Text = RsMst_Assayer!AssayerID
Else
    RsMst_Assayer.MoveFirst
    RsMst_Assayer.Find "AssayerName= '" & TxtAssayerName.Text & "'"
    If Not RsMst_Assayer.EOF Then
       If RsMst_Assayer!AssayerID <> TxtAssayerID.Text Then
          MsgBox "Duplicate Assayer Name Not Allowed !!! "
          TxtAssayerName.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Assayer.MoveFirst
    RsMst_Assayer.Find " AssayerID= " & TxtAssayerID.Text
End If

RsMst_Assayer!AssayerID = TxtAssayerID.Text
RsMst_Assayer!AssayerName = TxtAssayerName.Text
RsMst_Assayer!Address = TxtAddress.Text
RsMst_Assayer!TelephoneNo = TxtTelephoneNo.Text
RsMst_Assayer!Mobile = TxtMobile.Text
RsMst_Assayer!EmailID = TxtEmailID.Text
RsMst_Assayer!Organization = TxtOrganization.Text
RsMst_Assayer!Flag = TxtFlag.Text
RsMst_Assayer.Update
Call Indata
End Sub
Private Function GetAssayerID() As Double
Dim Retval As Double
tmp = "Select max(AssayerID) from Mst_Assayer"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetAssayerID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8
MSHFlexGrid1.TextMatrix(0, 0) = "Assayer ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Assayer Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Address"
MSHFlexGrid1.TextMatrix(0, 3) = "Telephone No"
MSHFlexGrid1.TextMatrix(0, 4) = "Mobile"
MSHFlexGrid1.TextMatrix(0, 5) = "EmailID"
MSHFlexGrid1.TextMatrix(0, 6) = "Organization"
MSHFlexGrid1.TextMatrix(0, 7) = "Flag"


MSHFlexGrid1.ColWidth(0) = 400
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 500

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_Assayer.MoveFirst
RsMst_Assayer.Find "AssayerID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_Assayer.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
