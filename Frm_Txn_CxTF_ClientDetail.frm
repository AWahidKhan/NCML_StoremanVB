VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_CxTF_ClientDetail 
   Caption         =   "Txn_CxTF_ClientDetail"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12495
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   4000
      Left            =   8000
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   10
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   8490
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5865
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10000
      Begin VB.TextBox TxtClientDetailID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   3600
      End
      Begin VB.TextBox TxtCDTFNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Top             =   615
         Width           =   3600
      End
      Begin VB.ComboBox CmbCDTFNo 
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   615
         Width           =   4000
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   19
         Top             =   990
         Width           =   3600
      End
      Begin VB.ComboBox CmbClientID 
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   990
         Width           =   4000
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   22
         Top             =   1365
         Width           =   3600
      End
      Begin VB.Label LblClientDetailID 
         Caption         =   "ClientDetail ID"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblCDTFNo 
         Caption         =   "CDTF No"
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client ID"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1365
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   23
      Top             =   7200
      Width           =   10000
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   7320
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   7320
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   5520
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   3720
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1920
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmTxn_CxTF_ClientDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mCDTFNo, mClientID As Variant
Private Sub Form_Load()
Call TableTxn_CxTF_ClientDetail
If RsTxn_CxTF_ClientDetail.RecordCount = 0 Then
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
Call InData
End Sub
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Txn_CxTF_ClientDetail.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Txn_CxTF_ClientDetail.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_ClientDetail.MoveNext
If RsTxn_CxTF_ClientDetail.EOF Then
   RsTxn_CxTF_ClientDetail.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_ClientDetail.MovePrevious
If RsTxn_CxTF_ClientDetail.BOF Then
   RsTxn_CxTF_ClientDetail.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_CxTF_ClientDetail.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_ClientDetail.MoveLast
Call InData
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
Frame0.Height = Me.Height - 350
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_CxTF_ClientDetail.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsTxn_CxTF_ClientDetail.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CxTF_ClientDetail.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsTxn_CxTF_ClientDetail.Fields(C)), "", RsTxn_CxTF_ClientDetail.Fields(C))
    Next
    RsTxn_CxTF_ClientDetail.MoveNext
Next
End Sub
Public Sub InData()

TxtClientDetailID.Locked = True
TxtCDTFNo.Locked = True
CmbCDTFNo.Visible = False
TxtCDTFNo.Visible = True
TxtClientID.Locked = True
CmbClientID.Visible = False
TxtClientID.Visible = True
TxtFlag.Locked = True

TxtClientDetailID = IIf(IsNull(RsTxn_CxTF_ClientDetail!ClientDetailID), "", RsTxn_CxTF_ClientDetail!ClientDetailID)
TxtCDTFNo = (IIf(IsNull(RsTxn_CxTF_ClientDetail!CDTFNo), "", RsTxn_CxTF_ClientDetail!CDTFNo))
mCDTFNo = IIf(IsNull(RsTxn_CxTF_ClientDetail!CDTFNo), "", RsTxn_CxTF_ClientDetail!CDTFNo)
TxtClientID = GetClientName(IIf(IsNull(RsTxn_CxTF_ClientDetail!ClientID), "", RsTxn_CxTF_ClientDetail!ClientID))
mClientID = IIf(IsNull(RsTxn_CxTF_ClientDetail!ClientID), "", RsTxn_CxTF_ClientDetail!ClientID)
TxtFlag = IIf(IsNull(RsTxn_CxTF_ClientDetail!Flag), "", RsTxn_CxTF_ClientDetail!Flag)
If RsTxn_CxTF_ClientDetail.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtClientDetailID_LostFocus()
If TxtClientDetailID = "" Then Exit Sub
If IsNumeric(TxtClientDetailID) = False Then
   MsgBox "Invalid number format !!!!"
   TxtClientDetailID.SetFocus
   Exit Sub
End If
TxtClientDetailID = Val(TxtClientDetailID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtClientDetailID.Locked = True
TxtCDTFNo.Locked = False
CmbCDTFNo.Visible = True
TxtCDTFNo.Visible = False
TxtClientID.Locked = False
CmbClientID.Visible = True
TxtClientID.Visible = False
TxtFlag.Locked = False

TxtClientDetailID = ""

TxtCDTFNo = ""

TxtClientID = ""

TxtFlag = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtClientDetailID.Locked = True
TxtCDTFNo.Locked = False
CmbCDTFNo.Visible = True
TxtCDTFNo.Visible = False
TxtClientID.Locked = False
CmbClientID.Visible = True
TxtClientID.Visible = False
TxtFlag.Locked = False

CmbCDTFNo.Text = TxtCDTFNo
CmbClientID.Text = TxtClientID
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsTxn_CxTF_ClientDetail.Delete
   RsTxn_CxTF_ClientDetail.Update
   If RsTxn_CxTF_ClientDetail.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_ClientDetail.MoveNext
   If RsTxn_CxTF_ClientDetail.EOF() Then
      RsTxn_CxTF_ClientDetail.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CxTF_ClientDetail.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If RsTxn_CxTF_ClientDetail.RecordCount > 0 Then
      RsTxn_CxTF_ClientDetail.MoveFirst
      RsTxn_CxTF_ClientDetail.Find "= '" & Txt.Text & "'"
      If Not RsTxn_CxTF_ClientDetail.EOF Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsTxn_CxTF_ClientDetail.AddNew
      RsTxn_CxTF_ClientDetail!ClientDetailID = GetClientDetailID
      TxtClientDetailID.Text = RsTxn_CxTF_ClientDetail!ClientDetailID
Else
    RsTxn_CxTF_ClientDetail.MoveFirst
    RsTxn_CxTF_ClientDetail.Find "= '" & Txt.Text & "'"
    If Not RsTxn_CxTF_ClientDetail.EOF Then
       If RsTxn_CxTF_ClientDetail!ClientDetailID <> TxtClientDetailID.Text Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsTxn_CxTF_ClientDetail.MoveFirst
    RsTxn_CxTF_ClientDetail.Find " ClientDetailID= " & TxtClientDetailID.Text
End If

RsTxn_CxTF_ClientDetail!ClientDetailID = TxtClientDetailID.Text
RsTxn_CxTF_ClientDetail!CDTFNo = mCDTFNo
RsTxn_CxTF_ClientDetail!ClientID = mClientID
RsTxn_CxTF_ClientDetail!Flag = TxtFlag.Text
RsTxn_CxTF_ClientDetail.Update
Call InData
End Sub
Private Function GetClientDetailID() As Double
Dim RetVal As Double
tmp = "Select max(ClientDetailID) from Txn_CxTF_ClientDetail"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetClientDetailID = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4
MSHFlexGrid1.TextMatrix(0, 0) = "ClientDetail ID"
MSHFlexGrid1.TextMatrix(0, 1) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 2) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 3) = "Flag"
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_ClientDetail.MoveFirst
RsTxn_CxTF_ClientDetail.Find "ClientDetailID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsTxn_CxTF_ClientDetail.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdCDTFNo_Click()
    Frm_Txn_CxTF_Details.Show
End Sub
'----
Private Sub CmbCDTFNo_GotFocus()
tmp = CmbCDTFNo.Text
Call TableTxn_CxTF_Details
CmbCDTFNo.Clear
If RsTxn_CxTF_Details.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsTxn_CxTF_Details.RecordCount
    CmbCDTFNo.AddItem RsTxn_CxTF_Details!CDTFNo
    RsTxn_CxTF_Details.MoveNext
Next
CmbCDTFNo.Text = tmp
End Sub
'----
Private Sub CmbCDTFNo_LostFocus()
If CmbCDTFNo.Text = "" Then
   Exit Sub
End If
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "CDTFNo = '" & CmbCDTFNo.Text & "'"
If RsTxn_CxTF_Details.EOF Then
   MsgBox "Invalid  selection "
   CmbCDTFNo.SetFocus
   Exit Sub
End If
mCDTFNo = RsTxn_CxTF_Details!CDTFNo
End Sub
'----
Private Sub CmdClientID_Click()
    Frm_Mst_Client.Show
End Sub
'----
Private Sub CmbClientID_GotFocus()
tmp = CmbClientID.Text
Call TableMst_Client
CmbClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbClientID.AddItem RsMst_Client!ClientID
    RsMst_Client.MoveNext
Next
CmbClientID.Text = tmp
End Sub
'----
Private Sub CmbClientID_LostFocus()
If CmbClientID.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientID = '" & CmbClientID.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid  selection "
   CmbClientID.SetFocus
   Exit Sub
End If
mClientID = RsMst_Client!ClientID
End Sub
'----
