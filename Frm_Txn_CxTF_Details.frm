VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CxTF_Details 
   Caption         =   "Txn_CxTF_Details"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12495
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2445
      Left            =   10080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3480
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
      Height          =   4545
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9885
      Begin VB.TextBox TxtCxTFNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker TxtCxTFDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Top             =   615
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   48758785
         CurrentDate     =   36494
      End
      Begin VB.TextBox TxtCMPID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   18
         Top             =   990
         Width           =   3600
      End
      Begin VB.ComboBox CmbCMPID 
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   990
         Width           =   4000
      End
      Begin VB.TextBox TxtCommodityID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   21
         Top             =   1365
         Width           =   3600
      End
      Begin VB.ComboBox CmbCommodityID 
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   1365
         Width           =   4000
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   24
         Top             =   1740
         Width           =   3600
      End
      Begin VB.ComboBox CmbExchangeTypeID 
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Top             =   1740
         Width           =   4000
      End
      Begin VB.TextBox TxtBookedSpaceFlag 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   27
         Top             =   2115
         Width           =   1080
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   29
         Top             =   2490
         Width           =   3600
      End
      Begin VB.TextBox TxtAgent 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   31
         Top             =   2865
         Width           =   3600
      End
      Begin VB.TextBox TxtRemark 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   33
         Top             =   3240
         Width           =   3600
      End
      Begin VB.TextBox TxtTxnType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   35
         Top             =   3615
         Width           =   3600
      End
      Begin VB.TextBox TxtContractID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   37
         Top             =   3990
         Width           =   3600
      End
      Begin VB.ComboBox CmbContractID 
         Height          =   285
         Left            =   1800
         TabIndex        =   38
         Top             =   3990
         Width           =   4000
      End
      Begin VB.Label LblCxTFNo 
         Caption         =   "CDTF No"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblCxTFDate 
         Caption         =   "CDTF Date"
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblCMPID 
         Caption         =   "CMP"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label LblCommodityID 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   1365
         Width           =   1800
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "Exchange Type"
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   1740
         Width           =   1800
      End
      Begin VB.Label LblBookedSpaceFlag 
         Caption         =   "a"
         Height          =   300
         Left            =   240
         TabIndex        =   26
         Top             =   2115
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   2490
         Width           =   1800
      End
      Begin VB.Label LblAgent 
         Caption         =   "Agent"
         Height          =   300
         Left            =   240
         TabIndex        =   30
         Top             =   2865
         Width           =   1800
      End
      Begin VB.Label LblRemark 
         Caption         =   "Remark"
         Height          =   300
         Left            =   240
         TabIndex        =   32
         Top             =   3240
         Width           =   1800
      End
      Begin VB.Label LblTxnType 
         Caption         =   "Txn Type"
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   3615
         Width           =   1800
      End
      Begin VB.Label LblContractID 
         Caption         =   "Booked Space ID"
         Height          =   300
         Left            =   240
         TabIndex        =   36
         Top             =   3990
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   39
      Top             =   7200
      Width           =   10000
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   7320
         TabIndex        =   40
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   7320
         TabIndex        =   41
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
Attribute VB_Name = "FrmTxn_CxTF_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mCMPID, mCommodityID, mExchangeTypeID, mContractID As Variant
Private Sub Form_Load()
Call TableTxn_CxTF_Details
If RsTxn_CxTF_Details.RecordCount = 0 Then
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
Call Gen_Create_Xls("Txn_CxTF_Details.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Txn_CxTF_Details.xls")
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
RsTxn_CxTF_Details.MoveNext
If RsTxn_CxTF_Details.EOF Then
   RsTxn_CxTF_Details.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_Details.MovePrevious
If RsTxn_CxTF_Details.BOF Then
   RsTxn_CxTF_Details.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_CxTF_Details.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_Details.MoveLast
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
RsTxn_CxTF_Details.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsTxn_CxTF_Details.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CxTF_Details.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsTxn_CxTF_Details.Fields(C)), "", RsTxn_CxTF_Details.Fields(C))
    Next
    RsTxn_CxTF_Details.MoveNext
Next
End Sub
Public Sub InData()

TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = False
TxtCMPID.Locked = True
CmbCMPID.Visible = False
TxtCMPID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
TxtCommodityID.Visible = True
TxtExchangeTypeID.Locked = True
CmbExchangeTypeID.Visible = False
TxtExchangeTypeID.Visible = True
TxtBookedSpaceFlag.Locked = True
TxtFlag.Locked = True
TxtAgent.Locked = True
TxtRemark.Locked = True
TxtTxnType.Locked = True
TxtContractID.Locked = True
CmbContractID.Visible = False
TxtContractID.Visible = True

TxtCxTFNo = IIf(IsNull(RsTxn_CxTF_Details!CxTFNo), "", RsTxn_CxTF_Details!CxTFNo)
TxtCxTFDate = IIf(IsNull(RsTxn_CxTF_Details!CxTFDate), "", RsTxn_CxTF_Details!CxTFDate)
TxtCMPID = GetCMPName(IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID))
mCMPID = IIf(IsNull(RsTxn_CxTF_Details!CMPID), "", RsTxn_CxTF_Details!CMPID)
TxtCommodityID = GetCommodity(IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID))
mCommodityID = IIf(IsNull(RsTxn_CxTF_Details!CommodityID), "", RsTxn_CxTF_Details!CommodityID)
TxtExchangeTypeID = GetExchangeType(IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID))
mExchangeTypeID = IIf(IsNull(RsTxn_CxTF_Details!ExchangeTypeID), "", RsTxn_CxTF_Details!ExchangeTypeID)
TxtBookedSpaceFlag = IIf(IsNull(RsTxn_CxTF_Details!BookedSpaceFlag), "", RsTxn_CxTF_Details!BookedSpaceFlag)
TxtFlag = IIf(IsNull(RsTxn_CxTF_Details!Flag), "", RsTxn_CxTF_Details!Flag)
TxtAgent = IIf(IsNull(RsTxn_CxTF_Details!Agent), "", RsTxn_CxTF_Details!Agent)
TxtRemark = IIf(IsNull(RsTxn_CxTF_Details!Remark), "", RsTxn_CxTF_Details!Remark)
TxtTxnType = IIf(IsNull(RsTxn_CxTF_Details!TxnType), "", RsTxn_CxTF_Details!TxnType)
TxtContractID = GetContractID(IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID))
mContractID = IIf(IsNull(RsTxn_CxTF_Details!ContractID), "", RsTxn_CxTF_Details!ContractID)
If RsTxn_CxTF_Details.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtCxTFNo_LostFocus()
If TxtCxTFNo = "" Then Exit Sub
If IsNumeric(TxtCxTFNo) = False Then
   MsgBox "Invalid number format !!!!"
   TxtCxTFNo.SetFocus
   Exit Sub
End If
TxtCxTFNo = Val(TxtCxTFNo)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = True
TxtCMPID.Locked = False
CmbCMPID.Visible = True
TxtCMPID.Visible = False
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = False
CmbExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False
TxtBookedSpaceFlag.Locked = False
TxtFlag.Locked = False
TxtAgent.Locked = False
TxtRemark.Locked = False
TxtTxnType.Locked = False
TxtContractID.Locked = False
CmbContractID.Visible = True
TxtContractID.Visible = False

TxtCxTFNo = ""

TxtCxTFDate = ""

TxtCMPID = ""

TxtCommodityID = ""

TxtExchangeTypeID = ""

TxtBookedSpaceFlag = ""

TxtFlag = ""

TxtAgent = ""

TxtRemark = ""

TxtTxnType = ""

TxtContractID = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtCxTFNo.Locked = True
TxtCxTFDate.Enabled = True
TxtCMPID.Locked = False
CmbCMPID.Visible = True
TxtCMPID.Visible = False
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = False
CmbExchangeTypeID.Visible = True
TxtExchangeTypeID.Visible = False
TxtBookedSpaceFlag.Locked = False
TxtFlag.Locked = False
TxtAgent.Locked = False
TxtRemark.Locked = False
TxtTxnType.Locked = False
TxtContractID.Locked = False
CmbContractID.Visible = True
TxtContractID.Visible = False

CmbCMPID.Text = TxtCMPID
CmbCommodityID.Text = TxtCommodityID
CmbExchangeTypeID.Text = TxtExchangeTypeID
CmbContractID.Text = TxtContractID
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsTxn_CxTF_Details.Delete
   RsTxn_CxTF_Details.Update
   If RsTxn_CxTF_Details.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_Details.MoveNext
   If RsTxn_CxTF_Details.EOF() Then
      RsTxn_CxTF_Details.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CxTF_Details.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If RsTxn_CxTF_Details.RecordCount > 0 Then
      RsTxn_CxTF_Details.MoveFirst
      RsTxn_CxTF_Details.Find "= '" & Txt.Text & "'"
      If Not RsTxn_CxTF_Details.EOF Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsTxn_CxTF_Details.AddNew
      RsTxn_CxTF_Details!CxTFNo = GetCxTFNo
      TxtCxTFNo.Text = RsTxn_CxTF_Details!CxTFNo
Else
    RsTxn_CxTF_Details.MoveFirst
    RsTxn_CxTF_Details.Find "= '" & Txt.Text & "'"
    If Not RsTxn_CxTF_Details.EOF Then
       If RsTxn_CxTF_Details!CxTFNo <> TxtCxTFNo.Text Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsTxn_CxTF_Details.MoveFirst
    RsTxn_CxTF_Details.Find " CxTFNo= " & TxtCxTFNo.Text
End If

RsTxn_CxTF_Details!CxTFNo = TxtCxTFNo.Text
RsTxn_CxTF_Details!CxTFDate = TxtCxTFDate.Text
RsTxn_CxTF_Details!CMPID = mCMPID
RsTxn_CxTF_Details!CommodityID = mCommodityID
RsTxn_CxTF_Details!ExchangeTypeID = mExchangeTypeID
RsTxn_CxTF_Details!BookedSpaceFlag = TxtBookedSpaceFlag.Text
RsTxn_CxTF_Details!Flag = TxtFlag.Text
RsTxn_CxTF_Details!Agent = TxtAgent.Text
RsTxn_CxTF_Details!Remark = TxtRemark.Text
RsTxn_CxTF_Details!TxnType = TxtTxnType.Text
RsTxn_CxTF_Details!ContractID = mContractID
RsTxn_CxTF_Details.Update
Call InData
End Sub
Private Function GetCxTFNo() As Double
Dim RetVal As Double
tmp = "Select max(CxTFNo) from Txn_CxTF_Details"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetCxTFNo = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.TextMatrix(0, 0) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 1) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 5) = "a"
MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
MSHFlexGrid1.TextMatrix(0, 7) = "Agent"
MSHFlexGrid1.TextMatrix(0, 8) = "Remark"
MSHFlexGrid1.TextMatrix(0, 9) = "Txn Type"
MSHFlexGrid1.TextMatrix(0, 10) = "Booked Space ID"
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "CxTFNo = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsTxn_CxTF_Details.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdCMPID_Click()
    Frm_Mst_CMP.Show
End Sub
'----
Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text
Call TableMst_CMP
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub
'----
Private Sub CmdCommodityID_Click()
    Frm_Mst_Commodity.Show
End Sub
'----
Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
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
   MsgBox "Invalid  selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
End Sub
'----
Private Sub CmdExchangeTypeID_Click()
    Frm_Mst_ExchangeType.Show
End Sub
'----
Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid  selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub
'----
Private Sub CmdContractID_Click()
    Frm_Mst_StorageContract.Show
End Sub
'----
Private Sub CmbContractID_GotFocus()
tmp = CmbContractID.Text
Call TableMst_StorageContract
CmbContractID.Clear
If RsMst_StorageContract.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_StorageContract.RecordCount
    CmbContractID.AddItem RsMst_StorageContract!ContractID
    RsMst_StorageContract.MoveNext
Next
CmbContractID.Text = tmp
End Sub
'----
Private Sub CmbContractID_LostFocus()
If CmbContractID.Text = "" Then
   Exit Sub
End If
RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractID = '" & CmbContractID.Text & "'"
If RsMst_StorageContract.EOF Then
   MsgBox "Invalid  selection "
   CmbContractID.SetFocus
   Exit Sub
End If
mContractID = RsMst_StorageContract!ContractID
End Sub
'----
