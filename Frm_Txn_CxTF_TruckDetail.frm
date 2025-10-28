Version 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CxTF_TruckDetail
   Caption = "Txn_CxTF_TruckDetail"
   ClientHeight = 8865
   ClientLeft = 60
   ClientTop = 450
   ClientWidth = 12500
   MDIChild = -1           'True
   LinkTopic = ""
   ScaleHeight = 4920
   ScaleWidth = 12500
   StartUpPosition = 2    'CenterScreen
   WindowState = 2        'Maximized
   Begin VB.Frame Frame0
         Caption = ""
         Height = 4000
         Left = 8000
         TabIndex = 0
         Top = 120
         Width = 3235
         Visible = False
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   80
            TabIndex        =   0
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
               WordWrap = -1           'True
               AllowUserResizing = 3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
         End
   End
   Begin MSComctlLib.StatusBar StatusBar1
         Align = 2              'Align Bottom
         Height = 375
         Left = 0
         TabIndex = 0
         Top = 4545
         Width = 11055
         _ExtentX        =   19500
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628}
            NumPanels = 1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628}
            EndProperty
         EndProperty
      End
   Begin VB.Frame Frame1
         Caption = ""
         Height = 5865
         Left = 120
         TabIndex = 0
         Top = 120
         Width = 10000
      Begin VB.Label LblTruckDetailID         Caption = "TruckDetail ID"         Height = 300         Left = 240         TabIndex = 462         Top = 240         Width = 1800      End      Begin VB.TextBox TxtTruckDetailID         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 463         Text = ""         Top = 240         Width = 3600      End      Begin VB.Label LblCDTFNo         Caption = "CDTF No"         Height = 300         Left = 240         TabIndex = 464         Top = 615         Width = 1800      End      Begin VB.TextBox TxtCDTFNo         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 465         Text = ""         Top = 615         Width = 3600      End         Begin VB.ComboBox CmbCDTFNo            Height = 285            Left = 1800            TabIndex = 465            Text = ""            Top = 615            Width = 4000         End      Begin VB.Label LblTruckNo         Caption = "Truck No"         Height = 300         Left = 240         TabIndex = 468         Top = 990         Width = 1800      End      Begin VB.TextBox TxtTruckNo         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 50         TabIndex = 469         Text = ""         Top = 990         Width = 3600      End      Begin VB.Label LblWeighBridgeSlipno         Caption = "WeighBridge Slip No"         Height = 300         Left = 240         TabIndex = 470         Top = 1365         Width = 1800      End      Begin VB.TextBox TxtWeighBridgeSlipno         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 50         TabIndex = 471         Text = ""         Top = 1365         Width = 3600      End      Begin VB.Label LblTotalWeight         Caption = "Total Weight"         Height = 300         Left = 240         TabIndex = 472         Top = 1740         Width = 1800      End      Begin VB.TextBox TxtTotalWeight         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 473         Text = ""         Top = 1740         Width = 3600      End      Begin VB.Label LblTareWeight         Caption = "Tare Weight"         Height = 300         Left = 240         TabIndex = 474         Top = 2115         Width = 1800      End      Begin VB.TextBox TxtTareWeight         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 475         Text = ""         Top = 2115         Width = 3600      End      Begin VB.Label LblGrossWeight         Caption = "Gross Weight"         Height = 300         Left = 240         TabIndex = 476         Top = 2490         Width = 1800      End      Begin VB.TextBox TxtGrossWeight         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 477         Text = ""         Top = 2490         Width = 3600      End      Begin VB.Label LblRemarks         Caption = "Remarks"         Height = 300         Left = 240         TabIndex = 478         Top = 2865         Width = 1800      End      Begin VB.TextBox TxtRemarks         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 255         TabIndex = 479         Text = ""         Top = 2865         Width = 3600      End      Begin VB.Label LblFlag         Caption = "Flag"         Height = 300         Left = 240         TabIndex = 480         Top = 3240         Width = 1800      End      Begin VB.TextBox TxtFlag         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 1         TabIndex = 481         Text = ""         Top = 3240         Width = 3600      End
      End
Begin VB.Frame ButtonFrm
      Height          =   1080
      Left            =   120
      TabIndex        =   482
      Top             =   7200
      Width           =   10000
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7320
            TabIndex        =   
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7320
            TabIndex        =   1
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   5520
            TabIndex        =   2
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   5520
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3720
            TabIndex        =   4
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3720
            TabIndex        =   5
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1920
            TabIndex        =   6
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1920
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   100
            TabIndex        =   8
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   100
            TabIndex        =   9
            Top             =   240
            Width           =   1815
         End
   End
End
Attribute VB_Name = "FrmTxn_CxTF_TruckDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant 
Dim mCDTFNo As Variant 
Private Sub Form_Load()
Call TableTxn_CxTF_TruckDetail
If RsTxn_CxTF_TruckDetail.RecordCount = 0 Then 
   If UserType = "Auditor"Then
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
Call Gen_Create_Xls("Txn_CxTF_TruckDetail.xls")
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
Pat = App.path
Set oWB = oXL.Workbooks.Open(Pat & "\Txn_CxTF_TruckDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C),"/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C),"MM/dd/yyyy")
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
RsTxn_CxTF_TruckDetail.MoveNext
If RsTxn_CxTF_TruckDetail.EOF Then
   RsTxn_CxTF_TruckDetail.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_TruckDetail.MovePrevious
If RsTxn_CxTF_TruckDetail.BOF Then
   RsTxn_CxTF_TruckDetail.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_CxTF_TruckDetail.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_TruckDetail.MoveLast
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
Frame0.Height = me.Height - 350
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100 
Frame0.Width = me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_CxTF_TruckDetail.MoveFirst
Call Grid_Head 
MSHFlexGrid1.Rows = RsTxn_CxTF_TruckDetail.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CxTF_TruckDetail.RecordCount
    For C = 0 To MSHFlexGrid1.Cols-1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsTxn_CxTF_TruckDetail.Fields(C)), "",  RsTxn_CxTF_TruckDetail.Fields(C))
    Next
    RsTxn_CxTF_TruckDetail.MoveNext
Next
End Sub
Public Sub InData()
TxtTruckDetailID.Locked = TrueTxtCDTFNo.Locked = TrueCmbCDTFNo.Visible = False TxtCDTFNo.Visible = True TxtTruckNo.Locked = TrueTxtWeighBridgeSlipno.Locked = TrueTxtTotalWeight.Locked = TrueTxtTareWeight.Locked = TrueTxtGrossWeight.Locked = TrueTxtRemarks.Locked = TrueTxtFlag.Locked = True
TxtTruckDetailID= IIf(IsNull( RsTxn_CxTF_TruckDetail!TruckDetailID), "", RsTxn_CxTF_TruckDetail!TruckDetailID)TxtCDTFNo= GetCDTFNo(IIf(IsNull( RsTxn_CxTF_TruckDetail!CDTFNo), "", RsTxn_CxTF_TruckDetail!CDTFNo))mCDTFNo= IIf(IsNull( RsTxn_CxTF_TruckDetail!CDTFNo), "", RsTxn_CxTF_TruckDetail!CDTFNo)TxtTruckNo= IIf(IsNull( RsTxn_CxTF_TruckDetail!TruckNo), "", RsTxn_CxTF_TruckDetail!TruckNo)TxtWeighBridgeSlipno= IIf(IsNull( RsTxn_CxTF_TruckDetail!WeighBridgeSlipno), "", RsTxn_CxTF_TruckDetail!WeighBridgeSlipno)TxtTotalWeight= IIf(IsNull( RsTxn_CxTF_TruckDetail!TotalWeight), "", RsTxn_CxTF_TruckDetail!TotalWeight)TxtTareWeight= IIf(IsNull( RsTxn_CxTF_TruckDetail!TareWeight), "", RsTxn_CxTF_TruckDetail!TareWeight)TxtGrossWeight= IIf(IsNull( RsTxn_CxTF_TruckDetail!GrossWeight), "", RsTxn_CxTF_TruckDetail!GrossWeight)TxtRemarks= IIf(IsNull( RsTxn_CxTF_TruckDetail!Remarks), "", RsTxn_CxTF_TruckDetail!Remarks)TxtFlag= IIf(IsNull( RsTxn_CxTF_TruckDetail!Flag), "", RsTxn_CxTF_TruckDetail!Flag)
If RsTxn_CxTF_TruckDetail.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub
Private Sub TxtTruckDetailID_LostFocus()If TxtTruckDetailID = "" Then Exit SubIf IsNumeric(TxtTruckDetailID) = False Then   MsgBox "Invalid number format !!!!"   TxtTruckDetailID.SetFocus   Exit SubEnd IfTxtTruckDetailID = Val(TxtTruckDetailID)End SubPrivate Sub TxtTotalWeight_LostFocus()If TxtTotalWeight = "" Then Exit SubIf IsNumeric(TxtTotalWeight) = False Then   MsgBox "Invalid number format !!!!"   TxtTotalWeight.SetFocus   Exit SubEnd IfTxtTotalWeight = Val(TxtTotalWeight)End SubPrivate Sub TxtTareWeight_LostFocus()If TxtTareWeight = "" Then Exit SubIf IsNumeric(TxtTareWeight) = False Then   MsgBox "Invalid number format !!!!"   TxtTareWeight.SetFocus   Exit SubEnd IfTxtTareWeight = Val(TxtTareWeight)End SubPrivate Sub TxtGrossWeight_LostFocus()If TxtGrossWeight = "" Then Exit SubIf IsNumeric(TxtGrossWeight) = False Then   MsgBox "Invalid number format !!!!"   TxtGrossWeight.SetFocus   Exit SubEnd IfTxtGrossWeight = Val(TxtGrossWeight)End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtTruckDetailID.Locked = TrueTxtCDTFNo.Locked = FalseCmbCDTFNo.Visible = true TxtCDTFNo.Visible = false TxtTruckNo.Locked = FalseTxtWeighBridgeSlipno.Locked = FalseTxtTotalWeight.Locked = FalseTxtTareWeight.Locked = FalseTxtGrossWeight.Locked = FalseTxtRemarks.Locked = FalseTxtFlag.Locked = False
TxtTruckDetailID= ""TxtCDTFNo= ""TxtTruckNo= ""TxtWeighBridgeSlipno= ""TxtTotalWeight= ""TxtTareWeight= ""TxtGrossWeight= ""TxtRemarks= ""TxtFlag= ""
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtTruckDetailID.Locked = TrueTxtCDTFNo.Locked = FalseCmbCDTFNo.Visible = true TxtCDTFNo.Visible = false TxtTruckNo.Locked = FalseTxtWeighBridgeSlipno.Locked = FalseTxtTotalWeight.Locked = FalseTxtTareWeight.Locked = FalseTxtGrossWeight.Locked = FalseTxtRemarks.Locked = FalseTxtFlag.Locked = False
CmbCDTFNo.Text = TxtCDTFNo
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsTxn_CxTF_TruckDetail.Delete
   RsTxn_CxTF_TruckDetail.Update
   If RsTxn_CxTF_TruckDetail.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_TruckDetail.MoveNext
   If RsTxn_CxTF_TruckDetail.EOF() Then
      RsTxn_CxTF_TruckDetail.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CxTF_TruckDetail.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If RsTxn_CxTF_TruckDetail.RecordCount > 0 Then
      RsTxn_CxTF_TruckDetail.MoveFirst
      RsTxn_CxTF_TruckDetail.Find "= '"&Txt.Text &"'"
      if Not RsTxn_CxTF_TruckDetail.eof Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsTxn_CxTF_TruckDetail.AddNew
      RsTxn_CxTF_TruckDetail!TruckDetailID = GetTruckDetailID
      TxtTruckDetailID.text  = RsTxn_CxTF_TruckDetail!TruckDetailID
Else
    RsTxn_CxTF_TruckDetail.MoveFirst
    RsTxn_CxTF_TruckDetail.Find "= '"&Txt.Text &"'"
    If Not RsTxn_CxTF_TruckDetail.EOF Then
       If RsTxn_CxTF_TruckDetail!TruckDetailID  <> TxtTruckDetailID.text  Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsTxn_CxTF_TruckDetail.MoveFirst
    RsTxn_CxTF_TruckDetail.Find " TruckDetailID= " &TxtTruckDetailID.Text 
End If
RsTxn_CxTF_TruckDetail!TruckDetailID= TxtTruckDetailID.Text RsTxn_CxTF_TruckDetail!CDTFNo= mCDTFNoRsTxn_CxTF_TruckDetail!TruckNo= TxtTruckNo.Text RsTxn_CxTF_TruckDetail!WeighBridgeSlipno= TxtWeighBridgeSlipno.Text RsTxn_CxTF_TruckDetail!TotalWeight= TxtTotalWeight.Text RsTxn_CxTF_TruckDetail!TareWeight= TxtTareWeight.Text RsTxn_CxTF_TruckDetail!GrossWeight= TxtGrossWeight.Text RsTxn_CxTF_TruckDetail!Remarks= TxtRemarks.Text RsTxn_CxTF_TruckDetail!Flag= TxtFlag.Text 
RsTxn_CxTF_TruckDetail.update
Call InData
End Sub
Private Function GetTruckDetailID() As Double
Dim RetVal As Double
tmp = "Select max(TruckDetailID) from Txn_CxTF_TruckDetail"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetTruckDetailID = RetVal
End Function
Private sub Grid_Head
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9
MSHFlexGrid1.TextMatrix(0,0) = "TruckDetail ID"
MSHFlexGrid1.TextMatrix(0,1) = "CDTF No"
MSHFlexGrid1.TextMatrix(0,2) = "Truck No"
MSHFlexGrid1.TextMatrix(0,3) = "WeighBridge Slip No"
MSHFlexGrid1.TextMatrix(0,4) = "Total Weight"
MSHFlexGrid1.TextMatrix(0,5) = "Tare Weight"
MSHFlexGrid1.TextMatrix(0,6) = "Gross Weight"
MSHFlexGrid1.TextMatrix(0,7) = "Remarks"
MSHFlexGrid1.TextMatrix(0,8) = "Flag"
End sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_TruckDetail.MoveFirst
RsTxn_CxTF_TruckDetail.Find "TruckDetailID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsTxn_CxTF_TruckDetail.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub  CmdCDTFNo_Click()
    Frm_Txn_CxTF_Details.Show
End Sub
'---- 
Private Sub  CmbCDTFNo_GotFocus()
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
