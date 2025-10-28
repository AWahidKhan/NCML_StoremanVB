Version 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CxTF_TTD
   Caption = "Txn_CxTF_TTD"
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
      Begin VB.Label LblTTDID         Caption = "TTD ID"         Height = 300         Left = 240         TabIndex = 483         Top = 240         Width = 1800      End      Begin VB.TextBox TxtTTDID         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 484         Text = ""         Top = 240         Width = 3600      End      Begin VB.Label LblCxTFNo         Caption = "CxTFNo"         Height = 300         Left = 240         TabIndex = 485         Top = 615         Width = 1800      End      Begin VB.TextBox TxtCxTFNo         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 486         Text = ""         Top = 615         Width = 3600      End         Begin VB.ComboBox CmbCxTFNo            Height = 285            Left = 1800            TabIndex = 486            Text = ""            Top = 615            Width = 4000         End      Begin VB.Label LblTruckNo         Caption = "Truck No"         Height = 300         Left = 240         TabIndex = 489         Top = 990         Width = 1800      End      Begin VB.TextBox TxtTruckNo         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 25         TabIndex = 490         Text = ""         Top = 990         Width = 3600      End
      End
Begin VB.Frame ButtonFrm
      Height          =   1080
      Left            =   120
      TabIndex        =   491
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
Attribute VB_Name = "FrmTxn_CxTF_TTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant 
Dim mCxTFNo As Variant 
Private Sub Form_Load()
Call TableTxn_CxTF_TTD
If RsTxn_CxTF_TTD.RecordCount = 0 Then 
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
Call Gen_Create_Xls("Txn_CxTF_TTD.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Txn_CxTF_TTD.xls")
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
RsTxn_CxTF_TTD.MoveNext
If RsTxn_CxTF_TTD.EOF Then
   RsTxn_CxTF_TTD.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_TTD.MovePrevious
If RsTxn_CxTF_TTD.BOF Then
   RsTxn_CxTF_TTD.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_CxTF_TTD.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_TTD.MoveLast
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
RsTxn_CxTF_TTD.MoveFirst
Call Grid_Head 
MSHFlexGrid1.Rows = RsTxn_CxTF_TTD.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CxTF_TTD.RecordCount
    For C = 0 To MSHFlexGrid1.Cols-1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsTxn_CxTF_TTD.Fields(C)), "",  RsTxn_CxTF_TTD.Fields(C))
    Next
    RsTxn_CxTF_TTD.MoveNext
Next
End Sub
Public Sub InData()
TxtTTDID.Locked = TrueTxtCxTFNo.Locked = TrueCmbCxTFNo.Visible = False TxtCxTFNo.Visible = True TxtTruckNo.Locked = True
TxtTTDID= IIf(IsNull( RsTxn_CxTF_TTD!TTDID), "", RsTxn_CxTF_TTD!TTDID)TxtCxTFNo= GetCxTFNo(IIf(IsNull( RsTxn_CxTF_TTD!CxTFNo), "", RsTxn_CxTF_TTD!CxTFNo))mCxTFNo= IIf(IsNull( RsTxn_CxTF_TTD!CxTFNo), "", RsTxn_CxTF_TTD!CxTFNo)TxtTruckNo= IIf(IsNull( RsTxn_CxTF_TTD!TruckNo), "", RsTxn_CxTF_TTD!TruckNo)
If RsTxn_CxTF_TTD.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub
Private Sub TxtTTDID_LostFocus()If TxtTTDID = "" Then Exit SubIf IsNumeric(TxtTTDID) = False Then   MsgBox "Invalid number format !!!!"   TxtTTDID.SetFocus   Exit SubEnd IfTxtTTDID = Val(TxtTTDID)End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtTTDID.Locked = TrueTxtCxTFNo.Locked = FalseCmbCxTFNo.Visible = true TxtCxTFNo.Visible = false TxtTruckNo.Locked = False
TxtTTDID= ""TxtCxTFNo= ""TxtTruckNo= ""
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtTTDID.Locked = TrueTxtCxTFNo.Locked = FalseCmbCxTFNo.Visible = true TxtCxTFNo.Visible = false TxtTruckNo.Locked = False
CmbCxTFNo.Text = TxtCxTFNo
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsTxn_CxTF_TTD.Delete
   RsTxn_CxTF_TTD.Update
   If RsTxn_CxTF_TTD.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_TTD.MoveNext
   If RsTxn_CxTF_TTD.EOF() Then
      RsTxn_CxTF_TTD.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CxTF_TTD.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtTruckNo.Text = "" Then    Msgbox"Blank TruckNo Not Allowed !!! "   TxtTruckNo.SetFocus    Exit SubEnd If
If Edit_Flg = "A" Then
   If RsTxn_CxTF_TTD.RecordCount > 0 Then
      RsTxn_CxTF_TTD.MoveFirst
      RsTxn_CxTF_TTD.Find "= '"&Txt.Text &"'"
      if Not RsTxn_CxTF_TTD.eof Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsTxn_CxTF_TTD.AddNew
      RsTxn_CxTF_TTD!TTDID = GetTTDID
      TxtTTDID.text  = RsTxn_CxTF_TTD!TTDID
Else
    RsTxn_CxTF_TTD.MoveFirst
    RsTxn_CxTF_TTD.Find "= '"&Txt.Text &"'"
    If Not RsTxn_CxTF_TTD.EOF Then
       If RsTxn_CxTF_TTD!TTDID  <> TxtTTDID.text  Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsTxn_CxTF_TTD.MoveFirst
    RsTxn_CxTF_TTD.Find " TTDID= " &TxtTTDID.Text 
End If
RsTxn_CxTF_TTD!TTDID= TxtTTDID.Text RsTxn_CxTF_TTD!CxTFNo= mCxTFNoRsTxn_CxTF_TTD!TruckNo= TxtTruckNo.Text 
RsTxn_CxTF_TTD.update
Call InData
End Sub
Private Function GetTTDID() As Double
Dim RetVal As Double
tmp = "Select max(TTDID) from Txn_CxTF_TTD"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetTTDID = RetVal
End Function
Private sub Grid_Head
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0,0) = "TTD ID"
MSHFlexGrid1.TextMatrix(0,1) = "CxTFNo"
MSHFlexGrid1.TextMatrix(0,2) = "Truck No"
End sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_TTD.MoveFirst
RsTxn_CxTF_TTD.Find "TTDID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsTxn_CxTF_TTD.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub  CmdCxTFNo_Click()
    Frm_Txn_CxTF_Details.Show
End Sub
'---- 
Private Sub  CmbCxTFNo_GotFocus()
tmp = CmbCxTFNo.Text
Call TableTxn_CxTF_Details
CmbCxTFNo.Clear
If RsTxn_CxTF_Details.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsTxn_CxTF_Details.RecordCount
    CmbCxTFNo.AddItem RsTxn_CxTF_Details!CxTFNo
    RsTxn_CxTF_Details.MoveNext
Next
CmbCxTFNo.Text = tmp
End Sub
'---- 
Private Sub CmbCxTFNo_LostFocus()
If CmbCxTFNo.Text = "" Then
   Exit Sub
End If
RsTxn_CxTF_Details.MoveFirst
RsTxn_CxTF_Details.Find "CxTFNo = '" & CmbCxTFNo.Text & "'"
If RsTxn_CxTF_Details.EOF Then
   MsgBox "Invalid  selection "
   CmbCxTFNo.SetFocus
   Exit Sub
End If
mCxTFNo = RsTxn_CxTF_Details!CxTFNo
End Sub
'---- 
