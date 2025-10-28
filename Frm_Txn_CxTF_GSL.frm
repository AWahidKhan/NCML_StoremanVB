Version 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CxTF_GSL
   Caption = "Txn_CxTF_GSL"
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
      Begin VB.Label LblGSLDetailID         Caption = "GSLDetail ID"         Height = 300         Left = 240         TabIndex = 429         Top = 240         Width = 1800      End      Begin VB.TextBox TxtGSLDetailID         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 430         Text = ""         Top = 240         Width = 3600      End      Begin VB.Label LblCxTFNo         Caption = "CxTF No"         Height = 300         Left = 240         TabIndex = 431         Top = 615         Width = 1800      End      Begin VB.TextBox TxtCxTFNo         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 432         Text = ""         Top = 615         Width = 3600      End         Begin VB.ComboBox CmbCxTFNo            Height = 285            Left = 1800            TabIndex = 432            Text = ""            Top = 615            Width = 4000         End      Begin VB.Label LblNoOfBags         Caption = "No Of Bags"         Height = 300         Left = 240         TabIndex = 435         Top = 990         Width = 1800      End      Begin VB.TextBox TxtNoOfBags         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 10         TabIndex = 436         Text = ""         Top = 990         Width = 3600      End      Begin VB.Label LblGrossWeight         Caption = "Gross Weight"         Height = 300         Left = 240         TabIndex = 437         Top = 1365         Width = 1800      End      Begin VB.TextBox TxtGrossWeight         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 438         Text = ""         Top = 1365         Width = 3600      End      Begin VB.Label LblDematWeight         Caption = "Demat Weight"         Height = 300         Left = 240         TabIndex = 439         Top = 1740         Width = 1800      End      Begin VB.TextBox TxtDematWeight         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 440         Text = ""         Top = 1740         Width = 3600      End      Begin VB.Label LblClientID         Caption = "Client ID"         Height = 300         Left = 240         TabIndex = 441         Top = 2115         Width = 1800      End      Begin VB.TextBox TxtClientID         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 50         TabIndex = 442         Text = ""         Top = 2115         Width = 3600      End         Begin VB.ComboBox CmbClientID            Height = 285            Left = 1800            TabIndex = 442            Text = ""            Top = 2115            Width = 4000         End      Begin VB.Label LblRemarks         Caption = "Remarks"         Height = 300         Left = 240         TabIndex = 445         Top = 2490         Width = 1800      End      Begin VB.TextBox TxtRemarks         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 255         TabIndex = 446         Text = ""         Top = 2490         Width = 3600      End      Begin VB.Label LblFlag         Caption = "Flag"         Height = 300         Left = 240         TabIndex = 447         Top = 2865         Width = 1800      End      Begin VB.TextBox TxtFlag         Appearance = 0         'Flat         Height = 360         Left = 1800         MaxLength = 1         TabIndex = 448         Text = ""         Top = 2865         Width = 3600      End      Begin VB.Label LblGunnyWeight         Caption = "Gunny Weight"         Height = 300         Left = 240         TabIndex = 449         Top = 3240         Width = 1800      End      Begin VB.TextBox TxtGunnyWeight         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 450         Text = ""         Top = 3240         Width = 3600      End      Begin VB.Label LblStdDeduction         Caption = "Std Deduction"         Height = 300         Left = 240         TabIndex = 451         Top = 3615         Width = 1800      End      Begin VB.TextBox TxtStdDeduction         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 452         Text = ""         Top = 3615         Width = 3600      End      Begin VB.Label LblAdjNCMSLWeight         Caption = "Adj NCMSL Weight"         Height = 300         Left = 240         TabIndex = 453         Top = 3990         Width = 1800      End      Begin VB.TextBox TxtAdjNCMSLWeight         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 454         Text = ""         Top = 3990         Width = 3600      End      Begin VB.Label LblGSLID         Caption = "Lot No"         Height = 300         Left = 240         TabIndex = 457         Top = 4740         Width = 1800      End      Begin VB.TextBox TxtGSLID         Appearance = 0         'Flat         Height = 360         Left = 1800         TabIndex = 458         Text = ""         Top = 4740         Width = 3600      End         Begin VB.ComboBox CmbGSLID            Height = 285            Left = 1800            TabIndex = 458            Text = ""            Top = 4740            Width = 4000         End
      End
Begin VB.Frame ButtonFrm
      Height          =   1080
      Left            =   120
      TabIndex        =   461
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
Attribute VB_Name = "FrmTxn_CxTF_GSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant 
Dim mCxTFNo, mClientID, mGSLID As Variant 
Private Sub Form_Load()
Call TableTxn_CxTF_GSL
If RsTxn_CxTF_GSL.RecordCount = 0 Then 
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
Call Gen_Create_Xls("Txn_CxTF_GSL.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Txn_CxTF_GSL.xls")
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
RsTxn_CxTF_GSL.MoveNext
If RsTxn_CxTF_GSL.EOF Then
   RsTxn_CxTF_GSL.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CxTF_GSL.MovePrevious
If RsTxn_CxTF_GSL.BOF Then
   RsTxn_CxTF_GSL.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_CxTF_GSL.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CxTF_GSL.MoveLast
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
RsTxn_CxTF_GSL.MoveFirst
Call Grid_Head 
MSHFlexGrid1.Rows = RsTxn_CxTF_GSL.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CxTF_GSL.RecordCount
    For C = 0 To MSHFlexGrid1.Cols-1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsTxn_CxTF_GSL.Fields(C)), "",  RsTxn_CxTF_GSL.Fields(C))
    Next
    RsTxn_CxTF_GSL.MoveNext
Next
End Sub
Public Sub InData()
TxtGSLDetailID.Locked = TrueTxtCxTFNo.Locked = TrueCmbCxTFNo.Visible = False TxtCxTFNo.Visible = True TxtNoOfBags.Locked = TrueTxtGrossWeight.Locked = TrueTxtDematWeight.Locked = TrueTxtClientID.Locked = TrueCmbClientID.Visible = False TxtClientID.Visible = True TxtRemarks.Locked = TrueTxtFlag.Locked = TrueTxtGunnyWeight.Locked = TrueTxtStdDeduction.Locked = TrueTxtAdjNCMSLWeight.Locked = TrueTxtGSLID.Locked = TrueCmbGSLID.Visible = False TxtGSLID.Visible = True 
TxtGSLDetailID= IIf(IsNull( RsTxn_CxTF_GSL!GSLDetailID), "", RsTxn_CxTF_GSL!GSLDetailID)TxtCxTFNo= GetCxTFNo(IIf(IsNull( RsTxn_CxTF_GSL!CxTFNo), "", RsTxn_CxTF_GSL!CxTFNo))mCxTFNo= IIf(IsNull( RsTxn_CxTF_GSL!CxTFNo), "", RsTxn_CxTF_GSL!CxTFNo)TxtNoOfBags= IIf(IsNull( RsTxn_CxTF_GSL!NoOfBags), "", RsTxn_CxTF_GSL!NoOfBags)TxtGrossWeight= IIf(IsNull( RsTxn_CxTF_GSL!GrossWeight), "", RsTxn_CxTF_GSL!GrossWeight)TxtDematWeight= IIf(IsNull( RsTxn_CxTF_GSL!DematWeight), "", RsTxn_CxTF_GSL!DematWeight)TxtClientID= GetClientID(IIf(IsNull( RsTxn_CxTF_GSL!ClientID), "", RsTxn_CxTF_GSL!ClientID))mClientID= IIf(IsNull( RsTxn_CxTF_GSL!ClientID), "", RsTxn_CxTF_GSL!ClientID)TxtRemarks= IIf(IsNull( RsTxn_CxTF_GSL!Remarks), "", RsTxn_CxTF_GSL!Remarks)TxtFlag= IIf(IsNull( RsTxn_CxTF_GSL!Flag), "", RsTxn_CxTF_GSL!Flag)TxtGunnyWeight= IIf(IsNull( RsTxn_CxTF_GSL!GunnyWeight), "", RsTxn_CxTF_GSL!GunnyWeight)TxtStdDeduction= IIf(IsNull( RsTxn_CxTF_GSL!StdDeduction), "", RsTxn_CxTF_GSL!StdDeduction)TxtAdjNCMSLWeight= IIf(IsNull( RsTxn_CxTF_GSL!AdjNCMSLWeight), "", RsTxn_CxTF_GSL!AdjNCMSLWeight)TxtGSLID= GetLotNo(IIf(IsNull( RsTxn_CxTF_GSL!GSLID), "", RsTxn_CxTF_GSL!GSLID))mGSLID= IIf(IsNull( RsTxn_CxTF_GSL!GSLID), "", RsTxn_CxTF_GSL!GSLID)
If RsTxn_CxTF_GSL.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub
Private Sub TxtGSLDetailID_LostFocus()If TxtGSLDetailID = "" Then Exit SubIf IsNumeric(TxtGSLDetailID) = False Then   MsgBox "Invalid number format !!!!"   TxtGSLDetailID.SetFocus   Exit SubEnd IfTxtGSLDetailID = Val(TxtGSLDetailID)End SubPrivate Sub TxtGrossWeight_LostFocus()If TxtGrossWeight = "" Then Exit SubIf IsNumeric(TxtGrossWeight) = False Then   MsgBox "Invalid number format !!!!"   TxtGrossWeight.SetFocus   Exit SubEnd IfTxtGrossWeight = Val(TxtGrossWeight)End SubPrivate Sub TxtDematWeight_LostFocus()If TxtDematWeight = "" Then Exit SubIf IsNumeric(TxtDematWeight) = False Then   MsgBox "Invalid number format !!!!"   TxtDematWeight.SetFocus   Exit SubEnd IfTxtDematWeight = Val(TxtDematWeight)End SubPrivate Sub TxtGunnyWeight_LostFocus()If TxtGunnyWeight = "" Then Exit SubIf IsNumeric(TxtGunnyWeight) = False Then   MsgBox "Invalid number format !!!!"   TxtGunnyWeight.SetFocus   Exit SubEnd IfTxtGunnyWeight = Val(TxtGunnyWeight)End SubPrivate Sub TxtStdDeduction_LostFocus()If TxtStdDeduction = "" Then Exit SubIf IsNumeric(TxtStdDeduction) = False Then   MsgBox "Invalid number format !!!!"   TxtStdDeduction.SetFocus   Exit SubEnd IfTxtStdDeduction = Val(TxtStdDeduction)End SubPrivate Sub TxtAdjNCMSLWeight_LostFocus()If TxtAdjNCMSLWeight = "" Then Exit SubIf IsNumeric(TxtAdjNCMSLWeight) = False Then   MsgBox "Invalid number format !!!!"   TxtAdjNCMSLWeight.SetFocus   Exit SubEnd IfTxtAdjNCMSLWeight = Val(TxtAdjNCMSLWeight)End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtGSLDetailID.Locked = TrueTxtCxTFNo.Locked = FalseCmbCxTFNo.Visible = true TxtCxTFNo.Visible = false TxtNoOfBags.Locked = FalseTxtGrossWeight.Locked = FalseTxtDematWeight.Locked = FalseTxtClientID.Locked = FalseCmbClientID.Visible = true TxtClientID.Visible = false TxtRemarks.Locked = FalseTxtFlag.Locked = FalseTxtGunnyWeight.Locked = FalseTxtStdDeduction.Locked = FalseTxtAdjNCMSLWeight.Locked = FalseTxtGSLID.Locked = FalseCmbGSLID.Visible = true TxtGSLID.Visible = false 
TxtGSLDetailID= ""TxtCxTFNo= ""TxtNoOfBags= ""TxtGrossWeight= ""TxtDematWeight= ""TxtClientID= ""TxtRemarks= ""TxtFlag= ""TxtGunnyWeight= ""TxtStdDeduction= ""TxtAdjNCMSLWeight= ""TxtGSLID= ""
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtGSLDetailID.Locked = TrueTxtCxTFNo.Locked = FalseCmbCxTFNo.Visible = true TxtCxTFNo.Visible = false TxtNoOfBags.Locked = FalseTxtGrossWeight.Locked = FalseTxtDematWeight.Locked = FalseTxtClientID.Locked = FalseCmbClientID.Visible = true TxtClientID.Visible = false TxtRemarks.Locked = FalseTxtFlag.Locked = FalseTxtGunnyWeight.Locked = FalseTxtStdDeduction.Locked = FalseTxtAdjNCMSLWeight.Locked = FalseTxtGSLID.Locked = FalseCmbGSLID.Visible = true TxtGSLID.Visible = false 
CmbCxTFNo.Text = TxtCxTFNoCmbClientID.Text = TxtClientIDCmbGSLID.Text = TxtGSLID
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsTxn_CxTF_GSL.Delete
   RsTxn_CxTF_GSL.Update
   If RsTxn_CxTF_GSL.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CxTF_GSL.MoveNext
   If RsTxn_CxTF_GSL.EOF() Then
      RsTxn_CxTF_GSL.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CxTF_GSL.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If Edit_Flg = "A" Then
   If RsTxn_CxTF_GSL.RecordCount > 0 Then
      RsTxn_CxTF_GSL.MoveFirst
      RsTxn_CxTF_GSL.Find "= '"&Txt.Text &"'"
      if Not RsTxn_CxTF_GSL.eof Then
        MsgBox "Duplicate Party Name Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsTxn_CxTF_GSL.AddNew
      RsTxn_CxTF_GSL!GSLDetailID = GetGSLDetailID
      TxtGSLDetailID.text  = RsTxn_CxTF_GSL!GSLDetailID
Else
    RsTxn_CxTF_GSL.MoveFirst
    RsTxn_CxTF_GSL.Find "= '"&Txt.Text &"'"
    If Not RsTxn_CxTF_GSL.EOF Then
       If RsTxn_CxTF_GSL!GSLDetailID  <> TxtGSLDetailID.text  Then
          MsgBox "Duplicate Party Name Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsTxn_CxTF_GSL.MoveFirst
    RsTxn_CxTF_GSL.Find " GSLDetailID= " &TxtGSLDetailID.Text 
End If
RsTxn_CxTF_GSL!GSLDetailID= TxtGSLDetailID.Text RsTxn_CxTF_GSL!CxTFNo= mCxTFNoRsTxn_CxTF_GSL!NoOfBags= TxtNoOfBags.Text RsTxn_CxTF_GSL!GrossWeight= TxtGrossWeight.Text RsTxn_CxTF_GSL!DematWeight= TxtDematWeight.Text RsTxn_CxTF_GSL!ClientID= mClientIDRsTxn_CxTF_GSL!Remarks= TxtRemarks.Text RsTxn_CxTF_GSL!Flag= TxtFlag.Text RsTxn_CxTF_GSL!GunnyWeight= TxtGunnyWeight.Text RsTxn_CxTF_GSL!StdDeduction= TxtStdDeduction.Text RsTxn_CxTF_GSL!AdjNCMSLWeight= TxtAdjNCMSLWeight.Text RsTxn_CxTF_GSL!GSLID= mGSLID
RsTxn_CxTF_GSL.update
Call InData
End Sub
Private Function GetGSLDetailID() As Double
Dim RetVal As Double
tmp = "Select max(GSLDetailID) from Txn_CxTF_GSL"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetGSLDetailID = RetVal
End Function
Private sub Grid_Head
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13
MSHFlexGrid1.TextMatrix(0,0) = "GSLDetail ID"
MSHFlexGrid1.TextMatrix(0,1) = "CxTF No"
MSHFlexGrid1.TextMatrix(0,2) = "No Of Bags"
MSHFlexGrid1.TextMatrix(0,3) = "Gross Weight"
MSHFlexGrid1.TextMatrix(0,4) = "Demat Weight"
MSHFlexGrid1.TextMatrix(0,5) = "Client ID"
MSHFlexGrid1.TextMatrix(0,6) = "Remarks"
MSHFlexGrid1.TextMatrix(0,7) = "Flag"
MSHFlexGrid1.TextMatrix(0,8) = "Gunny Weight"
MSHFlexGrid1.TextMatrix(0,9) = "Std Deduction"
MSHFlexGrid1.TextMatrix(0,10) = "Adj NCMSL Weight"
MSHFlexGrid1.TextMatrix(0,11) = "Weight Gain Loss"
MSHFlexGrid1.TextMatrix(0,12) = "Lot No"
End sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CxTF_GSL.MoveFirst
RsTxn_CxTF_GSL.Find "GSLDetailID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsTxn_CxTF_GSL.EOF Then
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
Private Sub  CmdClientID_Click()
    Frm_Mst_Client.Show
End Sub
'---- 
Private Sub  CmbClientID_GotFocus()
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
Private Sub  CmdGSLID_Click()
    Frm_Mst_GSL.Show
End Sub
'---- 
Private Sub  CmbGSLID_GotFocus()
tmp = CmbGSLID.Text
Call TableMst_GSL
CmbGSLID.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_GSL.RecordCount
    CmbGSLID.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next
CmbGSLID.Text = tmp
End Sub
'---- 
Private Sub CmbGSLID_LostFocus()
If CmbGSLID.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbGSLID.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid  selection "
   CmbGSLID.SetFocus
   Exit Sub
End If
mGSLID = RsMst_GSL!GSLID
End Sub
'---- 
