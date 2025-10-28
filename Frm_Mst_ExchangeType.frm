VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ExchangeType 
   Caption         =   "ExchangeType Master"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   MDIChild        =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   7755
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   9360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   16
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   15
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
      Height          =   2745
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   7245
      Begin VB.TextBox TxtMult2 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2130
         Width           =   1560
      End
      Begin VB.TextBox TxtMult1 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1560
         Width           =   1560
      End
      Begin VB.ComboBox CmbFlag 
         Height          =   315
         ItemData        =   "Frm_Mst_ExchangeType.frx":0000
         Left            =   1800
         List            =   "Frm_Mst_ExchangeType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Active"
         Top             =   1110
         Width           =   4695
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtExchangeType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   660
         Width           =   4680
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1110
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label Label2 
         Caption         =   "Outstanding Amout Multiple of( for AM)"
         Height          =   420
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Outstanding Amout Multiple of (for All)"
         Height          =   405
         Left            =   240
         TabIndex        =   24
         Top             =   1530
         Width           =   1800
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "ExchangeTypeID"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblExchangeType 
         Caption         =   "Exchange Type"
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
         Left            =   240
         TabIndex        =   20
         Top             =   690
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Status"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1117
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   7245
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   8
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   11
         Top             =   660
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
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   12
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   9
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2895
         TabIndex        =   13
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   10
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   14
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   6
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   7
         Top             =   660
         Width           =   1400
      End
   End
End
Attribute VB_Name = "FrmMst_ExchangeType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub

If CmbFlag.Text <> "Active" And CmbFlag.Text <> "De-Active" Then
   MsgBox "Invalid selection!!! "
   CmbFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_ExchangeType
If RsMst_ExchangeType.RecordCount = 0 Then
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
Call Gen_Create_Xls("Mst_ExchangeType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & "Mst_ExchangeType.xls")
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
RsMst_ExchangeType.MoveNext
If RsMst_ExchangeType.EOF Then
   RsMst_ExchangeType.MoveLast
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
RsMst_ExchangeType.MovePrevious
If RsMst_ExchangeType.BOF Then
   RsMst_ExchangeType.MoveFirst
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
RsMst_ExchangeType.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_ExchangeType.MoveLast
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
   If RsMst_ExchangeType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
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
RsMst_ExchangeType.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_ExchangeType.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For I = 1 To RsMst_ExchangeType.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(RsMst_ExchangeType.Fields(c)), "", RsMst_ExchangeType.Fields(c))
    Next
    RsMst_ExchangeType.MoveNext
Next
End Sub
Public Sub Indata()

TxtExchangeTypeID.Locked = True
TxtExchangeType.Locked = True
TxtMult1.Locked = True
TxtMult2.Locked = True
TxtFlag.Locked = True

CmbFlag.Visible = False
TxtFlag.Visible = True


TxtExchangeTypeID = IIf(IsNull(RsMst_ExchangeType!ExchangeTypeID), "", RsMst_ExchangeType!ExchangeTypeID)
TxtExchangeType = IIf(IsNull(RsMst_ExchangeType!ExchangeType), "", RsMst_ExchangeType!ExchangeType)
TxtFlag = IIf(IsNull(RsMst_ExchangeType!Flag), "", RsMst_ExchangeType!Flag)
TxtMult1 = IIf(IsNull(RsMst_ExchangeType!MultOfAll), "", RsMst_ExchangeType!MultOfAll)
TxtMult2 = IIf(IsNull(RsMst_ExchangeType!MultOfAM), "", RsMst_ExchangeType!MultOfAM)

If RsMst_ExchangeType!Flag = "A" Then
   TxtFlag = "Active" 'TxtFlag.Text
Else
   TxtFlag = "De-Active"  'TxtFlag.Text
End If

If RsMst_ExchangeType.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtExchangeType_LostFocus()
TxtExchangeType = Trim(TxtExchangeType)
End Sub

Private Sub TxtExchangeTypeID_LostFocus()
If TxtExchangeTypeID = "" Then Exit Sub
If IsNumeric(TxtExchangeTypeID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtExchangeTypeID.SetFocus
   Exit Sub
End If
TxtExchangeTypeID = Val(TxtExchangeTypeID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtExchangeTypeID.Locked = True
TxtExchangeType.Locked = False
TxtMult1.Locked = False
TxtMult2.Locked = False
TxtFlag.Locked = False

TxtExchangeTypeID = ""

TxtExchangeType = ""
TxtMult1 = ""
TxtMult2 = ""
TxtFlag = ""

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtExchangeTypeID.Locked = True
TxtExchangeType.Locked = False
TxtMult1.Locked = False
TxtMult2.Locked = False
TxtFlag.Locked = False

End Sub
Private Sub DeleteCmd_Click()
Exit Sub
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_ExchangeType.Delete
   RsMst_ExchangeType.Update
   If RsMst_ExchangeType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ExchangeType.MoveNext
   If RsMst_ExchangeType.EOF() Then
      RsMst_ExchangeType.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_ExchangeType.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtExchangeType.Text = "" Then
   MsgBox "Blank ExchangeType Not Allowed !!! "
   TxtExchangeType.SetFocus
   Exit Sub
End If

If CmbFlag.Text <> "Active" And CmbFlag.Text <> "De-Active" Then
   MsgBox "Invalid selection!!! "
   CmbFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_ExchangeType.RecordCount > 0 Then
      RsMst_ExchangeType.MoveFirst
      RsMst_ExchangeType.Find "ExchangeType= '" & TxtExchangeType.Text & "'"
      If Not RsMst_ExchangeType.EOF Then
        MsgBox "Duplicate Exchange Name Not Allowed !!! "
         TxtExchangeType.SetFocus
         Exit Sub
      End If
   End If
      RsMst_ExchangeType.AddNew
      RsMst_ExchangeType!ExchangeTypeID = GetExchangeTypeID
      RsMst_ExchangeType!G_UID = GetGUID
      TxtExchangeTypeID.Text = RsMst_ExchangeType!ExchangeTypeID
Else
    RsMst_ExchangeType.MoveFirst
    RsMst_ExchangeType.Find "ExchangeType= '" & TxtExchangeType.Text & "'"
    If Not RsMst_ExchangeType.EOF Then
       If RsMst_ExchangeType!ExchangeTypeID <> TxtExchangeTypeID.Text Then
          MsgBox "Duplicate Exchange Name Not Allowed !!! "
          TxtExchangeType.SetFocus
          Exit Sub
       End If
    End If
    RsMst_ExchangeType.MoveFirst
    RsMst_ExchangeType.Find " ExchangeTypeID= " & TxtExchangeTypeID.Text
End If

RsMst_ExchangeType!ExchangeTypeID = TxtExchangeTypeID.Text
RsMst_ExchangeType!ExchangeType = TxtExchangeType.Text
RsMst_ExchangeType!MultOfAll = Val(TxtMult1.Text)
RsMst_ExchangeType!MultOfAM = Val(TxtMult2.Text)

If CmbFlag = "Active" Then
   RsMst_ExchangeType!Flag = "A" 'TxtFlag.Text
Else
   RsMst_ExchangeType!Flag = "D" 'TxtFlag.Text
End If

RsMst_ExchangeType.Update
Call Indata
End Sub
Private Function GetExchangeTypeID() As Double
Dim Retval As Double
tmp = "Select max(ExchangeTypeID) from Mst_ExchangeType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetExchangeTypeID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0, 0) = "ExchangeTypeID"
MSHFlexGrid1.TextMatrix(0, 1) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 2) = "Flag"

MSHFlexGrid1.ColWidth(1) = 3200

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeTypeID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_ExchangeType.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtMult1_LostFocus()
If TxtMult1 = "" Then Exit Sub
If Not IsNumeric(TxtMult1) Or Val(TxtMult1) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtMult1.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtMult2_LostFocus()
If TxtMult2 = "" Then Exit Sub
If Not IsNumeric(TxtMult2) Or Val(TxtMult2) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtMult2.SetFocus
   Exit Sub
End If
End Sub

