VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_IncomeExpenseType 
   Caption         =   "Income Expense Type"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   7455
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2625
      Left            =   120
      TabIndex        =   0
      Top             =   4215
      Visible         =   0   'False
      Width           =   10695
      Begin VB.ComboBox CmbSFlag 
         Height          =   315
         ItemData        =   "FrmMst_IncomeExpenseType.frx":0000
         Left            =   6135
         List            =   "FrmMst_IncomeExpenseType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   450
         Width           =   1620
      End
      Begin VB.ComboBox CmbSIEFlag 
         Height          =   315
         Left            =   3510
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   450
         Width           =   2565
      End
      Begin VB.TextBox TxtSIEType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   75
         MaxLength       =   50
         TabIndex        =   26
         Top             =   450
         Width           =   3375
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   7815
         TabIndex        =   25
         Top             =   450
         Width           =   1020
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   8850
         TabIndex        =   15
         Top             =   450
         Width           =   1020
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   14
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
         Height          =   195
         Left            =   6720
         TabIndex        =   31
         Top             =   210
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Income Expense Flag"
         Height          =   195
         Left            =   4027
         TabIndex        =   30
         Top             =   210
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Income Expense Type"
         Height          =   195
         Left            =   967
         TabIndex        =   29
         Top             =   210
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3420
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7380
      Begin VB.ComboBox CmbExchange 
         Height          =   315
         Left            =   2430
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   1200
         Width           =   4680
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2910
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2445
         Width           =   4680
      End
      Begin VB.ComboBox CmbFlag 
         Height          =   315
         ItemData        =   "FrmMst_IncomeExpenseType.frx":0021
         Left            =   2430
         List            =   "FrmMst_IncomeExpenseType.frx":002B
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   2040
         Width           =   4680
      End
      Begin VB.ComboBox CmbIEFlag 
         Height          =   315
         ItemData        =   "FrmMst_IncomeExpenseType.frx":0042
         Left            =   2430
         List            =   "FrmMst_IncomeExpenseType.frx":004C
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1635
         Width           =   4695
      End
      Begin VB.TextBox TxtIETypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox TxtIEType 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   1
         Top             =   693
         Width           =   4680
      End
      Begin VB.TextBox TxtIEFlag 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1635
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.TextBox TxtExchange 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   4680
      End
      Begin VB.Label Label7 
         Caption         =   "Exchange Type"
         Height          =   225
         Left            =   240
         TabIndex        =   38
         Top             =   1275
         Width           =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "Updated :"
         Height          =   240
         Left            =   240
         TabIndex        =   35
         Top             =   2985
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Created  :"
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   2490
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   2085
         Width           =   1800
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "Income Expense TypeID"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   300
         Width           =   1920
      End
      Begin VB.Label LblExchangeType 
         Caption         =   "Income/Expense Type"
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
         TabIndex        =   19
         Top             =   753
         Width           =   1950
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
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
         TabIndex        =   20
         Top             =   1665
         Width           =   2280
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   22
      Top             =   3510
      Width           =   7365
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   7
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   10
         Top             =   660
         Width           =   1425
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1523
         TabIndex        =   4
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1523
         TabIndex        =   11
         Top             =   660
         Width           =   1425
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2946
         TabIndex        =   8
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2946
         TabIndex        =   12
         Top             =   660
         Width           =   1425
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4369
         TabIndex        =   9
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4369
         TabIndex        =   13
         Top             =   660
         Width           =   1425
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5790
         TabIndex        =   5
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5790
         TabIndex        =   6
         Top             =   660
         Width           =   1425
      End
   End
End
Attribute VB_Name = "FrmMst_IncomeExpenseType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mExchangeTypeID As Double

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub

If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbIEFlag_LostFocus()
If CmbIEFlag.Text = "" Then Exit Sub

If LCase(CmbIEFlag.Text) <> "income" And LCase(CmbIEFlag.Text) <> "expense" Then
   MsgBox "Invalid selection!!! "
   CmbIEFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSFlag_LostFocus()
If CmbSFlag.Text = "" Then Exit Sub

If LCase(CmbSFlag.Text) <> "active" And LCase(CmbSFlag.Text) <> "de-active" Then
   MsgBox "Invalid selection!!! "
   CmbSFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSIEFlag_LostFocus()
If CmbSIEFlag.Text = "" Then Exit Sub

If LCase(CmbSIEFlag.Text) <> "income" And LCase(CmbSIEFlag.Text) <> "expense" Then
   MsgBox "Invalid selection!!! "
   CmbSIEFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdSearch_Click()
Dim x As Integer
Dim wc As String

wc = ""
RsMst_IncomeExpenseType.MoveFirst
Call Grid_Head

If TxtSIEType <> "" Then
   wc = " where IEType like '%" & Trim(TxtSIEType) & "%' "
End If

If CmbSIEFlag.Text <> "" Then
   If wc = "" Then
      wc = " where IEFlag='" & Left(CmbSIEFlag.Text, 1) & "'"
   Else
      wc = wc & " and IEFlag='" & Left(CmbSIEFlag.Text, 1) & "'"
   End If
End If

If CmbSFlag.Text <> "" Then
   If wc = "" Then
      wc = " where Flag='" & Left(CmbSFlag.Text, 1) & "'"
   Else
      wc = wc & " and Flag='" & Left(CmbSFlag.Text, 1) & "'"
   End If
End If

tmp = "Select IETypeID, IEType, IEFlag, Flag FROM Mst_IncomeExpenseType"
tmp = tmp & wc
tmp = tmp & " order by IETypeID "
Call Tmp4Table

'If TmpRs4.RecordCount <= 0 Then
'   MsgBox "No Such Income Expense Type Exist in DataBase!!!"
'   TxtSIEType.SetFocus
'   Exit Sub
'End If
Call ShowGrid
End Sub

Private Sub ShowGrid()
MSHFlexGrid1.Rows = TmpRs4.RecordCount + 2
For x = 1 To TmpRs4.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs4.Fields(c)), "", TmpRs4.Fields(c))
    Next
    MSHFlexGrid1.TextMatrix(x, 2) = IIf(MSHFlexGrid1.TextMatrix(x, 2) = "E", "Expense", "Income")
    MSHFlexGrid1.TextMatrix(x, 3) = IIf(MSHFlexGrid1.TextMatrix(x, 3) = "A", "Active", "De-active")
    TmpRs4.MoveNext
Next
TmpRs4.Close
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
TxtSIEType.Text = ""
CmbSIEFlag.Text = ""
CmbSFlag.Text = ""

Call TableMst_IncomeExpenseType
If RsMst_IncomeExpenseType.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_IncomeExpenseType.xls")

'Call Gen_Create_Xls("Mst_IncomeExpenseType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_IncomeExpenseType.xls")
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
RsMst_IncomeExpenseType.MoveNext
If RsMst_IncomeExpenseType.EOF Then
   RsMst_IncomeExpenseType.MoveLast
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
RsMst_IncomeExpenseType.MovePrevious
If RsMst_IncomeExpenseType.BOF Then
   RsMst_IncomeExpenseType.MoveFirst
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
RsMst_IncomeExpenseType.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_IncomeExpenseType.MoveLast
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
   If RsMst_IncomeExpenseType.RecordCount = 0 Then Exit Sub
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

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'TxtSIEType.Text = ""
'CmbSIEFlag.Text = ""
'CmbSFlag.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Public Sub Indata()

TxtIETypeID.Locked = True
TxtIEType.Locked = True


CmbIEFlag.Visible = False
TxtIEFlag.Visible = True

CmbFlag.Visible = False
TxtFlag.Visible = True

CmbExchange.Visible = False
TxtExchange.Visible = True

TxtIETypeID = RsMst_IncomeExpenseType!IETypeID '), "", RsMst_IncomeExpenseType!IETypeID)
TxtIEType = IIf(IsNull(RsMst_IncomeExpenseType!IEType), "", RsMst_IncomeExpenseType!IEType)
TxtIEFlag = IIf(IsNull(RsMst_IncomeExpenseType!IEFlag), "", RsMst_IncomeExpenseType!IEFlag)
TxtFlag = IIf(IsNull(RsMst_IncomeExpenseType!Flag), "", RsMst_IncomeExpenseType!Flag)

mExchangeTypeID = IIf(IsNull(RsMst_IncomeExpenseType!ExchangeTypeID), 0, RsMst_IncomeExpenseType!ExchangeTypeID)
TxtExchange.Text = GetExchangeName(mExchangeTypeID)

If RsMst_IncomeExpenseType!IEFlag = "I" Then
   TxtIEFlag = "Income" 'TxtIEFlag.Text
Else
   TxtIEFlag = "Expense"  'TxtIEFlag.Text
End If
If RsMst_IncomeExpenseType!Flag = "A" Then
   TxtFlag = "Active" 'TxtFlag.Text
Else
   TxtFlag = "De-Active"  'TxtFlag.Text
End If
TxtCreated = IIf(IsNull(RsMst_IncomeExpenseType!CreatedBy), "", RsMst_IncomeExpenseType!CreatedBy) & " :- " & IIf(IsNull(RsMst_IncomeExpenseType!CreatedDate), "", RsMst_IncomeExpenseType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_IncomeExpenseType!UpdatedBy), "", RsMst_IncomeExpenseType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_IncomeExpenseType!UpdatedDate), "", RsMst_IncomeExpenseType!UpdatedDate)

If RsMst_IncomeExpenseType.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtIEType_LostFocus()
TxtIEType = Trim(TxtIEType)
End Sub

Private Sub TxtIETypeID_LostFocus()
If TxtIETypeID = "" Then Exit Sub
If IsNumeric(TxtIETypeID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtIETypeID.SetFocus
   Exit Sub
End If
TxtIETypeID = Val(TxtIETypeID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtIEType.Locked = False

TxtIEFlag.Visible = False
TxtFlag.Visible = False

CmbIEFlag.Visible = True
CmbFlag.Visible = True

CmbIEFlag.Text = ""
CmbFlag.Text = ""


TxtIETypeID = ""

TxtIEType = ""
TxtCreated = ""
TxtUpdated = ""

TxtExchange.Visible = False
CmbExchange.Visible = True
CmbExchange.Text = ""


End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtIEType.Locked = False

TxtIEFlag.Locked = False
TxtFlag.Locked = False

TxtIEFlag.Visible = False
TxtFlag.Visible = False

CmbIEFlag.Visible = True
CmbIEFlag.Text = TxtIEFlag.Text
CmbFlag.Visible = True
CmbFlag.Text = TxtFlag.Text

TxtExchange.Visible = False
CmbExchange.Visible = True
CmbExchange.Text = TxtExchange


End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_IncomeExpenseType.Delete
   RsMst_IncomeExpenseType.Update
   If RsMst_IncomeExpenseType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_IncomeExpenseType.MoveNext
   If RsMst_IncomeExpenseType.EOF() Then
      RsMst_IncomeExpenseType.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_IncomeExpenseType.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtIEType.Text = "" Then
   MsgBox "Blank Income Expense Type Not Allowed !!! "
   TxtIEType.SetFocus
   Exit Sub
End If

If CmbIEFlag.Text = "" Then
   MsgBox "Blank Flag not allowed!!! "
   CmbIEFlag.SetFocus
   Exit Sub
End If

If CmbFlag.Text = "" Then
   MsgBox "Blank Status not allowed!!! "
   CmbFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_IncomeExpenseType.RecordCount > 0 Then
      RsMst_IncomeExpenseType.MoveFirst
      RsMst_IncomeExpenseType.Find "IEType= '" & TxtIEType.Text & "'"
      If Not RsMst_IncomeExpenseType.EOF Then
         MsgBox "Duplicate Income Expense Type Not Allowed !!! "
         TxtIEType.SetFocus
         Exit Sub
      End If
   End If
   RsMst_IncomeExpenseType.AddNew
   RsMst_IncomeExpenseType!IETypeID = GetIncomeExpenseTypeID
   RsMst_IncomeExpenseType!G_UID = GetGUID
   TxtIETypeID.Text = RsMst_IncomeExpenseType!IETypeID
Else
   RsMst_IncomeExpenseType.MoveFirst
   RsMst_IncomeExpenseType.Find "IEType= '" & TxtIEType.Text & "'"
   If Not RsMst_IncomeExpenseType.EOF Then
      If RsMst_IncomeExpenseType!IETypeID <> TxtIETypeID.Text Then
         MsgBox "Duplicate Income Expense Type  Not Allowed !!! "
         TxtIEType.SetFocus
         Exit Sub
      End If
   End If
   RsMst_IncomeExpenseType.MoveFirst
   RsMst_IncomeExpenseType.Find " IETypeID= " & TxtIETypeID.Text
End If

'RsMst_IncomeExpenseType!IETypeID = TxtIETypeID.Text
RsMst_IncomeExpenseType!IEType = TxtIEType.Text

If LCase(CmbIEFlag.Text) = "income" Then
   RsMst_IncomeExpenseType!IEFlag = "I" 'TxtIEFlag.Text
Else
   RsMst_IncomeExpenseType!IEFlag = "E" 'TxtIEFlag.Text
End If

If CmbExchange.Text <> "" Then
   RsMst_IncomeExpenseType!ExchangeTypeID = mExchangeTypeID
Else
   RsMst_IncomeExpenseType!ExchangeTypeID = Null
End If

If LCase(CmbFlag.Text) = "active" Then
   RsMst_IncomeExpenseType!Flag = "A" 'TxtFlag.Text
Else
   RsMst_IncomeExpenseType!Flag = "D" 'TxtFlag.Text
End If
If IsNull(RsMst_IncomeExpenseType!CreatedBy) = True Or RsMst_IncomeExpenseType!CreatedBy = "" Then
   RsMst_IncomeExpenseType!CreatedBy = Gen_UserLoginID
   RsMst_IncomeExpenseType!CreatedDate = Now
Else
   RsMst_IncomeExpenseType!UpdatedBy = Gen_UserLoginID
   RsMst_IncomeExpenseType!UpdatedDate = Now
End If


RsMst_IncomeExpenseType.Update
Call Indata
End Sub
Private Function GetIncomeExpenseTypeID() As Double
Dim Retval As Double
tmp = "Select max(IETypeID) from Mst_IncomeExpenseType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetIncomeExpenseTypeID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4
MSHFlexGrid1.TextMatrix(0, 0) = "Income/Expense TypeID"
MSHFlexGrid1.TextMatrix(0, 1) = "Income/Expense Type"
MSHFlexGrid1.TextMatrix(0, 2) = "Flag"
MSHFlexGrid1.TextMatrix(0, 3) = "Status"

MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 1400

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_IncomeExpenseType.MoveFirst
RsMst_IncomeExpenseType.Find "IETypeID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_IncomeExpenseType.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType
CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    'If RsMst_ExchangeType.Fields(0) <> 1 Then
    CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    'End If
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp

End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   Exit Sub
End If

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchange.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub
