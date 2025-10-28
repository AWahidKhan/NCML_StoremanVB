VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_CoExpense 
   Caption         =   "Corporate Office Expense"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1305
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   7335
      Begin VB.ComboBox CmbSFinanYr 
         Height          =   315
         ItemData        =   "FrmMst_CoExpense.frx":0000
         Left            =   120
         List            =   "FrmMst_CoExpense.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   2700
      End
      Begin VB.ComboBox CmbSFlag 
         Height          =   315
         ItemData        =   "FrmMst_CoExpense.frx":0004
         Left            =   2880
         List            =   "FrmMst_CoExpense.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   27
         Text            =   "Active"
         Top             =   480
         Width           =   1620
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   4590
         TabIndex        =   26
         Top             =   450
         Width           =   1020
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   5625
         TabIndex        =   18
         Top             =   450
         Width           =   1020
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   17
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
         Left            =   3360
         TabIndex        =   29
         Top             =   210
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Financial Year"
         Height          =   195
         Left            =   405
         TabIndex        =   28
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   14580
      Begin VB.TextBox TxtBudgetIncomeWord 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1602
         Width           =   7935
      End
      Begin VB.TextBox TxtBankInterestWord 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2056
         Width           =   7935
      End
      Begin VB.TextBox TxtAOROCMGExpenseWord 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2520
         Width           =   7935
      End
      Begin VB.TextBox TxtCoExpenseWord 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1148
         Width           =   7935
      End
      Begin VB.TextBox TxtAOROCMGExpense 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2520
         Width           =   3960
      End
      Begin VB.TextBox TxtUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3945
         Width           =   3960
      End
      Begin VB.TextBox TxtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3450
         Width           =   3960
      End
      Begin VB.TextBox TxtBudgetIncome 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1602
         Width           =   3960
      End
      Begin VB.TextBox TxtBankInterest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2056
         Width           =   3960
      End
      Begin VB.TextBox TxtExpenseAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1148
         Width           =   3960
      End
      Begin VB.ComboBox CmbFinancialYr 
         Height          =   315
         ItemData        =   "FrmMst_CoExpense.frx":0025
         Left            =   2430
         List            =   "FrmMst_CoExpense.frx":0027
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   694
         Width           =   2175
      End
      Begin VB.TextBox TxtCOExpenseID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtFinancialYr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   694
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox CmbFlag 
         Height          =   315
         ItemData        =   "FrmMst_CoExpense.frx":0029
         Left            =   2430
         List            =   "FrmMst_CoExpense.frx":0033
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2985
         Width           =   3960
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2430
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2985
         Visible         =   0   'False
         Width           =   3960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "AO RO CMG Expense"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   2580
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Budget Income"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1662
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bank Interest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   2116
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CO Expense Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1208
         Width           =   1740
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Updated :"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   4005
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Created  :"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3510
         Width           =   690
      End
      Begin VB.Label LblExchangeTypeID 
         AutoSize        =   -1  'True
         Caption         =   "CO Expense ID"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label LblFlag 
         AutoSize        =   -1  'True
         Caption         =   "Financial Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   754
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   3045
         Width           =   555
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   14565
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   7
         Top             =   240
         Width           =   2865
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   12
         Top             =   660
         Width           =   2865
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2970
         TabIndex        =   8
         Top             =   240
         Width           =   2865
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2970
         TabIndex        =   13
         Top             =   660
         Width           =   2865
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5835
         TabIndex        =   9
         Top             =   240
         Width           =   2865
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5835
         TabIndex        =   14
         Top             =   660
         Width           =   2865
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   8700
         TabIndex        =   10
         Top             =   240
         Width           =   2865
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   8700
         TabIndex        =   15
         Top             =   660
         Width           =   2865
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11565
         TabIndex        =   11
         Top             =   240
         Width           =   2865
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11565
         TabIndex        =   16
         Top             =   660
         Width           =   2865
      End
   End
End
Attribute VB_Name = "FrmMst_CoExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub CmbFinancialYr_GotFocus()
Dim mCurYr      As Variant
Dim mNxtYr      As Variant
tmp = CmbFinancialYr.Text
CmbFinancialYr.Clear
CmbFinancialYr.AddItem "2009-2010"
For I = 2010 To Year(Date)
    mCurYr = I
    mNxtYr = I + 1
    If Year(Date) = mCurYr Then
        If Month(Date) >= 4 Then
           CmbFinancialYr.AddItem mCurYr & "-" & mNxtYr
           Exit For
        Else
           Exit For
        End If
    End If
    CmbFinancialYr.AddItem mCurYr & "-" & mNxtYr
Next
CmbFinancialYr.Text = tmp
End Sub

Private Sub CmbFinancialYr_LostFocus()
If CmbFinancialYr.Text = "" Then Exit Sub
Dim flg As Variant
Dim myr As Variant

flg = "N"

If Month(Now) >= 4 Then
   myr = Year(Now) + 1
Else
   myr = Year(Now)
End If


For I = 2009 To myr 'Year(Now)
    If CmbFinancialYr.Text = I & "-" & (I + 1) Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbFinancialYr.SetFocus
   Exit Sub
End If
End Sub
Private Sub CmbSFinanYr_GotFocus()
Dim mCurYr      As Variant
Dim mNxtYr      As Variant
tmp = CmbSFinanYr
CmbSFinanYr.Clear
CmbSFinanYr.AddItem "2009-2010"
For I = 2010 To Year(Date)
    mCurYr = I
    mNxtYr = I + 1
    If Year(Date) = mCurYr Then
        If Month(Date) >= 4 Then
           CmbSFinanYr.AddItem mCurYr & "-" & mNxtYr
           Exit For
        Else
           Exit For
        End If
    End If
    CmbSFinanYr.AddItem mCurYr & "-" & mNxtYr
Next
CmbSFinanYr = tmp
End Sub

Private Sub CmbSFinanYr_LostFocus()
If CmbSFinanYr.Text = "" Then Exit Sub
Dim flg As Variant
Dim myr As Variant

flg = "N"

If Month(Now) >= 4 Then
   myr = Year(Now) + 1
Else
   myr = Year(Now)
End If


For I = 2009 To myr 'Year(Now)
    If CmbSFinanYr.Text = I & "-" & (I + 1) Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbSFinanYr.SetFocus
   Exit Sub
End If
End Sub

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

If CmbIEFlag.Text <> "Income" And CmbIEFlag.Text <> "Expense" Then
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

If CmbSIEFlag.Text <> "Income" And CmbSIEFlag.Text <> "Expense" Then
   MsgBox "Invalid selection!!! "
   CmbSIEFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdSearch_Click()
Dim x As Integer
Dim wc As String

wc = ""
RsMst_CoExpense.MoveFirst
Call Grid_Head


If CmbSFinanYr.Text <> "" Then
   wc = " where FinancialYr='" & CmbSFinanYr.Text & "'"
End If

If CmbSFlag.Text <> "" Then
   If wc = "" Then
      wc = " where Flag='" & Left(CmbSFlag.Text, 1) & "'"
   Else
      wc = wc & " and Flag='" & Left(CmbSFlag.Text, 1) & "'"
   End If
End If

tmp = "select CoExpenseID,FinancialYr,ExpenseAmt, Flag, BudgetIncome, BankInterest, AOROCMGExpense FROM Mst_CoExpense"
tmp = tmp & wc
tmp = tmp & " order by CoExpenseID "
Call Tmp4Table

If TmpRs4.RecordCount <= 0 Then
   MsgBox "No record found!!!"
   CmbSFinanYr.SetFocus
   Exit Sub
End If
Call ShowGrid
End Sub

Private Sub ShowGrid()
MSHFlexGrid1.Rows = TmpRs4.RecordCount + 2
For x = 1 To TmpRs4.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs4.Fields(c)), "", TmpRs4.Fields(c))
        If c = 3 Then
            MSHFlexGrid1.TextMatrix(x, 3) = IIf(TmpRs4.Fields(3) = "A", "Active", "De-active")
        End If
    Next
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
CmbSFinanYr.Text = ""
CmbSFlag.Text = ""

Call TableMst_CoExpense
If RsMst_CoExpense.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_CoExpense.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_CoExpense.xls")
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
RsMst_CoExpense.MoveNext
If RsMst_CoExpense.EOF Then
   RsMst_CoExpense.MoveLast
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
RsMst_CoExpense.MovePrevious
If RsMst_CoExpense.BOF Then
   RsMst_CoExpense.MoveFirst
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
RsMst_CoExpense.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_CoExpense.MoveLast
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
   If RsMst_CoExpense.RecordCount = 0 Then Exit Sub
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
'CmbSFinanYr.Text = ""
'CmbSFlag.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Public Sub Indata()

TxtCOExpenseID.Locked = True
TxtFinancialYr.Locked = True
TxtExpenseAmt.Locked = True
TxtBudgetIncome.Locked = True
TxtBankInterest.Locked = True
TxtAOROCMGExpense.Locked = True
TxtFlag.Locked = True

CmbFinancialYr.Visible = False
TxtFinancialYr.Visible = True

CmbFlag.Visible = False
TxtFlag.Visible = True


TxtCOExpenseID = IIf(IsNull(RsMst_CoExpense!CoExpenseID), "", RsMst_CoExpense!CoExpenseID)
TxtFinancialYr = IIf(IsNull(RsMst_CoExpense!FinancialYr), "", RsMst_CoExpense!FinancialYr)
TxtExpenseAmt = IIf(IsNull(RsMst_CoExpense!ExpenseAmt), "", RsMst_CoExpense!ExpenseAmt)
TxtBudgetIncome = IIf(IsNull(RsMst_CoExpense!BudgetIncome), "", RsMst_CoExpense!BudgetIncome)
TxtBankInterest = IIf(IsNull(RsMst_CoExpense!BankInterest), "", RsMst_CoExpense!BankInterest)
TxtAOROCMGExpense = IIf(IsNull(RsMst_CoExpense!aorocmgexpense), "", RsMst_CoExpense!aorocmgexpense)
TxtFlag = IIf(IsNull(RsMst_CoExpense!Flag), "", RsMst_CoExpense!Flag)
If RsMst_CoExpense!Flag = "A" Then
   TxtFlag = "Active" 'TxtFlag.Text
Else
   TxtFlag = "De-Active"  'TxtFlag.Text
End If
TxtCreated = IIf(IsNull(RsMst_CoExpense!CreatedBy), "", RsMst_CoExpense!CreatedBy) & " :- " & IIf(IsNull(RsMst_CoExpense!CreatedDate), "", RsMst_CoExpense!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_CoExpense!UpdatedBy), "", RsMst_CoExpense!UpdatedBy) & " :- " & IIf(IsNull(RsMst_CoExpense!UpdatedDate), "", RsMst_CoExpense!UpdatedDate)

TxtCoExpenseWord.Text = NumStrConv(Val(TxtExpenseAmt.Text))
TxtBudgetIncomeWord.Text = NumStrConv(Val(TxtBudgetIncome.Text))
TxtBankInterestWord.Text = NumStrConv(Val(TxtBankInterest.Text))
TxtAOROCMGExpenseWord.Text = NumStrConv(Val(TxtAOROCMGExpense.Text))

If RsMst_CoExpense.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtCOExpenseID_LostFocus()
If TxtCOExpenseID = "" Then Exit Sub
If IsNumeric(TxtCOExpenseID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtCOExpenseID.SetFocus
   Exit Sub
End If
TxtCOExpenseID = Val(TxtCOExpenseID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtExpenseAmt.Locked = False
TxtBudgetIncome.Locked = False
TxtBankInterest.Locked = False
TxtAOROCMGExpense.Locked = False
TxtFlag.Visible = False
CmbFinancialYr.Visible = True
CmbFlag.Visible = True

TxtCOExpenseID = ""
TxtFinancialYr.Text = ""
TxtExpenseAmt.Text = ""
TxtBudgetIncome.Text = ""
TxtBankInterest.Text = ""
TxtAOROCMGExpense.Text = ""

TxtCoExpenseWord.Text = ""
TxtBudgetIncomeWord.Text = ""
TxtBankInterestWord.Text = ""
TxtAOROCMGExpenseWord.Text = ""

CmbFinancialYr = ""
CmbFlag.Text = ""
TxtFlag = ""
TxtCreated = ""
TxtUpdated = ""
If RsMst_CoExpense.RecordCount > 0 Then
   CmbFinancialYr.SetFocus
End If
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtCOExpenseID.Locked = True
TxtExpenseAmt.Locked = False
TxtBudgetIncome.Locked = False
TxtBankInterest.Locked = False
TxtAOROCMGExpense.Locked = False

'TxtFinancialYr.Visible = False
TxtFlag.Visible = False

'CmbFinancialYr.Visible = True
CmbFinancialYr.Text = TxtFinancialYr.Text
CmbFlag.Visible = True
CmbFlag.Text = TxtFlag.Text

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_CoExpense.Delete
   RsMst_CoExpense.Update
   If RsMst_CoExpense.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_CoExpense.MoveNext
   If RsMst_CoExpense.EOF() Then
      RsMst_CoExpense.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_CoExpense.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
Dim dtStart, dtEnd          As Variant
Dim strSplitYr()          As String

If CmbFinancialYr.Text = "" Then
   MsgBox "Blank Financial Year Not Allowed !!! "
   CmbFinancialYr.SetFocus
   Exit Sub
End If

If TxtExpenseAmt.Text = "" Then
   MsgBox "Blank Expense Amount Not Allowed !!! "
   TxtExpenseAmt.SetFocus
   Exit Sub
End If
If TxtBudgetIncome.Text = "" Then
   MsgBox "Blank Budget Income Not Allowed !!! "
   TxtBudgetIncome.SetFocus
   Exit Sub
End If
If TxtBankInterest.Text = "" Then
   MsgBox "Blank Bank Interest Not Allowed !!! "
   TxtBankInterest.SetFocus
   Exit Sub
End If
If TxtAOROCMGExpense.Text = "" Then
   MsgBox "Blank AO RO CMG Expense Not Allowed !!! "
   TxtAOROCMGExpense.SetFocus
   Exit Sub
End If
If CmbFlag.Text = "" Then
   MsgBox "Please select Status!!! "
   CmbFlag.SetFocus
   Exit Sub
End If

strSplitYr = Split(CmbFinancialYr, "-")
dtStart = "04-01-" & strSplitYr(0)
dtEnd = "03-31-" & strSplitYr(1)

If Edit_Flg = "A" Then
   If RsMst_CoExpense.RecordCount > 0 Then
      RsMst_CoExpense.MoveFirst
      RsMst_CoExpense.Find "FinancialYr= '" & CmbFinancialYr.Text & "'"
      If Not RsMst_CoExpense.EOF Then
        MsgBox "Duplicate Financial Year Not Allowed !!! "
        CmbFinancialYr.SetFocus
        Exit Sub
      End If
   End If
      RsMst_CoExpense.AddNew
      RsMst_CoExpense!CoExpenseID = GetCoExpenseID
      RsMst_CoExpense!G_UID = GetGUID
      TxtCOExpenseID.Text = RsMst_CoExpense!CoExpenseID
Else
    RsMst_CoExpense.MoveFirst
    RsMst_CoExpense.Find "CoExpenseID= '" & TxtCOExpenseID.Text & "'"
    If Not RsMst_CoExpense.EOF Then
       If RsMst_CoExpense!CoExpenseID <> TxtCOExpenseID.Text Then
          MsgBox "Duplicate Co Expense  Not Allowed !!! "
          CmbFinancialYr.SetFocus
          Exit Sub
       End If
    End If
    RsMst_CoExpense.MoveFirst
    RsMst_CoExpense.Find " CoExpenseID= " & TxtCOExpenseID.Text
End If

RsMst_CoExpense!FinancialYr = CmbFinancialYr.Text
RsMst_CoExpense!FromDate = Format(dtStart, Gen_DatFormat)
RsMst_CoExpense!ToDate = Format(dtEnd, Gen_DatFormat)
RsMst_CoExpense!ExpenseAmt = Val(TxtExpenseAmt.Text)
RsMst_CoExpense!BudgetIncome = Val(TxtBudgetIncome.Text)
RsMst_CoExpense!BankInterest = Val(TxtBankInterest.Text)
RsMst_CoExpense!aorocmgexpense = Val(TxtAOROCMGExpense.Text)
If CmbFlag = "Active" Then
   RsMst_CoExpense!Flag = "A" 'TxtFlag.Text
Else
   RsMst_CoExpense!Flag = "D" 'TxtFlag.Text
End If
If IsNull(RsMst_CoExpense!CreatedBy) = True Or RsMst_CoExpense!CreatedBy = "" Then
   RsMst_CoExpense!CreatedBy = Gen_UserLoginID
   RsMst_CoExpense!CreatedDate = Now
Else
   RsMst_CoExpense!UpdatedBy = Gen_UserLoginID
   RsMst_CoExpense!UpdatedDate = Now
End If


RsMst_CoExpense.Update
Call Indata
End Sub
Private Function GetCoExpenseID() As Double
Dim Retval As Double
tmp = "Select max(CoExpenseID) from Mst_CoExpense"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCoExpenseID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 7
MSHFlexGrid1.TextMatrix(0, 0) = "Co Expense TypeID"
MSHFlexGrid1.TextMatrix(0, 1) = "Financial Year"
MSHFlexGrid1.TextMatrix(0, 2) = "Expense Amount"
MSHFlexGrid1.TextMatrix(0, 3) = "Status"
MSHFlexGrid1.TextMatrix(0, 4) = "Budget Income"
MSHFlexGrid1.TextMatrix(0, 5) = "Bank Interest(in %)"
MSHFlexGrid1.TextMatrix(0, 6) = "AO RO CMG Expense"

MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_CoExpense.MoveFirst
RsMst_CoExpense.Find "CoExpenseID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_CoExpense.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub TxtExpenseAmt_LostFocus()
If TxtExpenseAmt = "" Then
   TxtCoExpenseWord.Text = ""
   Exit Sub
End If
If IsNumeric(TxtExpenseAmt) = False Or Val(TxtExpenseAmt) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtExpenseAmt.SetFocus
   Exit Sub
End If
TxtCoExpenseWord.Text = NumStrConv(Val(TxtExpenseAmt.Text))
End Sub
Private Sub Txtbudgetincome_LostFocus()
If TxtBudgetIncome = "" Then
   TxtBudgetIncomeWord.Text = ""
   Exit Sub
End If
If IsNumeric(TxtBudgetIncome) = False Or Val(TxtBudgetIncome) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtBudgetIncome.SetFocus
   Exit Sub
End If
TxtBudgetIncomeWord.Text = NumStrConv(Val(TxtBudgetIncome.Text))
End Sub
Private Sub TxtBankInterest_LostFocus()
If TxtBankInterest = "" Then
   TxtBankInterestWord.Text = ""
   Exit Sub
End If
If IsNumeric(TxtBankInterest) = False Or Val(TxtBankInterest) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtBankInterest.SetFocus
   Exit Sub
End If
TxtBankInterestWord.Text = NumStrConv(Val(TxtBankInterest.Text))
End Sub

Private Sub TxtAOROCMGExpense_LostFocus()
If TxtAOROCMGExpense = "" Then
   TxtAOROCMGExpenseWord.Text = ""
   Exit Sub
End If
If IsNumeric(TxtAOROCMGExpense) = False Or Val(TxtAOROCMGExpense) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtAOROCMGExpense.SetFocus
   Exit Sub
End If
TxtAOROCMGExpenseWord.Text = NumStrConv(Val(TxtAOROCMGExpense.Text))
End Sub

