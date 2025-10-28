VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_NewCRDepositeDetails 
   Caption         =   "CR Deposit Details"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10500
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   630
      Left            =   120
      TabIndex        =   18
      Top             =   8565
      Width           =   14985
      Begin VB.CommandButton ExitCmd 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   12990
         TabIndex        =   10
         Top             =   165
         Width           =   1740
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11220
         TabIndex        =   9
         Top             =   165
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   14985
      Begin VB.TextBox TxtSearchBy 
         Appearance      =   0  'Flat
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
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1845
         Width           =   3885
      End
      Begin VB.ComboBox CmbSearchBy 
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
         ItemData        =   "FrmTxn_NewCRDepositeDetails.frx":0000
         Left            =   2040
         List            =   "FrmTxn_NewCRDepositeDetails.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1845
         Width           =   3030
      End
      Begin VB.TextBox TxtRemark 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1155
         Width           =   12030
      End
      Begin VB.ComboBox CmbBankName 
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   705
         Width           =   12030
      End
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   12030
      End
      Begin VB.CommandButton CmdFAGBank 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14205
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   735
         Width           =   375
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
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
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "Search Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   20
         Top             =   1875
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   1875
         Width           =   1440
      End
      Begin VB.Label Label9 
         Caption         =   "Remark / Reason for delay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   17
         Top             =   1230
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   758
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Prefix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11490
         TabIndex        =   15
         Top             =   300
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label8 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cash/Cheque Receipts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6165
      Left            =   120
      TabIndex        =   6
      Top             =   2415
      Width           =   14970
      Begin MSComCtl2.DTPicker TxtFAGPostingDate 
         Height          =   330
         Left            =   2280
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   107675651
         CurrentDate     =   40254
      End
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   315
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5445
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   9604
         _Version        =   393216
         RowHeightMin    =   370
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmTxn_NewCRDepositeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFAGBankID, mSFAGBankID As Variant
Dim mLocationID As Double
Dim Edit_Flg As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mSelRow, mSelCol As Variant
Private Sub ChkCheckAll_Click()
If SaveCmd.Enabled = False Then
   ChkCheckAll.Value = 0
   Exit Sub
End If
If MSHFlexGrid1.Rows = 3 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed Then
          If MSHFlexGrid1.TextMatrix(I, 0) <> strChecked Then
             MSHFlexGrid1.TextMatrix(I, 0) = strChecked
             MSHFlexGrid1.TextMatrix(I, 10) = Date
             MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))
          End If
       End If
   Next
Else
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed Then
          MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
          MSHFlexGrid1.TextMatrix(0, 8) = 0
          MSHFlexGrid1.TextMatrix(I, 10) = ""
       End If
   Next
End If



End Sub

Private Sub CmbSearchBy_LostFocus()

If CmbSearchBy.Text = "" Then
   TxtSearchBy = ""
   TxtSearchBy.Locked = True
   Exit Sub
End If

If LCase(CmbSearchBy.Text) <> "crid" And LCase(CmbSearchBy.Text) <> "amount" And LCase(CmbSearchBy.Text) <> "instrument type" And LCase(CmbSearchBy.Text) <> "instrument no" Then
   MsgBox "Invalid Selection!!!"
   CmbSearchBy.SetFocus
   Exit Sub
End If
TxtSearchBy.Locked = False
TxtSearchBy = ""

End Sub

Private Sub CmdFAGBank_Click()
FrmMst_FAGBank.Show
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
'   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

'If Gen_WcLocation <> "" Then
'   Call TableTxn_NewCRDepositDetails(" Where  " & Gen_WcLocation & " Order by CRDepositID ")
'Else
'   Call TableTxn_NewCRDepositDetails(" Order by CRDepositID ")
'End If
'
'If RsTxn_NewCRDepositDetails.RecordCount = 0 Then
'   Call AddCmd_Click
'   Exit Sub
'End If
'Call Indata
'Call GButtonLock(Me, True)

Call ClearData

End Sub
Private Sub ClearData()

CmbLocationID.Text = ""
CmbBankName.Text = ""
Call Grid_Head
TxtRemark = ""
TxtFAGPostingDate.Visible = False

End Sub
Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_FAGBank("")
CmbBankName.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbBankName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbBankName.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbBankName.Text = tmp
End Sub
Private Sub CmbBankName_LostFocus()
Dim strSplitString() As String
Dim AccountNo_, BankName_ As String

If CmbBankName.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbBankName, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbBankName.SetFocus
   Exit Sub
End If

BankName_ = Trim(strSplitString(0))
AccountNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo = '" & AccountNo_ & "'"

If RsMst_FAGBank.EOF Then
   MsgBox "Invalid FAG Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If

'If tmp = "" Then
'   If CmbLocationID.Text = "" Then
If CmbBankName.Text = "" Then Exit Sub
mFAGBankID = RsMst_FAGBank!FAGBankID
If tmp <> CmbBankName.Text Then
   Call Grid_Head
   Call FillCRDetails '(mLocationID)
End If
   'End If
'End If



End Sub
Private Sub FillCRDetails() '(mLocationID_ As Double, Optional CRDepositeID_ As Double)

tmp = " Select ML.LocationName,TCR.SM_Prefix,TCR.CRID,TCR.CRDate,TCR.InstrumentType,TCR.InstrumentNo, TCR.DrawnOn + '~'+ TCR.BranchName ,TCR.InstrumentAmount, "
tmp = tmp & " TCR.ReceivedFrom , TCR.DepositDate,TCR.FAGPostingDate,TCR.Flag "
tmp = tmp & " From Txn_NewCashReceipt TCR "
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID "
tmp = tmp & " Where (TCR.RejectFlag is Null Or TCR.RejectFlag = 'N')  And  TCR.BatchID = 0 " 'TCR.SM_prefix = '" & TxtPreFix & "' And  " 'And (CRDepositID is Null Or CRDepositID = 0 "

If CmbBankName.Text <> "" Then
   tmp = tmp & " And (TCR.FAGBankID = " & mFAGBankID & " or TCR.FAGBankID is null)"
End If

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And TCR.SM_prefix = '" & TxtPreFix & "' "
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & "ML." & Gen_WcLocation
   End If
End If

tmp = tmp & " Order By TCR.SM_Prefix,TCR.CRID "

Call TmpTable

Call Grid_Head
ChkCheckAll.Value = 0
If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       For C = 0 To TmpRs.Fields.Count - 2
           MSHFlexGrid1.TextMatrix(x, C + 1) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(x, 4) = Format(MSHFlexGrid1.TextMatrix(x, 4), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(x, 8) = Format(Val(MSHFlexGrid1.TextMatrix(x, 8)), "#####0.00")
       
       MSHFlexGrid1.row = x
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.TextMatrix(x, 0) = strUnChecked
       If IsNull(TmpRs!DepositDate) = False Then
          MSHFlexGrid1.TextMatrix(x, 10) = Format(MSHFlexGrid1.TextMatrix(x, 10), "dd-MMM-yyyy")
       Else
          MSHFlexGrid1.TextMatrix(x, 10) = "" 'Format(MSHFlexGrid1.TextMatrix(X, 10), "dd-MMM-yyyy")
       End If
       If IsNull(TmpRs!FAGPostingDate) = False Then
          'MSHFlexGrid1.TextMatrix(X, 0) = strChecked
          MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 8))
          MSHFlexGrid1.TextMatrix(x, 11) = Format(MSHFlexGrid1.TextMatrix(x, 11), "dd-MMM-yyyy")
       Else
          MSHFlexGrid1.TextMatrix(x, 11) = ""
       End If
'       If LCase(TmpRs!Flag) <> "p" Or IsNull(TmpRs!Flag) = True Then
'          'MSHFlexGrid1.Row = i
'          For o = 0 To MSHFlexGrid1.Cols - 1
'              MSHFlexGrid1.Col = o
'              MSHFlexGrid1.CellBackColor = vbRed
'          Next
'       End If
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       x = x + 1
       TmpRs.MoveNext
   Next
Else
   MsgBox "No Cash Receipt record found!!!"
End If

End Sub
Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .Cols = 12
    .Font.Size = 10
    .FixedCols = 0
    .FixedRows = 2
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Select"
    .TextMatrix(1, 1) = "Location"
    .TextMatrix(1, 2) = "SM_Prefix"
    .TextMatrix(1, 3) = "CR ID"
    .TextMatrix(1, 4) = "CR Date"
    .TextMatrix(1, 5) = "Instrument Type"
    .TextMatrix(1, 6) = "Instrument No"
    .TextMatrix(1, 7) = "Drawn On"
    .TextMatrix(1, 8) = "Amount"
    .TextMatrix(1, 9) = "Receveid From"
    .TextMatrix(1, 10) = "Deposit Date"
    .TextMatrix(1, 11) = "FAG Posting Date"
        
    .ColAlignment(5) = vbLeftJustify
    .ColWidth(0) = 800
    .ColWidth(1) = 1800
    .ColWidth(2) = 0
    .ColWidth(3) = 800
    .ColWidth(4) = 1200
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 3200
    .ColWidth(8) = 1200
    .ColWidth(9) = 3200
    .ColWidth(10) = 2400
    .ColWidth(11) = 2400
    .WordWrap = True
    .RowHeight(1) = 600
End With
End Sub

Private Sub MSHFlexGrid1_Click()
Dim SelRow As Variant
SelRow = MSHFlexGrid1.RowSel
'If SaveCmd.Enabled = False Then Exit Sub
'With MSHFlexGrid1
'    If .TextMatrix(.Row, 2) = "" Then Exit Sub
'    If .Col = 0 Then
'        If (.CellBackColor) <> vbRed Then
'           If .TextMatrix(.Row, 0) = strUnChecked Then
'              .TextMatrix(.Row, 0) = strChecked
'               .TextMatrix(.Row, 10) = Date
'              .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.Row, 8))
'           Else
'              .TextMatrix(.Row, 0) = strUnChecked
'              .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.Row, 8))
'              .TextMatrix(0, 10) = ""
'           End If
'        End If
'    End If
'End With

If LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "power" Then

   With MSHFlexGrid1
       If .TextMatrix(.row, 2) = "" Then Exit Sub
       If .Col = 0 Then
           If (.CellBackColor) <> vbRed Then
              If .TextMatrix(.row, 0) = strUnChecked Then
                 .TextMatrix(.row, 0) = strChecked
                 .TextMatrix(.row, 10) = Date
                 .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.row, 8))
              Else
                 .TextMatrix(.row, 0) = strUnChecked
                 .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.row, 8))
                 .TextMatrix(0, 10) = ""
              End If
           End If
       End If
   End With

   If MSHFlexGrid1.TextMatrix(SelRow, 0) = strChecked Then
      If MSHFlexGrid1.row = SelRow And MSHFlexGrid1.Col = 10 Then
         mSelRow = MSHFlexGrid1.row
         mSelCol = MSHFlexGrid1.Col
         TxtFAGPostingDate.Visible = True
         
         With TxtFAGPostingDate
             .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
             .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
             .Width = MSHFlexGrid1.CellWidth '- 10
             .Height = MSHFlexGrid1.CellHeight ' - 10
             If MSHFlexGrid1.Text <> "" Then
                .Value = CDate(MSHFlexGrid1.Text)
             Else
                .Value = Date
             End If
             .ZOrder
             .SetFocus
         End With
      End If
   End If
Else
   With MSHFlexGrid1
       If .TextMatrix(.row, 2) = "" Then Exit Sub
       If .Col = 0 Then
           If (.CellBackColor) <> vbRed Then
              If .TextMatrix(.row, 0) = strUnChecked Then
                 .TextMatrix(.row, 0) = strChecked
                 If .TextMatrix(.row, 10) = "" Then
                    .TextMatrix(.row, 10) = Date
                 End If
                 .TextMatrix(.row, 11) = Date
                 '.TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.Row, 8))
              Else
                 .TextMatrix(.row, 0) = strUnChecked
                 '.TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.Row, 8))
                 If CDate(.TextMatrix(.row, 10)) = Date Then
                    .TextMatrix(.row, 10) = ""
                 End If
                 .TextMatrix(0, 11) = ""
              End If
           End If
       End If
   End With
   
   
   If MSHFlexGrid1.TextMatrix(SelRow, 0) = strChecked Then
      If MSHFlexGrid1.row = SelRow And (MSHFlexGrid1.Col = 10 Or MSHFlexGrid1.Col = 11) Then
         mSelRow = MSHFlexGrid1.row
         mSelCol = MSHFlexGrid1.Col
         TxtFAGPostingDate.Visible = True
         With TxtFAGPostingDate
             .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
             .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
             .Width = MSHFlexGrid1.CellWidth '- 10
             .Height = MSHFlexGrid1.CellHeight ' - 10
             If MSHFlexGrid1.Text <> "" Then
                .Value = CDate(MSHFlexGrid1.Text)
             Else
                .Value = Date
             End If
             .ZOrder
             .SetFocus
         End With
      End If
   End If
End If

End Sub

Private Sub SaveCmd_Click()

If SaveCmd.Enabled = False Then
   Exit Sub
End If


'If CmbLocationID.Text = "" Then
'   MsgBox "Blank Location not Allowed"
'   CmbLocationID.SetFocus
'   Exit Sub
'End If

If CmbBankName.Text = "" Then
   MsgBox "Blank Bank Name not Allowed"
   CmbBankName.SetFocus
   Exit Sub
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not Allowed"
   TxtRemark.SetFocus
   Exit Sub
End If

If MSHFlexGrid1.Rows < 4 Then
   MsgBox "Blank Cash Receipt Details not Allowed"
   Exit Sub
End If

If CheckCRSelect = False Then
   MsgBox "Select atleast one Cash Receipt ID from CR Details!!!"
   Exit Sub
End If

'If CheckCRDate = True Then
'   MsgBox "'Date of Deposit' should be greater than or equal to CR Date!!"
'   TxtDepositeDate.SetFocus
'   Exit Sub
'End If

SaveCmd.Enabled = False

Call UpdateCRDetails

Call ClearData
SaveCmd.Enabled = True

End Sub
Private Sub UpdateCRDetails()
For I = 2 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       'If MSHFlexGrid1.TextMatrix(I, 10) <> "" Then
          If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
             If MSHFlexGrid1.TextMatrix(I, 11) <> "" Then
'                If CDate(MSHFlexGrid1.TextMatrix(I, 10)) < CDate(MSHFlexGrid1.TextMatrix(I, 4)) Then
'                    MsgBox "Deposite Date should not less than Cash Reeceipt Date!!!"
'                    Exit Sub
'                End If
'                If CDate(MSHFlexGrid1.TextMatrix(I, 11)) < CDate(MSHFlexGrid1.TextMatrix(I, 10)) Then
'                    MsgBox "FAG Posting Date should not less than Deposite Date!!!"
'                    Exit Sub
'                End If
                tmp = " Update Txn_NewCashReceipt Set FAGPostingDate = '" & CDate(MSHFlexGrid1.TextMatrix(I, 11)) & "',DepositDate = '" & CDate(MSHFlexGrid1.TextMatrix(I, 10)) & "',FAGBankID = " & mFAGBankID & ",FAGRemarks = '" & Trim(TxtRemark) & "',Flag = 'P',HologramNo = 0,FAGUpdatedBy = '" & Gen_UserLoginID & "',FAGUpdatedDate = getDate()"
                tmp = tmp & " Where CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 3)) & " And SM_Prefix= '" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "'"
                Call TmpTable
             End If
          Else
             If MSHFlexGrid1.TextMatrix(I, 10) <> "" Then
'                If CDate(MSHFlexGrid1.TextMatrix(I, 10)) < CDate(MSHFlexGrid1.TextMatrix(I, 4)) Then
'                    MsgBox "Deposite Date should not less than Cash Reeceipt Date!!!"
'                    Exit Sub
'                End If
                tmp = " Update Txn_NewCashReceipt Set DepositDate = '" & CDate(MSHFlexGrid1.TextMatrix(I, 10)) & "',FAGBankID = " & mFAGBankID & ",FAGRemarks = '" & Trim(TxtRemark) & "',Flag = 'P',HologramNo = 0,UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getDate()"
                tmp = tmp & " Where CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 3)) & " And SM_Prefix= '" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "'"
                Call TmpTable
             End If
          End If
          
       'End If
    Else
       If LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Then
          If MSHFlexGrid1.TextMatrix(I, 11) = "" Then
             tmp = " Update Txn_NewCashReceipt Set FAGPostingDate = Null,FAGBankID = Null,FAGRemarks = Null,Flag = Null,HologramNo = Null,FAGUpdatedBy = '" & Gen_UserLoginID & "',FAGUpdatedDate = getDate()"
             tmp = tmp & " Where CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 3)) & " And SM_Prefix= '" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "'"
             Call TmpTable
          End If
       Else
          If MSHFlexGrid1.TextMatrix(I, 11) = "" And MSHFlexGrid1.TextMatrix(I, 10) = "" Then
             tmp = " Update Txn_NewCashReceipt Set FAGPostingDate = Null ,DepositDate = Null,FAGBankID = Null,FAGRemarks = Null,Flag = Null,HologramNo = Null,UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getDate()"
             tmp = tmp & " Where CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 3)) & " And SM_Prefix= '" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "'"
             Call TmpTable
          End If
       End If

    End If
Next
End Sub
Private Function CheckCRSelect() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckCRSelect = Retval
End Function

Private Function CheckCRDate() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If CDate(MSHFlexGrid1.TextMatrix(I, 4)) > TxtDepositeDate.Value Then
          Retval = True
          Exit For
       End If
    End If
Next
CheckCRDate = Retval
End Function

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

'Call TableMst_Location(" Where LocationId > 0 ")

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
If CmbBankName.Text = "" Then Exit Sub
If tmp <> CmbLocationID.Text Then
   Call Grid_Head
   Call FillCRDetails '(mLocationID)
End If

End Sub
Private Sub TxtFAGPostingDate_LostFocus()

If TxtFAGPostingDate.Value > Date Then
   MsgBox "Future date not allowed!!"
   TxtFAGPostingDate.SetFocus
   Exit Sub
End If
If mSelCol = 10 Then
    If TxtFAGPostingDate.Value < CDate(MSHFlexGrid1.TextMatrix(mSelRow, 4)) Then
        MsgBox "Deposit Date should not less than Cash Reeceipt Date!!!"
        TxtFAGPostingDate.SetFocus
        Exit Sub
    End If

    If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFAGPostingDate.Value, "yyyy-mm") Then
       MsgBox "Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
       TxtFAGPostingDate.SetFocus
       Exit Sub
    End If

End If

If mSelCol = 11 Then
    If MSHFlexGrid1.TextMatrix(mSelRow, 10) <> "" Then
        If TxtFAGPostingDate.Value < CDate(MSHFlexGrid1.TextMatrix(mSelRow, 10)) Then
            MsgBox "FAG Posting Date should not less than Deposite Date!!!"
            TxtFAGPostingDate.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Blank Deposite Date not Allowed!!!"
        'TxtFAGPostingDate.SetFocus
        Exit Sub
    End If
End If

MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = TxtFAGPostingDate.Value
TxtFAGPostingDate.Visible = False

End Sub

Private Sub TxtSearchBy_LostFocus()

If CmbSearchBy.Text = "" Then Exit Sub

If TxtSearchBy = "" Then
   MSHFlexGrid1.TopRow = 2
   Exit Sub
End If


For I = 2 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.row = I
    If LCase(CmbSearchBy.Text) = "crid" Then
       MSHFlexGrid1.Col = 3
       If Val(TxtSearchBy) = Val(MSHFlexGrid1.TextMatrix(I, 3)) Then
          MSHFlexGrid1.TopRow = I
          Exit Sub
       End If
    ElseIf LCase(CmbSearchBy.Text) = "amount" Then
       MSHFlexGrid1.Col = 8
       If Val(TxtSearchBy) = Val(MSHFlexGrid1.TextMatrix(I, 8)) Then
          MSHFlexGrid1.TopRow = I
          Exit Sub
       End If
    ElseIf LCase(CmbSearchBy.Text) = "instrument type" Then
       MSHFlexGrid1.Col = 5
       If LCase(TxtSearchBy) = LCase(MSHFlexGrid1.TextMatrix(I, 5)) Then
          MSHFlexGrid1.TopRow = I
          Exit Sub
       End If
    ElseIf LCase(CmbSearchBy.Text) = "instrument no" Then
       MSHFlexGrid1.Col = 6
       If Trim(TxtSearchBy) = Trim(MSHFlexGrid1.TextMatrix(I, 6)) Then
          MSHFlexGrid1.TopRow = I
          Exit Sub
       End If
    End If
Next

End Sub
