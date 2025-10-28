VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CRSightedDetails 
   Caption         =   "CR Sighted Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Cash/Cheque Receipts not yet Sighted/Deposited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6765
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   14970
      Begin MSComCtl2.DTPicker TxtSightedDate 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Format          =   48300035
         CurrentDate     =   40254
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4020
         TabIndex        =   17
         Top             =   300
         Width           =   495
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
         TabIndex        =   5
         Top             =   315
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5805
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   10239
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
      Begin VB.Label Label6 
         Caption         =   "Cash/Cheque Receipt not yet printed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4590
         TabIndex        =   18
         Top             =   315
         Width           =   3600
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   915
      Left            =   120
      TabIndex        =   15
      Top             =   8520
      Width           =   14970
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12660
         TabIndex        =   9
         Top             =   240
         Width           =   2145
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
         Height          =   525
         Left            =   10500
         TabIndex        =   8
         Top             =   240
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      Begin VB.TextBox TxtAmtTo 
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
         Left            =   13140
         MaxLength       =   12
         TabIndex        =   30
         Top             =   495
         Width           =   1605
      End
      Begin VB.TextBox TxtAmtFrom 
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
         Left            =   11595
         MaxLength       =   12
         TabIndex        =   28
         Top             =   495
         Width           =   1485
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10530
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1035
         Width           =   960
      End
      Begin VB.ComboBox CmbInstrumentType 
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
         ItemData        =   "FrmTxn_CRSightedDetails.frx":0000
         Left            =   7845
         List            =   "FrmTxn_CRSightedDetails.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   510
         Width           =   1860
      End
      Begin VB.TextBox TxtInstrumentNo 
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
         Left            =   9780
         MaxLength       =   50
         TabIndex        =   19
         Top             =   510
         Width           =   1725
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9720
         TabIndex        =   4
         Top             =   1133
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   3
         Top             =   1133
         Value           =   -1  'True
         Width           =   1215
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
         Left            =   7800
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1095
         Width           =   375
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
         Left            =   3795
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   510
         Width           =   3975
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   510
         Visible         =   0   'False
         Width           =   1635
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
         Left            =   1530
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   6255
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   345
         Left            =   2160
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
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
         Format          =   48300033
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   345
         Left            =   315
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   510
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
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
         Format          =   48300033
         CurrentDate     =   36494
      End
      Begin VB.Label Label7 
         Caption         =   "Amount To"
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
         Left            =   13440
         TabIndex        =   31
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "Amount From"
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
         Left            =   11760
         TabIndex        =   29
         Top             =   210
         Width           =   1320
      End
      Begin VB.Label Label44 
         Caption         =   "Deposit From Date"
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
         TabIndex        =   26
         Top             =   225
         Width           =   2055
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
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
         Left            =   2520
         TabIndex        =   25
         Top             =   225
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Instrument Type"
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
         Left            =   8040
         TabIndex        =   22
         Top             =   225
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Instrument No"
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
         Left            =   10080
         TabIndex        =   20
         Top             =   225
         Width           =   1320
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
         Left            =   5400
         TabIndex        =   14
         Top             =   225
         Width           =   1200
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
         Left            =   4320
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   720
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
         Left            =   195
         TabIndex        =   12
         Top             =   1095
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmTxn_CRSightedDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFAGBankID As Double
Dim mLocationID As Double
Dim Edit_Flg As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mSelRow, mSelCol As Variant
Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .Cols = 13
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
    .TextMatrix(1, 9) = "Received From"
    .TextMatrix(1, 10) = "Deposited Date"
    .TextMatrix(1, 11) = "CRDepositID"
    .TextMatrix(1, 12) = "Sighted Date"
    
    .ColAlignment(5) = vbLeftJustify
    .ColWidth(0) = 800
    .ColWidth(1) = 1600
    .ColWidth(2) = 0
    .ColWidth(3) = 700
    .ColWidth(4) = 1200
    .ColWidth(5) = 1000
    .ColWidth(6) = 1200
    .ColWidth(7) = 2500
    .ColWidth(8) = 1000
    .ColWidth(9) = 2000
    .ColWidth(10) = 1300
    .ColWidth(11) = 0
    .ColWidth(12) = 1200
    .RowHeight(1) = 600
    .WordWrap = True
End With
End Sub

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
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed Then
          If MSHFlexGrid1.TextMatrix(I, 0) <> strChecked Then
             MSHFlexGrid1.TextMatrix(I, 0) = strChecked
             MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8))
             If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 12) = "" Then MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 12) = Format(Date, "mm/dd/yyyy")
          End If
       End If
   Next
Else
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed Then
          MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
          MSHFlexGrid1.TextMatrix(0, 8) = 0 'Val(MSHFlexGrid1.TextMatrix(0, 8)) - Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8))
          MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 12) = ""
       End If
   Next
End If
End Sub

Private Sub CmdFAGBank_Click()
FrmMst_FAGBank.Show
End Sub

Private Sub CmdGo_Click()

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To Date' always!!!"
   SFrom.SetFocus
   Exit Sub
End If

If CmbBankName.Text = "" Then
   MsgBox "Blank Bank Name not Allowed"
   CmbBankName.SetFocus
   Exit Sub
End If

If Val(TxtAmtFrom) > Val(TxtAmtTo) Then
   MsgBox "'Amount From' should be less than 'Amount To'"
   TxtAmtFrom.SetFocus
   Exit Sub
End If
Call FillCRDetails
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   ButtonFrm.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
ButtonFrm.Enabled = True
Call ClearControls
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

mFAGBankID = RsMst_FAGBank!FAGBankId
'If tmp <> CmbBankName.Text Then
'   Call Grid_Head
'   If CmbLocationID.Text <> "" Then
'      Call FillCRDetails(mLocationID, mFAGBankID)
'   End If
'End If
End Sub

Private Sub FillCRDetails()
Dim Wc As String
Wc = ""

tmp = " Select ML.LocationName,TCR.SM_Prefix,TCR.CRID,TCR.CRDate,TCR.InstrumentType,TCR.InstrumentNo,TCR.DrawnOn + '~'+ TCR.BranchName ,TCR.InstrumentAmount,"
tmp = tmp & " TCR.ReceivedFrom,TN.DepositDate ,TCR.CRDepositID,TCR.SightedDate, TCR.Flag "
tmp = tmp & " From Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " left Join Txn_NewCashReceiptDepositDetails  TN On TN.CRDepositID = TCR.CRDepositID"

If Gen_DBType = "MDB" Then
   Wc = " Where ((TN.DepositDate Between #" & SFrom.Value & "# And #" & STo.Value & "#) or TN.DepositDate is null)"
Else
   Wc = " Where ((TN.DepositDate Between '" & SFrom.Value & "' And '" & STo.Value & "') or TN.DepositDate is null)"
End If
If CmbLocationID.Text <> "" Then
   Wc = Wc & " And TCR.LocationID =  " & mLocationID & ""
End If

If CmbInstrumentType.Text <> "" Then
    Wc = Wc & " And TCR.InstrumentType='" & Trim(CmbInstrumentType.Text) & "'"
End If

If TxtInstrumentNo.Text <> "" Then
    Wc = Wc & " And TCR.InstrumentNo Like '%" & Trim(TxtInstrumentNo.Text) & "%'"
End If
If CmbBankName.Text <> "" Then
   Wc = Wc & " And (TCR.CRDepositID is Null Or TCR.CRDepositID = 0 or TN.FAGBankID=" & mFAGBankID & ")"
End If
If TxtAmtFrom.Text <> "" And TxtAmtTo.Text <> "" Then
   Wc = Wc & " And (TCR.InstrumentAmount >=" & Val(TxtAmtFrom.Text) & " and TCR.InstrumentAmount <= " & Val(TxtAmtTo.Text) & ")"
End If
If Option1.Value = True Then
   Wc = Wc & " And TCR.SightedDate is Null"
End If
tmp = tmp & Wc & " Order By TCR.CRID,ML.LocationName"

Call TmpTable

Call Grid_Head
ChkCheckAll.Value = 0
If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 2
           MSHFlexGrid1.TextMatrix(x, c + 1) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(x, 4) = Format(MSHFlexGrid1.TextMatrix(x, 4), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(x, 10) = Format(MSHFlexGrid1.TextMatrix(x, 10), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(x, 8) = Format(Val(MSHFlexGrid1.TextMatrix(x, 8)), "#####0.00")
       
       MSHFlexGrid1.Row = x
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.TextMatrix(x, 0) = strUnChecked
       If LCase(TmpRs!Flag) <> "p" Or IsNull(TmpRs!Flag) = True Then
          'MSHFlexGrid1.Row = i
          For o = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.Col = o
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       End If
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       x = x + 1
       TmpRs.MoveNext
   Next
Else
   MsgBox "No Cash Receipt record found!!!"
End If
End Sub
Private Sub MSHFlexGrid1_Click()
Dim SelRow As Variant
If SaveCmd.Enabled = False Then Exit Sub
SelRow = MSHFlexGrid1.RowSel
With MSHFlexGrid1
    If .TextMatrix(.Row, 1) = "" Then Exit Sub
    If .Col = 0 Then
        If (.CellBackColor) <> vbRed Then
           If .TextMatrix(.Row, 0) = strUnChecked Then
              .TextMatrix(.Row, 0) = strChecked
              .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.Row, 8))
              If .TextMatrix(.Row, 12) = "" Then .TextMatrix(.Row, 12) = Format(Date, "mm/dd/yyyy")
           Else
              .TextMatrix(.Row, 0) = strUnChecked
              .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.Row, 8))
              .TextMatrix(.Row, 12) = ""
           End If
        End If
    End If
End With
If MSHFlexGrid1.TextMatrix(SelRow, 0) = strChecked Then
   If MSHFlexGrid1.Row = SelRow And MSHFlexGrid1.Col = 12 Then
      mSelRow = MSHFlexGrid1.Row
      mSelCol = MSHFlexGrid1.Col
      TxtSightedDate.Visible = True
      
      With TxtSightedDate
          .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
          .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
          .Width = MSHFlexGrid1.CellWidth - 10
          .Height = MSHFlexGrid1.CellHeight - 10
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
End Sub
'
'Private Sub Option1_Click()
'If CmbLocationID.Text <> "" And CmbBankName.Text <> "" Then
'   Call FillCRDetails(mLocationID, mFAGBankID)
'End If
'End Sub
'
'Private Sub Option2_Click()
'If CmbLocationID.Text <> "" And CmbBankName.Text <> "" Then
'   Call FillCRDetails(mLocationID, mFAGBankID)
'End If
'
'End Sub

Private Sub SaveCmd_Click()
Dim mCRDepID As Variant
'If CmbLocationID = "" Then
'   MsgBox "Blank Location Name not Allowed"
'   CmbLocationID.SetFocus
'   Exit Sub
'End If

If CmbBankName = "" Then
   MsgBox "Blank Bank Name not Allowed"
   CmbBankName.SetFocus
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
Call TableTxn_NewCashReceipt("")
Call TableTxn_NewCRDepositDetails("")
For I = 2 To MSHFlexGrid1.Rows - 2
    If RsTxn_NewCashReceipt.RecordCount > 0 Then RsTxn_NewCashReceipt.MoveFirst
    RsTxn_NewCashReceipt.Filter = "CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 3)) & " And SM_Prefix='" & MSHFlexGrid1.TextMatrix(I, 2) & "'"
    If RsTxn_NewCashReceipt.RecordCount > 0 Then
       If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And MSHFlexGrid1.TextMatrix(I, 12) <> "" Then
          RsTxn_NewCashReceipt!SightedDate = CDate(MSHFlexGrid1.TextMatrix(I, 12))
          If MSHFlexGrid1.TextMatrix(I, 10) = "" Then
             mCRDepID = AddDepositeDetails()
             RsTxn_NewCashReceipt!CRDepositID = mCRDepID
          End If
       End If
       RsTxn_NewCashReceipt.Update
       RsTxn_NewCashReceipt.Requery
    End If
    RsTxn_NewCashReceipt.Filter = ""
Next
MsgBox "Done!!"
Call ClearControls
End Sub
Private Function AddDepositeDetails() As Double
Dim Retval As Double

RsTxn_NewCRDepositDetails.AddNew
RsTxn_NewCRDepositDetails!CRDepositID = GetMaxCRDepositID
Retval = RsTxn_NewCRDepositDetails!CRDepositID
RsTxn_NewCRDepositDetails!G_UID = GetGUID
RsTxn_NewCRDepositDetails!LocationID = RsTxn_NewCashReceipt!LocationID
RsTxn_NewCRDepositDetails!FAGBankId = mFAGBankID
RsTxn_NewCRDepositDetails!DepositDate = RsTxn_NewCashReceipt!SightedDate
RsTxn_NewCRDepositDetails!Remark = "Created via Sighted Module."
RsTxn_NewCRDepositDetails!CreatedBy = Gen_UserLoginID
RsTxn_NewCRDepositDetails!CreatedDate = Now
RsTxn_NewCRDepositDetails.Update
RsTxn_NewCRDepositDetails.Requery

AddDepositeDetails = Retval
End Function
Private Function GetMaxCRDepositID() As Double
Dim Retval As Double
tmp = "Select max(CRDepositID) from Txn_NewCashReceiptDepositDetails"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxCRDepositID = Retval
End Function

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
'If tmp <> CmbLocationID.Text Then
'   Call Grid_Head
'   If CmbBankName.Text <> "" Then
'      Call FillCRDetails(mLocationID, mFAGBankID)
'   End If
'End If
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")

End Sub
Private Sub TxtSightedDate_LostFocus()
If TxtSightedDate.Value > Date Then
   MsgBox "Future date not allowed!!"
   TxtSightedDate.SetFocus
   Exit Sub
End If
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = TxtSightedDate.Value
TxtSightedDate.Visible = False
End Sub

Private Sub ClearControls()
CmbLocationID.Text = ""
CmbBankName.Text = ""
Option1.Value = True
Call Grid_Head
ChkCheckAll.Value = 0
TxtInstrumentNo = ""
CmbInstrumentType.Text = ""
SFrom.Value = Date - 1
STo.Value = Date
TxtAmtFrom = ""
TxtAmtTo = ""
End Sub

Private Sub CmbInstrumentType_LostFocus()
If CmbInstrumentType.Text = "" Then Exit Sub
CmbInstrumentType.Text = Trim(CmbInstrumentType.Text)
If LCase(CmbInstrumentType.Text) = "cash" Then
   TxtInstrumentNo.Text = ""
   TxtInstrumentNo.Enabled = False
   Exit Sub
End If
If LCase(CmbInstrumentType.Text) = "cheque" Or LCase(CmbInstrumentType.Text) = "dd" Or LCase(CmbInstrumentType.Text) = "rtgs" Then
   TxtInstrumentNo.Enabled = True
   Exit Sub
End If
MsgBox "Invalid selection !!!!"
CmbInstrumentType.SetFocus
End Sub

