VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CashReceiptnotdepositedinBank 
   Caption         =   "Cash Receipt not deposited in Bank"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1965
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   15000
      Begin VB.OptionButton Option2 
         Caption         =   "Cash Receipt deposited in bank but not yet sighted by FAG"
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
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   5535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cash Receipt not deposited in bank"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.ComboBox CmbState 
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
         Left            =   7200
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   2955
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13560
         TabIndex        =   12
         Top             =   1410
         Width           =   1245
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13560
         TabIndex        =   11
         Top             =   1065
         Width           =   1245
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13560
         TabIndex        =   10
         Top             =   720
         Width           =   1245
      End
      Begin VB.ComboBox CmbLocation 
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
         ItemData        =   "FrmRpt_CashReceiptnotdepositedinBank.frx":0000
         Left            =   10290
         List            =   "FrmRpt_CashReceiptnotdepositedinBank.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   3045
      End
      Begin VB.ComboBox CmbDepositedBank 
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
         ItemData        =   "FrmRpt_CashReceiptnotdepositedinBank.frx":0004
         Left            =   1920
         List            =   "FrmRpt_CashReceiptnotdepositedinBank.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1425
         Width           =   11445
      End
      Begin MSComCtl2.DTPicker TxtCRTo 
         Height          =   360
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   635
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
         Format          =   64290817
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker TxtCRFrom 
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   635
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
         Format          =   64290817
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker TxtDTo 
         Height          =   360
         Left            =   5640
         TabIndex        =   6
         Top             =   960
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   635
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
         Format          =   64290817
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker TxtDFrom 
         Height          =   360
         Left            =   3840
         TabIndex        =   5
         Top             =   960
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   635
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
         Format          =   64290817
         CurrentDate     =   36494
      End
      Begin VB.Label Label2 
         Caption         =   "Total Record Found :-"
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
         Left            =   9720
         TabIndex        =   24
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lbltotalrecord 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12120
         TabIndex        =   23
         Top             =   247
         Width           =   90
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   8145
         TabIndex        =   22
         Top             =   645
         Width           =   1065
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
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
         Left            =   2265
         TabIndex        =   21
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   11430
         TabIndex        =   20
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Cash Receipt From"
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
         Left            =   90
         TabIndex        =   19
         Top             =   645
         Width           =   1725
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Deposited Bank"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Deposite From Date"
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
         Left            =   3645
         TabIndex        =   17
         Top             =   645
         Width           =   1815
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
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
         Left            =   5985
         TabIndex        =   16
         Top             =   645
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7980
      Left            =   120
      TabIndex        =   0
      Top             =   1935
      Width           =   15000
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   13
         Top             =   255
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   12515
         _Version        =   393216
         WordWrap        =   -1  'True
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   7440
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_CashReceiptnotdepositedinBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mBankID As Variant

Private Sub CmbLocation_GotFocus()
If CmbState.Text = "" Then
   MsgBox "Please Select State Name !!"
   'CmbLocation.SetFocus
   Exit Sub
End If

tmp = CmbLocation.Text
CmbLocation.Clear

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbDepositedBank_GotFocus()
tmp = CmbDepositedBank.Text
Call TableMst_FAGBank("")
CmbDepositedBank.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbDepositedBank.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
   CmbDepositedBank.AddItem RsMst_FAGBank!Bankname & "~" & RsMst_FAGBank!AccountNo
   RsMst_FAGBank.MoveNext
Next
CmbDepositedBank.Text = tmp
End Sub

Private Sub CmbDepositedBank_LostFocus()
Dim strSplitString() As String
Dim AccountNo_, BankName_ As String
If CmbDepositedBank.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbDepositedBank, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbDepositedBank.SetFocus
   Exit Sub
End If

BankName_ = Trim(strSplitString(0))
AccountNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo = '" & AccountNo_ & "'"

If RsMst_FAGBank.EOF Then
   MsgBox "Invalid Deposited Bank selection "
   CmbDepositedBank.SetFocus
   Exit Sub
End If
mBankID = RsMst_FAGBank!FAGBankId
End Sub

Private Sub CmbState_GotFocus()
tmp = CmbState.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub

Private Sub CmbState_LostFocus()
If tmp <> CmbState.Text Then
   CmbLocation.Text = ""
End If
If CmbState.Text = "" Then
   CmbLocation.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
TxtCRFrom.Value = Date - 365
TxtCRTo.Value = Date
TxtDFrom.Value = Date - 365
TxtDTo.Value = Date
CmbState.Text = ""
CmbLocation.Text = ""
Option1.Value = True
Call Option2_Click
Call Grid_Head
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant
If TxtCRFrom.Value > Date Then
   MsgBox "Future Cash Receipt Date not Allowed!!!"
   TxtCRFrom.SetFocus
   Exit Sub
End If
If TxtCRTo.Value > Date Then
   MsgBox "Future Cash Receipt Date not Allowed!!!"
   TxtCRTo.SetFocus
   Exit Sub
End If
If TxtCRFrom.Value > TxtCRTo.Value Then
   MsgBox "From Cash Receipt Date should not greater than To Cash Receipt Date!!!"
   TxtCRTo.SetFocus
   Exit Sub
End If

If Option2.Value = True Then
   If TxtDFrom.Value > Date Then
      MsgBox "Future Deposite Date not Allowed!!!"
      TxtDFrom.SetFocus
      Exit Sub
   End If
   If TxtDTo.Value > Date Then
      MsgBox "Future Deposite Date not Allowed!!!"
      TxtDTo.SetFocus
      Exit Sub
   End If
   If TxtDFrom.Value > TxtDTo.Value Then
      MsgBox "From Deposite Date should not greater than To Depsoite Date!!!"
      TxtDTo.SetFocus
      Exit Sub
   End If
End If

Call Grid_Head

wc = " Where CR.CRDate between '" & Format(TxtCRFrom.Value, "yyyy-mm-dd") & "' and '" & Format(TxtCRTo.Value, "yyyy-mm-dd") & "'"
wc = wc & " And isnull(ReturnFlag, '') <> 'Y' and isnull(RejectFlag, '') <> 'Y'"
If Option2.Value = True Then
   wc = wc & " And CR.DepositDate Between '" & Format(TxtDFrom.Value, "yyyy-mm-dd") & "' and '" & Format(TxtDTo.Value, "yyyy-mm-dd") & "'"
   wc = wc & " and CR.DepositDate is not null and CR.FAGPostingDate is null"
   If CmbDepositedBank.Text <> "" Then
      wc = wc & " And MFB.FAGBankId = " & mBankID
   End If
Else
   wc = wc & " and CR.DepositDate is null"
End If
If CmbLocation.Text <> "" Then
   wc = wc & " and ML.LocationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = wc & " and ML." & Gen_WcLocation
End If
If CmbState.Text <> "" Then
   wc = wc & " and MS.StateID = " & mStateID
ElseIf Gen_WcState <> "" Then
   wc = wc & " and MS." & Gen_WcState
End If

tmp = "Select MS.StateName, ML.LocationName, CR.CRID, CR.CRDate, CR.InstrumentType, CR.InstrumentAmount, CR.ReceivedFrom"
If Option2.Value = True Then
   tmp = tmp & " ,CR.DepositDate, MFB.BankName, MFB.AccountNo "
End If
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location ML On CR.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
If Option2.Value = True Then
   tmp = tmp & " Inner Join Mst_FAGBank MFB On CR.FAGBankId = MFB.FAGBankId"
End If
tmp = tmp & wc
tmp = tmp & " order by MS.StateName, ML.LocationName, CR.CRID, CR.CRDate"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   CmdExcel.Enabled = True
   lbltotalrecord.Caption = TmpRs.RecordCount
   ProgressBar1.Max = TmpRs.RecordCount + 2
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For I = 1 To TmpRs.RecordCount
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!CRID), "", TmpRs!CRID)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!CRDate), "", Format(TmpRs!CRDate, "dd-MMM-yyyy"))
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!ReceivedFrom), "", TmpRs!ReceivedFrom)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!InstrumentType), "", TmpRs!InstrumentType)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!InstrumentAmount), "", TmpRs!InstrumentAmount)
         If Option2.Value = True Then
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!Bankname), "", TmpRs!Bankname)
            .TextMatrix(I, 8) = IIf(IsNull(TmpRs!AccountNo), "", TmpRs!AccountNo)
            .TextMatrix(I, 9) = IIf(IsNull(TmpRs!DepositDate), "", Format(TmpRs!DepositDate, "dd-MMM-yyyy"))
         End If
         TmpRs.MoveNext
      Next
   End With
End If
MsgBox "Done!!!"
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CashReceiptnotdepositedinBank.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CashReceiptnotdepositedinBank.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1 ' + 6
   For c = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
   Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Grid_Head()
lbltotalrecord.Caption = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 7
   .FixedCols = 1
   .WordWrap = True
   .TextMatrix(0, 0) = "State Name"
   .TextMatrix(0, 1) = "Location Name"
   .TextMatrix(0, 2) = "CR No"
   .TextMatrix(0, 3) = "CR Date"
   .TextMatrix(0, 4) = "Received From"
   .TextMatrix(0, 5) = "Nature of Instrument"
   .TextMatrix(0, 6) = "Amount"
   If Option2.Value = True Then
      .Cols = 10
      .TextMatrix(0, 7) = "Deposited Bank"
      .TextMatrix(0, 8) = "Deposited Account No"
      .TextMatrix(0, 9) = "Deposited Date"
   End If
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 800
   .ColWidth(3) = 1200
   .ColWidth(4) = 5500
   .ColWidth(5) = 1500
   .ColWidth(6) = 1200
   
   .RowHeight(0) = 800
End With
End Sub

Private Sub Form_Load()
Call CmdClear_Click
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame0.Enabled = False
   Exit Sub
End If

Frame0.Enabled = True

End Sub

Private Sub Option1_Click()
   Call Option2_Click
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   CmbDepositedBank.Enabled = True
   TxtDFrom.Enabled = True
   TxtDTo.Enabled = True
ElseIf Option1.Value = True Then
   CmbDepositedBank.Text = ""
   TxtDFrom.Value = Date - 365
   TxtDTo.Enabled = Date
   CmbDepositedBank.Enabled = False
   TxtDFrom.Enabled = False
   TxtDTo.Enabled = False
End If
End Sub
