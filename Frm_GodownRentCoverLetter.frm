VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_GodownRentCoverLetter 
   Caption         =   "Godown Rent Cover Letter"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   9900
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.TextBox TxtChequeNo 
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
         Height          =   375
         Left            =   2085
         TabIndex        =   5
         Top             =   800
         Width           =   1815
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
         Left            =   10380
         TabIndex        =   19
         Top             =   840
         Width           =   1140
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
         Left            =   7200
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   375
         Width           =   4995
      End
      Begin VB.ComboBox CmbSLessor 
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
         Left            =   2085
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   375
         Width           =   4830
      End
      Begin VB.ComboBox CmbBatch 
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
         Left            =   12480
         TabIndex        =   4
         Top             =   375
         Width           =   2100
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
         Left            =   9240
         TabIndex        =   6
         Top             =   840
         Width           =   1140
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
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
         Left            =   11520
         TabIndex        =   7
         Top             =   840
         Width           =   1140
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
         Left            =   12660
         TabIndex        =   8
         Top             =   840
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
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
         Height          =   350
         Left            =   13800
         TabIndex        =   9
         Top             =   840
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   360
         Left            =   75
         TabIndex        =   1
         Top             =   375
         Width           =   1860
         _ExtentX        =   3281
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
         Format          =   9240577
         CurrentDate     =   36494
      End
      Begin VB.Label Label2 
         Caption         =   "Cheque No"
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
         TabIndex        =   20
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label44 
         Caption         =   "Rent Entered From"
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
         TabIndex        =   18
         Top             =   120
         Width           =   1770
      End
      Begin VB.Label Label5 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9135
         TabIndex        =   17
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label37 
         Caption         =   "Lessor Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3870
         TabIndex        =   14
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No"
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
         Left            =   13088
         TabIndex        =   13
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   120
      TabIndex        =   10
      Top             =   1230
      Width           =   15015
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   11040
         TabIndex        =   15
         Top             =   240
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
         TabIndex        =   11
         Top             =   250
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7440
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13123
         _Version        =   393216
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
      Begin VB.Label Label4 
         Caption         =   "Rent Payment Details not Uploaded"
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
         Left            =   11610
         TabIndex        =   16
         Top             =   255
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Frm_GodownRentCoverLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpFlag, MaxID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mMonth As Variant
Dim mBatchID As Variant
Dim mSMonth, mSYear, mSTotalDays As Variant
Dim mPrintOut As Variant
Dim mChequeNoFlag As Boolean

Private Sub ChkCheckAll_Click()
If MSHFlexGrid3.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       MSHFlexGrid3.Col = 0
       If MSHFlexGrid3.CellBackColor <> vbRed Then
          MSHFlexGrid3.row = I
          MSHFlexGrid3.TextMatrix(I, 0) = strChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       MSHFlexGrid3.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbBankName_GotFocus()
Dim tmp1 As Variant
tmp1 = CmbBankName.Text
tmp = "Select Distinct BankName from Txn_GodownRent Where BankName is not Null"
Call TmpTable

CmbBankName.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbBankName.AddItem TmpRs!BankName
    TmpRs.MoveNext
Next
CmbBankName.Text = tmp1

End Sub

Private Sub CmbBankName_LostFocus()

If CmbBankName.Text = "" Then
   Exit Sub
End If
TmpRs.MoveFirst
TmpRs.Find "BankName = '" & CmbBankName.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid selection "
   CmbBankName.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbBatch_GotFocus()
tmp = "select distinct BatchID As BatchID from Txn_GodownRent Where BatchID > 0  order by BatchID"
Call Tmp1Table
CmbBatch.Clear
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       CmbBatch.AddItem TmpRs1.Fields(0)
       MaxID = TmpRs1.Fields(0) + 1
       TmpRs1.MoveNext
   Next
Else
   MaxID = 1
End If
End Sub
Public Sub Grid_Head()
With MSHFlexGrid3

.Clear
.Rows = 2
.FixedRows = 1

.Cols = 20
.FixedCols = 0
.TextMatrix(0, 0) = "Select"
.TextMatrix(0, 1) = "Claim ID"
.TextMatrix(0, 2) = "Claim Date"
.TextMatrix(0, 3) = "Expense From"
.TextMatrix(0, 4) = "Expense To"
.TextMatrix(0, 5) = "Lessor ID"
.TextMatrix(0, 6) = "Lessor Name"
.TextMatrix(0, 7) = "Rent Amount"
.TextMatrix(0, 8) = "Claim Forwarded To"
.TextMatrix(0, 9) = "Status"
.TextMatrix(0, 10) = "Booked in NAV Date"
.TextMatrix(0, 11) = "FAG Remarks"
.TextMatrix(0, 12) = "Remarks"
.TextMatrix(0, 13) = "Batch No"
.TextMatrix(0, 14) = "G_UID"
.TextMatrix(0, 15) = "Bank Name"
.TextMatrix(0, 16) = "Cheque No"
.TextMatrix(0, 17) = "Paid Amount"
.TextMatrix(0, 18) = "TDS "
.TextMatrix(0, 19) = "Paid Date"
.ColWidth(0) = 800
.ColWidth(1) = 600
.ColWidth(2) = 1400
.ColWidth(3) = 1400
.ColWidth(4) = 1400
.ColWidth(5) = 0
.ColWidth(6) = 3000
.ColWidth(7) = 1400
.ColWidth(8) = 3000
.ColWidth(9) = 1000
.ColWidth(10) = 1400
.ColWidth(11) = 2600
.ColWidth(12) = 2600
.ColWidth(13) = 1200
.ColWidth(14) = 0
.ColWidth(15) = 2000
.ColWidth(16) = 1800
.ColWidth(17) = 1400
.ColWidth(18) = 1400
.ColWidth(19) = 1400
.RowHeight(0) = 500

  
End With

End Sub
Private Sub CmbBatch_LostFocus()
If CmbBatch.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "BatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbBatch.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()

mSMonth = Month(Date) - 1
mSYear = Year(Date)
mSTotalDays = Day(DateSerial(mYear, mSMonth + 1, 0))
If mSMonth = 0 Then
   mSMonth = 1
End If
SFrom.Value = CDate(mSYear & "-" & mSMonth & "-" & "01")

CmbSLessor.Text = ""
CmbBankName.Text = ""
CmbBatch.Text = ""
Call Grid_Head
mBatchID = ""
ChkCheckAll.Value = vbUnchecked
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_UserMaster.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_UserMaster.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid3.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant
Call Grid_Head
tmp = "Select TGR.ClaimID,TGR.ClaimDate,TGR.ExpenseFrom,TGR.ExpenseTo, "
tmp = tmp & "  ML.LessorID,ML.LessorName,TGR.Amount,EMP.EmployeeName +'~'+ EMP.EmployeeNo as ForwardedTo ,TGR.Flag, "
tmp = tmp & "  TGR.ApprovalDate,TGR.ApprovalRemarks, "
tmp = tmp & "  TGR.Remarks,TGR.BatchID,TGR.G_UID,TGR.BankName,TGR.ChequeNo,TGR.PaidAmount,TGR.TDS,TGR.PaidDate "
tmp = tmp & "  From Txn_GodownRent TGR "
tmp = tmp & "  Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID "
tmp = tmp & "  Left Outer Join Mst_Employee EMP On TGR.ForwardToID = EMP.EmployeeID "
tmp = tmp & "  Where TGR.Flag in ('S', 'A') "

If Trim(TxtChequeNo) <> "" Then
    mChequeNoFlag = True
    wc = " And TGR.ChequeNo = '" & Trim(TxtChequeNo) & "'"
Else
    mChequeNoFlag = False
    If Gen_DBType = "MDB" Then
       wc = " And Month(ExpenseFrom) = " & Month(SFrom) & " And Year(ExpenseFrom) = " & Year(SFrom) & ""
    Else
       wc = " And Month(ExpenseFrom) = " & Month(SFrom) & " And Year(ExpenseFrom) = " & Year(SFrom) & ""
    End If
    If CmbSLessor.Text <> "" Then
       wc = wc & " AND ML.LessorName = '" & CmbSLessor.Text & "' "
    End If
    If CmbBankName.Text <> "" Then
       wc = wc & " AND TGR.BankName = '" & CmbBankName.Text & "' "
    End If
    If CmbBatch.Text <> "" Then
       wc = wc & " And TGR.BatchID = " & Val(CmbBatch.Text) & ""
    End If
End If

tmp = tmp & wc
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mBatchID = CmbBatch.Text
   For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
    For c = 0 To MSHFlexGrid3.Cols - 2
        If InStr(TmpRs.Fields(c), "/") > 0 Then
           If IsDate(TmpRs.Fields(c)) Then
              MSHFlexGrid3.TextMatrix(iu, c + 1) = Format(TmpRs.Fields(c), "dd-MMM-yyyy")
           Else
              MSHFlexGrid3.TextMatrix(iu, c + 1) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
           End If
        Else
           MSHFlexGrid3.TextMatrix(iu, c + 1) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
        End If
        MSHFlexGrid3.row = iu
        MSHFlexGrid3.Col = 0
        MSHFlexGrid3.CellFontName = "Wingdings"
        MSHFlexGrid3.CellFontSize = 12
        MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
        MSHFlexGrid3.Text = strUnChecked
    Next

    If LCase(MSHFlexGrid3.TextMatrix(iu, 9)) = "a" Then
       MSHFlexGrid3.TextMatrix(iu, 9) = "Approved"
    ElseIf LCase(MSHFlexGrid3.TextMatrix(iu, 9)) = "p" Then
       MSHFlexGrid3.TextMatrix(iu, 9) = "Pending"
    ElseIf LCase(MSHFlexGrid3.TextMatrix(iu, 8)) = "r" Then
       MSHFlexGrid3.TextMatrix(iu, 9) = "Rejected"
    ElseIf LCase(MSHFlexGrid3.TextMatrix(iu, 8)) = "s" Then
       MSHFlexGrid3.TextMatrix(iu, 9) = "Settled"
    End If
    If MSHFlexGrid3.TextMatrix(iu, 15) = "" Then
       'Mark row as red
       MSHFlexGrid3.row = iu
       For o = 0 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.Col = o
           MSHFlexGrid3.CellBackColor = vbRed
       Next
    End If
    TmpRs.MoveNext
Next
End If
MsgBox "Done"
End Sub

Private Sub CmdPrint_Click()

Dim objPrinter As Printer
Dim ans As Variant
Set objPrinter = GetDefaultPrinter()

If CheckDPSelect = False Then
   MsgBox "Please select Atleast one DP"
   Exit Sub
End If

mPrintOut = False

If Gen_DefaultPrinter = False Then '
   ans = MsgBox("No default printer found, printout will not be generate !!" & vbCrLf & "Do you want to generate only excel file ?", vbYesNo)
   If ans = vbNo Then
      Exit Sub
   End If
Else
   ans = MsgBox("Printout will go to Printer : " & objPrinter.DeviceName & vbCrLf & "Do you want to continue ?", vbYesNo)
   If ans = vbYes Then
      mPrintOut = True
   End If
End If

MsgBox ("Wait till next message")
For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) = strChecked Then
       Gen_mTimeStamp = GetTimeStamp
       Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "GodownRentCoverLetter_" & MSHFlexGrid3.TextMatrix(I, 1) & "_" & CmbBatch.Text & ".xls")
       Call Xls_Detail_Rpt_New(MSHFlexGrid3.TextMatrix(I, 1), MSHFlexGrid3.TextMatrix(I, 5), MSHFlexGrid3.TextMatrix(I, 6), MSHFlexGrid3.TextMatrix(I, 3), MSHFlexGrid3.TextMatrix(I, 15), MSHFlexGrid3.TextMatrix(I, 16), MSHFlexGrid3.TextMatrix(I, 17), MSHFlexGrid3.TextMatrix(I, 18), MSHFlexGrid3.TextMatrix(I, 19))
    End If
Next
MsgBox ("Excel files created in 'Excel' folder")
End Sub


Public Sub Xls_Detail_Rpt_New(RentID_ As Variant, mLessorID_ As Variant, mLessorname_ As Variant, mExpFrom_ As Variant, mBankName_ As Variant, mChequeNo_ As Variant, mPaidAmount_ As Variant, mTDS_ As Variant, mPaidDate_ As Variant)
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow, mRow1 As Integer
Dim mCol As Integer
Dim c As Variant
Set oXL = New Excel.Application
Dim Pat As String
Dim mAmt, mSTAmt, mTotAmt As Double
Pat = App.Path

Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "GodownRentCoverLetter_" & RentID_ & "_" & CmbBatch.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows

oWS.Shapes.AddPicture(App.Path & "\logo.gif", msoTrue, msoTrue, 10, 10, 40, 40).Select

oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 18.86

tmp = "B" & 1 & ":f" & 1
oWS.Range(tmp).Merge
oWS.Cells(1, 2) = "National Collateral Management Services Limited"
oWS.Cells(1, 2).HorizontalAlignment = -4108
oWS.Range(tmp).Font.Bold = True

tmp = "B" & 2 & ":f" & 2
oWS.Range(tmp).Merge
oWS.Cells(2, 2) = "Lodha Supremus, Unit No 505 to 509, 5th Floor,"
oWS.Cells(2, 2).HorizontalAlignment = -4108

tmp = "B" & 3 & ":f" & 3
oWS.Range(tmp).Merge
oWS.Cells(3, 2) = "Kanjurmarg (East), Mumbai-400042"
oWS.Cells(3, 2).HorizontalAlignment = -4108

tmp = "B" & 4 & ":f" & 4
oWS.Range(tmp).Merge
oWS.Cells(4, 2) = "Tel.No-022 40419191, Fax. No-022 40419193"
oWS.Cells(4, 2).HorizontalAlignment = -4108

tmp = "a" & 7 & ":f" & 7
oWS.Range(tmp).Merge
oWS.Cells(7, 1) = ""
oWS.Cells(8, 1) = "To :"

tmp = "a" & 9 & ":f" & 9
oWS.Range(tmp).Merge
oWS.Cells(9, 1) = "      " & mLessorname_
oWS.Cells(9, 1).Font.Bold = True

tmp = "a" & 10 & ":f" & 10
oWS.Range(tmp).Merge
oWS.Cells(10, 1) = "" & GetLessorAddress(mLessorID_)  '""  'Address
oWS.Range(tmp).WrapText = True

'oWS.Cells(10, 1).RowHeight = 45
oWS.Cells(11, 1) = "Dear Sir,"
tmp = "a" & 13 & ":f" & 13
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
'oWS.Cells(13, 1) = "Please find enclosed Rent paid details for the month " & _
'                   " of " & MonthName(Month(mExpFrom_)) & " " & Year(mExpFrom_) & ". The details are given below."
'oWS.Cells(16, 1) = "Agreement ID"
'oWS.Cells(16, 2) = "Location"
'oWS.Cells(16, 3) = "Godown No."
'oWS.Cells(16, 4) = "Area in SqFt"
'oWS.Cells(16, 5) = "Agree. Expiry Date"
'oWS.Cells(16, 6) = "Gross Amount"
'tmp = "a" & 16 & ":f" & 16
'oWS.Range(tmp).Font.Bold = True

''-- Changed 19 March 2016
oWS.Cells(13, 1) = "Please find the Rent details as per below mentioned table."

tmp = "a" & 14 & ":d" & 14
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(14, 1) = "Cheque No.(" & mChequeNo_ & ")"

oWS.Cells(16, 1) = "Location"
oWS.Cells(16, 2) = "Godown No"
oWS.Cells(16, 3) = "Rent for the Month of"
oWS.Cells(16, 4) = "Claim ID"
oWS.Cells(16, 5) = "Area in SqFt"
oWS.Cells(16, 6) = "Gross Amount"

tmp = "a" & 16 & ":f" & 16
oWS.Range(tmp).Font.Bold = True
''-----------------

tmp = "Select TLA.LAID,ML.LocationName,MG.GodownNo,MG.Area,TLA.ExpiryDate,TGRD.Amount,TGR.ExpenseFrom,TGR.ClaimID,TGR.ChequeNo"
tmp = tmp & " from Txn_GodownRent TGR"
tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TGRD.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID = MG.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGRD.LAID=TLA.LAID"
If mChequeNoFlag = False Then
    tmp = tmp & " Where TGR.ClaimID=" & Val(RentID_)
Else
    tmp = tmp & " Where TGR.ChequeNo= '" & Trim(TxtChequeNo) & "'"
End If
Call TmpTable

tmp = "Select PANNo,RTGSNo From Mst_Lessor Where LessorID=" & mLessorID_
Call Tmp1Table
mRow = 17

mAmt = 0
mSTAmt = 0
mTotAmt = 0
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       c = 1
'       For Col = 0 To TmpRs.Fields.Count - 1
'           If InStr(TmpRs.Fields(Col), "/") > 0 Then
'              If IsDate(TmpRs.Fields(Col)) Then
'                 oWS.Cells(mRow, c) = Format(TmpRs.Fields(Col), "dd-MMM-yyyy")
'              Else
'                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
'              End If
'           Else
'              If Col = 3 Then
'                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
'                 mAmt = mAmt + Val(TmpRs.Fields(Col))
'              ElseIf Col = 4 Then
'                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
'              ElseIf Col = 5 Then
'                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
'                 mTotAmt = mTotAmt + Val(TmpRs.Fields(Col))
'              Else
'                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
'              End If
'           End If

             oWS.Cells(mRow, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
             oWS.Cells(mRow, 2) = IIf(IsNull(TmpRs!godownno), "", TmpRs!godownno)
             oWS.Cells(mRow, 3) = IIf(IsNull(TmpRs!ExpenseFrom), "", Format(TmpRs!ExpenseFrom, "MMMM"))
             oWS.Cells(mRow, 4) = IIf(IsNull(TmpRs!ClaimID), "", TmpRs!ClaimID)
             oWS.Cells(mRow, 5) = IIf(IsNull(TmpRs!Area), "", TmpRs!Area)
             oWS.Cells(mRow, 6) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
             mTotAmt = mTotAmt + Val(IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount))
'           c = c + 1
'       Next
       TmpRs.MoveNext
       

       mRow = mRow + 1
   Next
End If
tmp = "a" & 16 & ":f" & mRow
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "a" & mRow & ":e" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Font.Bold = True
oWS.Cells(mRow, 1) = "Total Gross Amount"
'oWS.Cells(mRow, 4) = Format(mAmt, "#####0.00")
oWS.Cells(mRow, 6) = Format(mTotAmt, "#####0.00")


tmp = "d" & mRow & ":f" & mRow
oWS.Range(tmp).Font.Bold = True

tmp = "b" & mRow + 2 & ":b" & mRow + 10
oWS.Range(tmp).Font.Bold = True
tmp = "b" & mRow + 2 & ":e" & mRow + 10
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "b" & mRow + 2 & ":c" & mRow + 2
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 2, 2) = "Bank Name"
tmp = "d" & mRow + 2 & ":e" & mRow + 2
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 2, 4) = mBankName_

tmp = "b" & mRow + 3 & ":c" & mRow + 3
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 3, 2) = "Total of Gross Amount"
tmp = "d" & mRow + 3 & ":e" & mRow + 3
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 3, 4) = Format(Val(mTotAmt), "#####0.00")

tmp = "b" & mRow + 4 & ":c" & mRow + 4
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 4, 2) = "Paid Amount"
tmp = "d" & mRow + 4 & ":e" & mRow + 4
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 4, 4) = Format(Val(mPaidAmount_), "#####0.00")

tmp = "b" & mRow + 5 & ":c" & mRow + 5
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 5, 2) = "TDS"
tmp = "d" & mRow + 5 & ":e" & mRow + 5
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 5, 4) = Format(Val(mTDS_), "#####0.00")

tmp = "b" & mRow + 6 & ":c" & mRow + 6
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 6, 2) = "Net Amount"
tmp = "d" & mRow + 6 & ":e" & mRow + 6
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 6, 4) = Format(Val(mPaidAmount_) + Val(mTDS_), "#####0.00")

tmp = "b" & mRow + 7 & ":c" & mRow + 7
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 7, 2) = "Cheque No."
tmp = "d" & mRow + 7 & ":e" & mRow + 7
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 7, 4) = "'" & mChequeNo_

tmp = "b" & mRow + 8 & ":c" & mRow + 8
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 8, 2) = "Paid Date"
tmp = "d" & mRow + 8 & ":e" & mRow + 8
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 8, 4) = "'" & mPaidDate_

tmp = "b" & mRow + 9 & ":c" & mRow + 9
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 9, 2) = "RTGS Details"
tmp = "d" & mRow + 9 & ":e" & mRow + 9
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
If LCase(TxtChequeNo) = "trf" Then
   oWS.Cells(mRow + 9, 4) = "'" & IIf(IsNull(TmpRs1!RTGSNo), "", TmpRs1!RTGSNo)
Else
   oWS.Cells(mRow + 9, 4) = ""
End If

tmp = "b" & mRow + 10 & ":c" & mRow + 10
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 10, 2) = "PAN Card No."
tmp = "d" & mRow + 10 & ":e" & mRow + 10
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 10, 4) = "'" & IIf(IsNull(TmpRs1!PANNo), "", TmpRs1!PANNo)

tmp = "a" & mRow + 13 & ":f" & mRow + 13
oWS.Range(tmp).Merge
oWS.Cells(mRow + 13, 1) = "For National Collateral Management Services Ltd."

tmp = "a" & mRow + 16 & ":f" & mRow + 16
oWS.Range(tmp).Merge
oWS.Cells(mRow + 16, 1) = "(Authorised Signatory)"

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If

oWS.Columns.AutoFit
oWS.PageSetup.Orientation = xlPortrait
oWB.Save

If mPrintOut = True Then
   oWS.PrintOut
Else
   oXL.Visible = True
End If
'oXL.Quit
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame5.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame5.Enabled = True
'Call Grid_Head
Call CmdClear_Click
End Sub
Private Function CheckDPSelect() As Boolean
Dim Retval As Boolean
Retval = False
For I = 0 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckDPSelect = Retval
End Function

Private Sub MSHFlexGrid3_Click()
With MSHFlexGrid3
    If .TextMatrix(.row, 1) = "" Then Exit Sub
    If .Col = 0 Then
        If (.CellBackColor) <> vbRed Then
           If .TextMatrix(.row, 0) = strUnChecked Then
              .TextMatrix(.row, 0) = strChecked
           Else
              .TextMatrix(.row, 0) = strUnChecked
           End If
        End If
    End If
End With
End Sub
Private Sub CmbSLessor_GotFocus()
tmp = CmbSLessor.Text
mSLessorID = ""
Call TableMst_Lessor("")
CmbSLessor.Clear
If RsMst_Lessor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Lessor.RecordCount
    CmbSLessor.AddItem RsMst_Lessor!LessorName
    RsMst_Lessor.MoveNext
Next
CmbSLessor.Text = tmp
End Sub

Private Sub CmbSLessor_LostFocus()
If CmbSLessor.Text = "" Then
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorName = '" & CmbSLessor.Text & "'"
If RsMst_Lessor.EOF Then
   MsgBox "Invalid selection "
   CmbSLessor.SetFocus
   Exit Sub
End If
mSLessorID = RsMst_Lessor!LessorID
End Sub

Private Sub SFrom_LostFocus()
mSMonth = Month(SFrom.Value)
mSYear = Year(SFrom.Value)
mSTotalDays = Day(DateSerial(mYear, mSMonth + 1, 0))
SFrom.Value = CDate(mSYear & "-" & mSMonth & "-" & "01")

End Sub

