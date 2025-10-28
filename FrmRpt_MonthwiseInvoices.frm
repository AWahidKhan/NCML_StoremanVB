VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_MonthwiseInvoices 
   Caption         =   "Monthwise Invoice"
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
      Height          =   8655
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8175
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14420
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
   End
   Begin VB.Frame Frame11 
      Height          =   1200
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.OptionButton OptDPIncome 
         Caption         =   "DP Income"
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
         Left            =   13200
         TabIndex        =   19
         Top             =   165
         Width           =   1335
      End
      Begin VB.OptionButton OptWithClient 
         Caption         =   "With Client Name"
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
         Left            =   10695
         TabIndex        =   18
         Top             =   165
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "With Client Name"
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
         Left            =   8040
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox ChkPendingInvoice 
         Caption         =   "With Pending Invoices"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         TabIndex        =   16
         Top             =   405
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Summary"
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
         Height          =   255
         Left            =   9000
         TabIndex        =   15
         Top             =   720
         Width           =   1200
      End
      Begin VB.CheckBox Check1 
         Caption         =   "AM/RM"
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
         Height          =   255
         Left            =   9000
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox CmbStateID 
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
         Left            =   3330
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   510
         Width           =   2685
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
         Left            =   6120
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   510
         Width           =   2760
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
         Height          =   375
         Left            =   13680
         TabIndex        =   3
         Top             =   705
         Width           =   1215
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
         Height          =   375
         Left            =   11250
         TabIndex        =   2
         Top             =   705
         Width           =   1215
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
         Height          =   375
         Left            =   12465
         TabIndex        =   1
         Top             =   705
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   510
         Width           =   1575
         _ExtentX        =   2778
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   111411203
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   510
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   111411203
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label LblStateID 
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
         Height          =   300
         Left            =   4185
         TabIndex        =   13
         Top             =   195
         Width           =   1200
      End
      Begin VB.Label LblLocationID 
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
         Left            =   7080
         TabIndex        =   12
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "From Month Year"
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
         TabIndex        =   7
         Top             =   195
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "To Month Year"
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
         Left            =   1800
         TabIndex        =   6
         Top             =   195
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmRpt_MonthwiseInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID As Variant
Dim mTotalDays, mTotalMonth As Variant
Dim mMonthLastDate As Date
Dim mDate As Date
Dim mStart As Integer
Dim mFlag As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Check2.Visible = True
   'Check3.Visible = True
   OptWithClient.Visible = True
   OptDPIncome.Visible = True
Else
   Check2.Value = 0
   Check2.Visible = False
   OptWithClient.Value = False
   'Check3.Visible = False
   OptWithClient.Visible = False
   OptDPIncome.Visible = False
End If
End Sub






Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState = "" Then
   Call TableMst_State
Else
   Call TableMst_State(" Where " & Gen_WcState)
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If mStateID <> RsMst_State!StateID Then
   CmbLocationID.Clear
   mLocationID = Null
End If

mStateID = RsMst_State!StateID

End Sub
Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmbStateID.Text = ""
CmbLocationID.Text = ""
Check2.Value = 0
Check1.Value = 0
Check2.Visible = False
OptWithClient.Visible = False
OptDPIncome.Visible = False
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_MonthWiseInvoices.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_MonthWiseInvoices.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
If Check1.Value = 0 Then
    oWS.Range("a:a").ColumnWidth = 10.86
    oWS.Range("f:f").ColumnWidth = 15
    oWS.Range("e:e").ColumnWidth = 16.14
    oWS.Range("d:d").ColumnWidth = 14
    oWS.Range("c:c").ColumnWidth = 17.14
    oWS.Range("b:b").ColumnWidth = 14
    oWS.Range("g:g").ColumnWidth = 14
    oWS.Range("h:h").ColumnWidth = 15
    oWS.Range("i:i").ColumnWidth = 14
    oWS.Range("j:j").ColumnWidth = 14
    oWS.Range("k:k").ColumnWidth = 15
    oWS.Range("l:l").ColumnWidth = 15
    oWS.Range("m:m").ColumnWidth = 15
    tmp = "a" & mRow & ":" & "m" & mRow
Else
    oWS.Range("a:a").ColumnWidth = 27
    oWS.Range("b:b").ColumnWidth = 27
    oWS.Range("c:c").ColumnWidth = 15
    oWS.Range("d:d").ColumnWidth = 15.3
    If Check2.Value = 1 And (OptWithClient.Value = False And OptDPIncome.Value = False) Then
       oWS.Range("e:e").ColumnWidth = 0
    ElseIf Check2.Value = 1 And (OptWithClient.Value = True Or OptDPIncome.Value = True) Then
       oWS.Range("e:e").ColumnWidth = 73.29
       
    Else
       oWS.Range("e:e").ColumnWidth = 73.29
    End If
    If Check2.Value = 1 And (OptWithClient.Value = True Or OptDPIncome.Value = True) Then
       oWS.Range("f:f").ColumnWidth = 0
    Else
       oWS.Range("f:f").ColumnWidth = 15
    End If
    oWS.Range("g:g").ColumnWidth = 16.14
    oWS.Range("h:h").ColumnWidth = 14
    oWS.Range("i:i").ColumnWidth = 17.14
    oWS.Range("j:j").ColumnWidth = 14
    oWS.Range("k:k").ColumnWidth = 14
    oWS.Range("l:l").ColumnWidth = 15
    oWS.Range("m:m").ColumnWidth = 14
    oWS.Range("n:n").ColumnWidth = 14
    oWS.Range("o:o").ColumnWidth = 15
    oWS.Range("p:p").ColumnWidth = 15
    oWS.Range("q:q").ColumnWidth = 15
    tmp = "a" & mRow & ":" & "q" & mRow
End If

oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Income Category wise Monthly Income MIS"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
'&H8080BF

For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    
    For C = 0 To MSHFlexGrid1.Cols - 1
        If Check2.Value = 1 And C = 4 And (OptWithClient.Value = False And OptDPIncome.Value = False) Then
        ElseIf Check2.Value = 1 And C = 5 And (OptWithClient.Value = True And OptDPIncome.Value = True) Then
        Else
            If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
               If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
                  oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
               Else
                  oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
               End If
            Else
               oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
            End If
            
            If I = 0 Then
               oWS.Cells(mRow, C + 1).Interior.Color = &HD9D9FF
               oWS.Cells(mRow, C + 1).Font.Bold = True
            ElseIf I = 1 Then
               oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
               oWS.Cells(mRow, C + 1).Font.Bold = True
               oWS.Cells(mRow, C + 1).RowHeight = 45
               oWS.Cells(mRow, C + 1).WrapText = True
            End If
        End If
    Next
Next
If Check1.Value = 1 Then
   tmp = "a" & 1 & ":" & "q" & mRow
Else
   tmp = "a" & 1 & ":" & "m" & mRow
End If
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
Dim Wc1 As String
Dim tmp1 As String
wc = ""
mFlag = False
Call Grid_Head

TxtFromDate.Value = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")
mTotalDays = Day(DateSerial(Year(TxtToDate.Value), Month(TxtToDate.Value) + 1, 0))
TxtToDate.Value = (Year(TxtToDate.Value) & "-" & Month(TxtToDate.Value) & "-" & mTotalDays)

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox " 'From Month Year' should be less than 'To Month Year'"
   Exit Sub
End If
mTotalMonth = DateDiff("m", TxtFromDate.Value, TxtToDate.Value)
mDate = TxtFromDate.Value

If Check1.Value = 1 Then
   Call AddLocationDetails
Else
   For I = 0 To mTotalMonth
       r = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(r, 0) = Format(mDate, "dd-MMM-yyyy")
       mDate = DateAdd("m", 1, mDate)
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Next
End If
If mFlag = False And (OptWithClient.Value = True Or OptDPIncome.Value = True) Then
   MsgBox "no Data Exists"
   Exit Sub
End If
mDate = TxtFromDate.Value

'--------Adhoc Inv
tmp = "Select TI.IETypeID,MIE.IEType,sum(TI.Amount)As IncomeAmt"
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,TI.InvoiceDate"
End If
If OptWithClient.Value = True Then
   tmp = tmp & " ,TI.ClientIDRow, MC.ClientName"
End If

tmp = tmp & " From Txn_TempInvoice TI"
tmp = tmp & " Inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.IETypeID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID=TI.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
If OptWithClient.Value = True Then
   tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = TI.ClientIDRow"
End If
If ChkPendingInvoice.Value = 0 Then
   tmp = tmp & " Where TI.InvoiceDate>='" & TxtFromDate.Value & "' And TI.InvoiceDate<='" & TxtToDate.Value & "' And MIE.Flag='A' And TI.Flag = 'A' And isnull(TI.ReverseFlag,'')<>'Y'"
Else
   tmp = tmp & " Where TI.InvoiceDate>='" & TxtFromDate.Value & "' And TI.InvoiceDate<='" & TxtToDate.Value & "' And MIE.Flag='A' And TI.Flag in ('A','P') And isnull(TI.ReverseFlag,'')<>'Y'"
End If
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID
Else
  If CmbStateID.Text <> "" Then
     tmp = tmp & " And ML.StateID=" & mStateID
  End If
End If

tmp = tmp & " Group By TI.IETypeID,MIE.IEType"
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,TI.InvoiceDate"
End If
If OptWithClient.Value = True Then
   tmp = tmp & " ,TI.ClientIDRow, MC.ClientName"
End If
Call TmpTable

'---------DP Inv
'COmmented on 22 July 2013 after mail from Jaimin and approved by Mayur Sir
'tmp = "Select sum(TotalAmt)As IncomeAmt "
'Amt

tmp = "Select sum(Amt)As IncomeAmt "
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,MDP.MonthYear"
End If
If OptDPIncome.Value = True Then
   tmp = tmp & " , MD.DepositoryName"
End If
tmp = tmp & " from Mst_DPInvoiceUpload MDP"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID=MDP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
If OptDPIncome.Value = True Then
   tmp = tmp & " Inner Join Mst_Depository MD on MD.NavCode = MDP.DPNavCode and MD.ExchangeTypeID = 2"
End If
tmp = tmp & " Where MDP.MonthYear>='" & TxtFromDate.Value & "' And MDP.MonthYear <='" & TxtToDate.Value & "' And isnull(MDP.ReverseFlag,'')<>'Y'"
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID
Else
  If CmbStateID.Text <> "" Then
     tmp = tmp & " And ML.StateID=" & mStateID
  End If
End If
Wc1 = " Group By"
If Check1.Value = 1 Then
   Wc1 = Wc1 & " ML.LocationName"
End If
If OptDPIncome.Value = True Then
   Wc1 = Wc1 & " , MD.DepositoryName"
End If
If Check2.Value = 0 Then
   If Wc1 = " Group By" Then
      Wc1 = Wc1 & " MDP.MonthYear"
   Else
      Wc1 = Wc1 & " ,MDP.MonthYear"
   End If
End If

tmp = tmp & Wc1

Call Tmp2Table

'-----------NNNR Inv
tmp = "Select TIG.PINo,Sum(TIG.InvoiceAmount)InvoiceAmount,max(Discount)Discount,max(Premium)Premium"
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,TIG.YearMonth As InvoiceDate "
End If
If OptWithClient.Value = True Then
   tmp = tmp & " ,TIG.ClientIDRow, MC.ClientName "
End If
tmp = tmp & " from  Txn_InvMonGSLDetail TIG"
tmp = tmp & " Left join Txn_ProInvoice TP on TP.InvoiceNo=TIG.PINo And (TP.Flag = 'C' Or TP.Flag = 'P') " 'TP.Flag = 'C'"
tmp = tmp & " inner Join Mst_Location ML on ML.SM_Prefix=TIG.SM_Prefix"
tmp = tmp & " inner Join Mst_State MS on ML.StateID=MS.StateID"
If OptWithClient.Value = True Then
   tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = TIG.ClientIDRow"
End If
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null) And PINo Is Not Null"
tmp = tmp & " And TIG.YearMonth >='" & TxtFromDate.Value & "' And TIG.YearMonth <='" & TxtToDate.Value & "'"
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID
Else
  If CmbStateID.Text <> "" Then
     tmp = tmp & " And ML.StateID=" & mStateID
  End If
End If
tmp = tmp & " Group By TIG.PINo"
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,TIG.YearMonth"
End If
If OptWithClient.Value = True Then
   tmp = tmp & " ,TIG.ClientIDRow, MC.ClientName "
End If
Call Tmp3Table

'----------Other Inv
'tmp = "Select TI.IETypeID,MIE.IEType,sum(TI.Amount)As IncomeAmt "
'If Check1.Value = 1 Then
'   tmp = tmp & " ,ML.LocationName"
'End If
'If Check2.Value = 0 Then
'   tmp = tmp & " ,TI.TxnDate"
'End If
'tmp = tmp & " from Txn_OtherIncome TI"
'tmp = tmp & " inner Join Mst_IncomeExpenseType MIE on MIE.IETypeID=TI.IETypeID"
'tmp = tmp & " inner Join Mst_Location ML on ML.LocationID=TI.LocationID"
'tmp = tmp & " inner Join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " Where TI.TxnDate>='" & TxtFromDate.Value & "' And TI.TxnDate<='" & TxtToDate.Value & "' And MIE.Flag='A'"
'If CmbLocationID.Text <> "" Then
'   tmp = tmp & " And ML.LocationID=" & mLocationID
'Else
'  If CmbStateID.Text <> "" Then
'     tmp = tmp & " And ML.StateID=" & mStateID
'  End If
'End If
'tmp = tmp & " Group By TI.IETypeID,MIE.IEType"
'If Check1.Value = 1 Then
'   tmp = tmp & " ,ML.LocationName"
'End If
'If Check2.Value = 0 Then
'   tmp = tmp & " ,TI.TxnDate"
'End If
'
'
'Call Tmp4Table


'---------Reservation Inv
tmp = "Select Sum(ActualAmount)As IncomeAmt"
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,TI.YearMonth"
End If
If OptWithClient.Value = True Then
   tmp = tmp & " ,MSC.ClientIDRow, MC.ClientName "
End If
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID=TI.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
If OptWithClient.Value = True Then
   tmp = tmp & " Inner Join Mst_StorageContract MSC on MSC.ContractID = TI.ContractID"
   tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = MSC.ClientIDRow"
End If
If ChkPendingInvoice.Value = 0 Then
   tmp = tmp & " Where TI.YearMonth>='" & TxtFromDate.Value & "' And TI.YearMonth<='" & TxtToDate.Value & "' And TI.Flag='A' And isnull(TI.ReverseFlag,'')<>'Y'"
Else
   tmp = tmp & " Where TI.YearMonth>='" & TxtFromDate.Value & "' And TI.YearMonth<='" & TxtToDate.Value & "' And TI.Flag in ('A','P') And isnull(TI.ReverseFlag,'')<>'Y'"
End If

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID
Else
  If CmbStateID.Text <> "" Then
     tmp = tmp & " And ML.StateID=" & mStateID
  End If
End If
Wc1 = " Group By"
If Check1.Value = 1 Then
   Wc1 = Wc1 & " ML.LocationName"
End If
If OptWithClient.Value = True Then
   Wc1 = Wc1 & " ,MSC.ClientIDRow, MC.ClientName "
End If
If Check2.Value = 0 Then
   If Wc1 = " Group By" Then
      Wc1 = Wc1 & " TI.YearMonth"
   Else
      Wc1 = Wc1 & " ,TI.YearMonth"
   End If
End If

tmp = tmp & Wc1

Call Tmp1Table

'---------7 Days Inv
tmp = "Select sum(Amount)As IncomeAmt, Sum(Discount)Discount,sum(Premium)Premium"
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,TC.BillingMonth"
End If
If OptWithClient.Value = True Then
   tmp = tmp & " ,TC.ClientIDRow, MC.ClientName"
End If
tmp = tmp & " From Txn_CDTFBOE TC"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID=TC.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
If OptWithClient.Value = True Then
   tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow=TC.ClientIDRow"
End If
tmp = tmp & " Where TC.TxnType='D' And TC.BillingMonth>='" & TxtFromDate.Value & "' And TC.BillingMonth <='" & TxtToDate.Value & "'"
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID
Else
  If CmbStateID.Text <> "" Then
     tmp = tmp & " And ML.StateID=" & mStateID
  End If
End If
Wc1 = " Group By"
If Check1.Value = 1 Then
   Wc1 = Wc1 & " ML.LocationName"
End If
If OptWithClient.Value = True Then
   Wc1 = Wc1 & " ,TC.ClientIDRow, MC.ClientName"
End If
If Check2.Value = 0 Then
   If Wc1 = " Group By" Then
      Wc1 = Wc1 & " TC.BillingMonth"
   Else
      Wc1 = Wc1 & " ,TC.BillingMonth"
   End If
End If
tmp = tmp & Wc1
Call Tmp5Table

tmp = "Select TC.TxnType,sum(Amount)As IncomeAmt, Sum(Discount)Discount,sum(Premium)Premium"
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,TC.BillingMonth"
End If
If OptWithClient.Value = True Then
   tmp = tmp & "  ,TC.ClientIDRow, MC.ClientName"
End If

tmp = tmp & " From Txn_CDTFBOE TC"
tmp = tmp & " inner Join Mst_Location ML on ML.LocationID=TC.LocationID"
tmp = tmp & " inner Join Mst_State MS on ML.StateID=MS.StateID"
If OptWithClient.Value = True Then
   tmp = tmp & "  Inner Join Mst_Client MC on MC.ClientIDRow=TC.ClientIDRow"
End If
tmp = tmp & " Where TC.BillingMonth>='" & TxtFromDate.Value & "' And TC.BillingMonth <='" & TxtToDate.Value & "'"
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID
Else
  If CmbStateID.Text <> "" Then
     tmp = tmp & " And ML.StateID=" & mStateID
  End If
End If

Wc1 = " Group By TC.TxnType"
If Check1.Value = 1 Then
   Wc1 = Wc1 & ", ML.LocationName"
End If
If OptWithClient.Value = True Then
   Wc1 = Wc1 & " ,TC.ClientIDRow, MC.ClientName"
End If
If Check2.Value = 0 Then
   Wc1 = Wc1 & " ,TC.BillingMonth"
End If
tmp = tmp & Wc1

Call Tmp6Table


'Commtrack
tmp = "Select sum(TotalAmt)As IncomeAmt "
If Check1.Value = 1 Then
   tmp = tmp & " ,ML.LocationName"
End If
If Check2.Value = 0 Then
   tmp = tmp & " ,MDP.MonthYear"
End If
If OptDPIncome.Value = True Then
   tmp = tmp & " , MD.DepositoryName"
End If
tmp = tmp & " from Txn_MonthlyCommtrackInvoice MDP"
tmp = tmp & " inner Join Mst_Location ML on ML.LocationID=MDP.LocationID"
tmp = tmp & " inner Join Mst_State MS on ML.StateID=MS.StateID"
If OptDPIncome.Value = True Then
   tmp = tmp & " Inner Join Mst_Depository MD on MD.NavCode = MDP.CPNavCode and MD.ExchangeTypeID = 2"
End If
tmp = tmp & " Where MDP.MonthYear>='" & TxtFromDate.Value & "' And MDP.MonthYear <='" & TxtToDate.Value & "' And isnull(MDP.ReverseFlag,'')<>'Y'"
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID
Else
  If CmbStateID.Text <> "" Then
     tmp = tmp & " And ML.StateID=" & mStateID
  End If
End If
Wc1 = " Group By"
If Check1.Value = 1 Then
   Wc1 = Wc1 & " ML.LocationName"
End If
If OptDPIncome.Value = True Then
   Wc1 = Wc1 & " , MD.DepositoryName"
End If
If Check2.Value = 0 Then
   If Wc1 = " Group By" Then
      Wc1 = Wc1 & " MDP.MonthYear"
   Else
      Wc1 = Wc1 & " ,MDP.MonthYear"
   End If
End If
tmp = tmp & Wc1

Call Tmp7Table

   For r = 2 To MSHFlexGrid1.Rows - 2
       If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
          mStart = 4
          TmpRs.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
          mStart = 4
          TmpRs.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
       ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
          mStart = 5
          'TmpRs.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
          TmpRs.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "' And InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
          mStart = 5
          'TmpRs.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
          TmpRs.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
       ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
          mStart = 5
          TmpRs.Filter = "LocationName='NA'"
       Else
          mStart = 0
          TmpRs.Filter = "InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 0) & "'"
       End If
       For rn = 1 To TmpRs.RecordCount
           If TmpRs!IETypeID = 1 Then ' 7 Days
              MSHFlexGrid1.TextMatrix(r, mStart + 1) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 1)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 2 Then ' Rejected
              MSHFlexGrid1.TextMatrix(r, mStart + 4) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 4)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 3 Then ' Spot
              MSHFlexGrid1.TextMatrix(r, mStart + 6) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 6)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 4 Then ' Reservation
              MSHFlexGrid1.TextMatrix(r, mStart + 5) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 5)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 5 Then ' Other
              MSHFlexGrid1.TextMatrix(r, mStart + 3) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 3)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 6 Then ' Spillage
              MSHFlexGrid1.TextMatrix(r, mStart + 8) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 8)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 7 Then ' NNNR
              MSHFlexGrid1.TextMatrix(r, mStart + 2) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 2)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 8 Then ' DP
              MSHFlexGrid1.TextMatrix(r, mStart + 7) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 7)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 9 Then ' Cheque return
              MSHFlexGrid1.TextMatrix(r, mStart + 9) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 9)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 11 Then ' SCM
              MSHFlexGrid1.TextMatrix(r, mStart + 10) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 10)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 12 Then ' Commtrack
              MSHFlexGrid1.TextMatrix(r, mStart + 11) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 11)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 13 Then ' Commtrack
              MSHFlexGrid1.TextMatrix(r, mStart + 14) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 14)) + TmpRs!IncomeAmt
           ElseIf TmpRs!IETypeID = 15 Then ' Space Charges
              MSHFlexGrid1.TextMatrix(r, mStart + 15) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 15)) + TmpRs!IncomeAmt
           
           End If
            TmpRs.MoveNext
        Next
        TmpRs.Filter = ""
    '---------For DP Income
 'If OptDPIncome.Value = True Then   'Mst_DPInvoiceUpload is NOT included in CLientwise Income
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       mTotalDays = Day(DateSerial(Year(MSHFlexGrid1.TextMatrix(r, 4)), Month(MSHFlexGrid1.TextMatrix(r, 4)) + 1, 0))
       mMonthLastDate = (Year(MSHFlexGrid1.TextMatrix(r, 4)) & "-" & Month(MSHFlexGrid1.TextMatrix(r, 4)) & "-" & mTotalDays)
       TmpRs2.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And MonthYear='" & mMonthLastDate & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs2.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = True Then
       mTotalDays = Day(DateSerial(Year(MSHFlexGrid1.TextMatrix(r, 5)), Month(MSHFlexGrid1.TextMatrix(r, 5)) + 1, 0))
       mMonthLastDate = (Year(MSHFlexGrid1.TextMatrix(r, 5)) & "-" & Month(MSHFlexGrid1.TextMatrix(r, 5)) & "-" & mTotalDays)
       TmpRs2.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And DepositoryName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And MonthYear='" & mMonthLastDate & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = True Then
       'mTotalDays = Day(DateSerial(Year(MSHFlexGrid1.TextMatrix(r, 5)), Month(MSHFlexGrid1.TextMatrix(r, 5)) + 1, 0))
       'mMonthLastDate = (Year(MSHFlexGrid1.TextMatrix(r, 5)) & "-" & Month(MSHFlexGrid1.TextMatrix(r, 5)) & "-" & mTotalDays)
       TmpRs2.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And DepositoryName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 0 Or Check2.Value = 1) And OptWithClient.Value = True And OptDPIncome.Value = False Then
       TmpRs2.Filter = "LocationName='NA'"
    Else
       mTotalDays = Day(DateSerial(Year(MSHFlexGrid1.TextMatrix(r, 0)), Month(CDate(MSHFlexGrid1.TextMatrix(r, 0))) + 1, 0))
       mMonthLastDate = (Year(MSHFlexGrid1.TextMatrix(r, 0)) & "-" & Month(CDate(MSHFlexGrid1.TextMatrix(r, 0))) & "-" & mTotalDays)
       TmpRs2.Filter = "MonthYear='" & mMonthLastDate & "'"
    End If
    For rn = 1 To TmpRs2.RecordCount
        MSHFlexGrid1.TextMatrix(r, mStart + 7) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 7)) + TmpRs2!IncomeAmt
        TmpRs2.MoveNext
    Next
    TmpRs2.Filter = ""
 'End If

    'TmpRs3.Filter = "InvoiceDate='" & mdate & "'"
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs3.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs3.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs3.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName = '" & MSHFlexGrid1.TextMatrix(r, 4) & "' And InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       TmpRs3.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow = '" & MSHFlexGrid1.TextMatrix(r, 18) & "' And InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs3.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName = '" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       TmpRs3.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow = '" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
       TmpRs3.Filter = "LocationName='NA'"
    Else
       TmpRs3.Filter = "InvoiceDate='" & MSHFlexGrid1.TextMatrix(r, 0) & "'"
    End If
    
    For rn = 1 To TmpRs3.RecordCount
        MSHFlexGrid1.TextMatrix(r, mStart + 2) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 2)) + (IIf(IsNull(TmpRs3!InvoiceAmount), 0, TmpRs3!InvoiceAmount) - IIf(IsNull(TmpRs3!Discount), 0, TmpRs3!Discount) + IIf(IsNull(TmpRs3!Premium), 0, TmpRs3!Premium))
        TmpRs3.MoveNext
    Next
    TmpRs3.Filter = ""
    
    'TmpRs4.Filter = "TxnDate='" & mdate & "'"
' If OptWithClient.Value = False Then '''Txn_OtherIncome is not included in Clientwise Income
'    If Check1.Value = 1 And Check2.Value = 0 Then
'       TmpRs4.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And TxnDate='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
'    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
'       TmpRs4.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
'    Else
'       TmpRs4.Filter = "TxnDate='" & MSHFlexGrid1.TextMatrix(r, 0) & "'"
'    End If
'    For rn = 1 To TmpRs4.RecordCount
'        If TmpRs4!IETypeID = 1 Then ' 7 Days
'           MSHFlexGrid1.TextMatrix(r, mStart + 1) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 1)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 2 Then ' Rejected
'           MSHFlexGrid1.TextMatrix(r, mStart + 4) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 4)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 3 Then ' Spot
'           MSHFlexGrid1.TextMatrix(r, mStart + 6) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 6)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 4 Then ' Reservation
'           MSHFlexGrid1.TextMatrix(r, mStart + 5) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 5)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 5 Then ' Other
'           MSHFlexGrid1.TextMatrix(r, mStart + 3) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 3)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 6 Then ' Spillage
'           MSHFlexGrid1.TextMatrix(r, mStart + 8) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 8)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 7 Then ' NNNR
'           MSHFlexGrid1.TextMatrix(r, mStart + 2) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 2)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 8 Then ' DP
'           MSHFlexGrid1.TextMatrix(r, mStart + 7) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 7)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 9 Then ' Cheque return
'           MSHFlexGrid1.TextMatrix(r, mStart + 9) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 9)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 11 Then ' SCM Income
'           MSHFlexGrid1.TextMatrix(r, mStart + 10) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 10)) + TmpRs4!IncomeAmt
'        ElseIf TmpRs4!IETypeID = 12 Then ' Commtrack
'           MSHFlexGrid1.TextMatrix(r, mStart + 11) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 11)) + TmpRs4!IncomeAmt
'        End If
'        TmpRs4.MoveNext
'    Next
'    TmpRs4.Filter = ""
' End If
'    If Check1.Value = 1 And Check2.Value = 0 Then
'       TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And YearMonth='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
'    ElseIf Check1.Value = 1 And Check2.Value = 1 Then
'       TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
'    Else
'       TmpRs1.Filter = "YearMonth='" & MSHFlexGrid1.TextMatrix(r, 0) & "'"
'    End If
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And YearMonth='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And YearMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "' And YearMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
        TmpRs1.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "'And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
       TmpRs1.Filter = "LocationName='NA'"
    Else
       TmpRs1.Filter = "YearMonth='" & MSHFlexGrid1.TextMatrix(r, 0) & "'"
    End If

'    If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
'    TmpRs5.Find "YearMonth='" & mdate & "'"

    For rn = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.TextMatrix(r, mStart + 5) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 5)) + TmpRs1!IncomeAmt
       TmpRs1.MoveNext
    Next
    TmpRs1.Filter = ""
    
''    If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
''    TmpRs5.Find "BillingMonth='" & mdate & "'"
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs5.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs5.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs5.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       TmpRs5.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs5.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       TmpRs5.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
       TmpRs5.Filter = "LocationName='NA'"
    Else
       TmpRs5.Filter = "BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 0) & "'"
    End If
    
    For rn = 1 To TmpRs5.RecordCount
       MSHFlexGrid1.TextMatrix(r, mStart + 1) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 1)) + (TmpRs5!IncomeAmt - TmpRs5!Discount + TmpRs5!Premium)
       TmpRs5.MoveNext
    Next
    TmpRs5.Filter = ""
    
    'Rejected Income
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
       TmpRs6.Filter = "LocationName='NA'"
    Else
       TmpRs6.Filter = "BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 0) & "' And TxnType = 'W'"
    End If
    
    For rn = 1 To TmpRs6.RecordCount
       MSHFlexGrid1.TextMatrix(r, mStart + 4) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 4)) + (TmpRs6!IncomeAmt - TmpRs6!Discount + TmpRs6!Premium)
       TmpRs6.MoveNext
    Next
    TmpRs6.Filter = ""
    
    
    'PWR Income
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
       TmpRs6.Filter = "LocationName='NA'"
    Else
       TmpRs6.Filter = "BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 0) & "' And TxnType = 'P'"
    End If
    
    For rn = 1 To TmpRs6.RecordCount
       MSHFlexGrid1.TextMatrix(r, mStart + 13) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 13)) + (TmpRs6!IncomeAmt - TmpRs6!Discount + TmpRs6!Premium)
       TmpRs6.MoveNext
    Next
    TmpRs6.Filter = ""
    
    'RRN Withdrawal Income
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
       TmpRs6.Filter = "LocationName='NA'"
    Else
       TmpRs6.Filter = "BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 0) & "' And TxnType = 'R'"
    End If
    
    For rn = 1 To TmpRs6.RecordCount
       MSHFlexGrid1.TextMatrix(r, mStart + 14) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 14)) + (TmpRs6!IncomeAmt - TmpRs6!Discount + TmpRs6!Premium)
       TmpRs6.MoveNext
    Next
    TmpRs6.Filter = ""
    
    'NNNR Yearly Income
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "' And BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 5) & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = True And OptDPIncome.Value = False Then
       'TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientName='" & MSHFlexGrid1.TextMatrix(r, 4) & "'"
       TmpRs6.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And ClientIDRow='" & MSHFlexGrid1.TextMatrix(r, 18) & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 1 Or Check2.Value = 0) And OptWithClient.Value = False And OptDPIncome.Value = True Then
       TmpRs6.Filter = "LocationName='NA'"
    Else
       TmpRs6.Filter = "BillingMonth='" & MSHFlexGrid1.TextMatrix(r, 0) & "' And TxnType = 'Y'"
    End If
    
    For rn = 1 To TmpRs6.RecordCount
       MSHFlexGrid1.TextMatrix(r, mStart + 12) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 12)) + (TmpRs6!IncomeAmt - TmpRs6!Discount + TmpRs6!Premium)
       TmpRs6.MoveNext
    Next
    TmpRs6.Filter = ""
    
    
    If Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       mTotalDays = Day(DateSerial(Year(MSHFlexGrid1.TextMatrix(r, 4)), Month(MSHFlexGrid1.TextMatrix(r, 4)) + 1, 0))
       mMonthLastDate = (Year(MSHFlexGrid1.TextMatrix(r, 4)) & "-" & Month(MSHFlexGrid1.TextMatrix(r, 4)) & "-" & mTotalDays)
       TmpRs7.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And MonthYear='" & mMonthLastDate & "'"
    ElseIf Check1.Value = 1 And Check2.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
       TmpRs7.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' "
    ElseIf Check1.Value = 1 And Check2.Value = 0 And OptWithClient.Value = False And OptDPIncome.Value = True Then
       mTotalDays = Day(DateSerial(Year(MSHFlexGrid1.TextMatrix(r, 5)), Month(MSHFlexGrid1.TextMatrix(r, 5)) + 1, 0))
       mMonthLastDate = (Year(MSHFlexGrid1.TextMatrix(r, 5)) & "-" & Month(MSHFlexGrid1.TextMatrix(r, 5)) & "-" & mTotalDays)
       TmpRs7.Filter = "LocationName='" & MSHFlexGrid1.TextMatrix(r, 3) & "' And DepositoryName='" & MSHFlexGrid1.TextMatrix(r, 4) & "' And MonthYear='" & mMonthLastDate & "'"
    ElseIf Check1.Value = 1 And (Check2.Value = 0 Or Check2.Value = 1) And OptWithClient.Value = True And OptDPIncome.Value = False Then
       TmpRs7.Filter = "LocationName='NA'"
    Else
       mTotalDays = Day(DateSerial(Year(MSHFlexGrid1.TextMatrix(r, 0)), Month(CDate(MSHFlexGrid1.TextMatrix(r, 0))) + 1, 0))
       mMonthLastDate = (Year(MSHFlexGrid1.TextMatrix(r, 0)) & "-" & Month(CDate(MSHFlexGrid1.TextMatrix(r, 0))) & "-" & mTotalDays)
       TmpRs7.Filter = "MonthYear='" & mMonthLastDate & "'"
    End If
    For rn = 1 To TmpRs7.RecordCount
            MSHFlexGrid1.TextMatrix(r, mStart + 11) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 11)) + TmpRs7!IncomeAmt
            TmpRs7.MoveNext
    Next
    TmpRs7.Filter = ""
    
    MSHFlexGrid1.TextMatrix(r, mStart + 16) = Val(MSHFlexGrid1.TextMatrix(r, mStart + 1)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 2)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 3)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 4)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 5)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 6)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 7)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 8)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 9)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 10)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 11)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 12)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 13)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 14)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 15))
    
    'MSHFlexGrid1.TextMatrix(0, mStart + 10) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 10)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 10))
    
    MSHFlexGrid1.TextMatrix(0, mStart + 1) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 1)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 1))
    MSHFlexGrid1.TextMatrix(0, mStart + 2) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 2)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 2))
    MSHFlexGrid1.TextMatrix(0, mStart + 3) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 3)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 3))
    MSHFlexGrid1.TextMatrix(0, mStart + 4) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 4)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 4))
    MSHFlexGrid1.TextMatrix(0, mStart + 5) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 5)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 5))
    MSHFlexGrid1.TextMatrix(0, mStart + 6) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 6)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 6))
    MSHFlexGrid1.TextMatrix(0, mStart + 7) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 7)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 7))
    MSHFlexGrid1.TextMatrix(0, mStart + 8) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 8)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 8))
    MSHFlexGrid1.TextMatrix(0, mStart + 9) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 9)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 9))
    MSHFlexGrid1.TextMatrix(0, mStart + 10) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 10)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 10))
    MSHFlexGrid1.TextMatrix(0, mStart + 11) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 11)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 11))
    MSHFlexGrid1.TextMatrix(0, mStart + 12) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 12)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 12))
    MSHFlexGrid1.TextMatrix(0, mStart + 13) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 13)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 13))
    MSHFlexGrid1.TextMatrix(0, mStart + 14) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 14)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 14))
    MSHFlexGrid1.TextMatrix(0, mStart + 15) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 15)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 15))
    MSHFlexGrid1.TextMatrix(0, mStart + 16) = Val(MSHFlexGrid1.TextMatrix(0, mStart + 16)) + Val(MSHFlexGrid1.TextMatrix(r, mStart + 16))
    
    
    For C = mStart + 1 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(r, C) = Format(MSHFlexGrid1.TextMatrix(r, C), "###,###,###,###.00")
    Next
    MSHFlexGrid1.TextMatrix(r, mStart) = Format(MSHFlexGrid1.TextMatrix(r, mStart), "MMM-yyyy")
Next
MSHFlexGrid1.TextMatrix(0, mStart + 1) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 1), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 2) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 2), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 3) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 3), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 4) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 4), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 5) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 5), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 6) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 6), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 7) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 7), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 8) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 8), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 9) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 9), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 10) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 10), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 11) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 11), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 12) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 12), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 13) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 13), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 14) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 14), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 15) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 15), "###,###,###,###.00")
MSHFlexGrid1.TextMatrix(0, mStart + 16) = Format(MSHFlexGrid1.TextMatrix(0, mStart + 16), "###,###,###,###.00")

'Else
    
'End If
CmdExcel.Enabled = True

MsgBox "Done!!"

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3

MSHFlexGrid1.FixedRows = 2
If Check1.Value = 1 And OptWithClient.Value = False And OptDPIncome.Value = False Then
   MSHFlexGrid1.Cols = 19
   MSHFlexGrid1.TextMatrix(1, 0) = "RM"
   MSHFlexGrid1.TextMatrix(1, 1) = "AM"
   MSHFlexGrid1.TextMatrix(1, 2) = "State"
   MSHFlexGrid1.TextMatrix(1, 3) = "Location"
   MSHFlexGrid1.TextMatrix(1, 4) = "Month-Year"
   MSHFlexGrid1.TextMatrix(1, 5) = "7 Days Income"
   MSHFlexGrid1.TextMatrix(1, 6) = "NNNR- Non NCDEX Non Reservation Income"
   MSHFlexGrid1.TextMatrix(1, 7) = "Other Income"
   MSHFlexGrid1.TextMatrix(1, 8) = "NCDEX Rejected Withdrawal Income"
   MSHFlexGrid1.TextMatrix(1, 9) = "Reservation Income"
   MSHFlexGrid1.TextMatrix(1, 10) = "Spot Income"
   MSHFlexGrid1.TextMatrix(1, 11) = "DP Income"
   MSHFlexGrid1.TextMatrix(1, 12) = "Spillage Income"
   MSHFlexGrid1.TextMatrix(1, 13) = "Cheque Return Charges"
   MSHFlexGrid1.TextMatrix(1, 14) = "SCM Income"
   MSHFlexGrid1.TextMatrix(1, 15) = "Commtrack"
   MSHFlexGrid1.TextMatrix(1, 16) = "NNNR Yearly"
   MSHFlexGrid1.TextMatrix(1, 17) = "PWR"
   MSHFlexGrid1.TextMatrix(1, 18) = "RRN Withdrawal"


   MSHFlexGrid1.TextMatrix(1, 19) = "Total"

   MSHFlexGrid1.ColWidth(0) = 2800
   MSHFlexGrid1.ColWidth(1) = 2800
   MSHFlexGrid1.ColWidth(2) = 2200
   MSHFlexGrid1.ColWidth(3) = 2200
   mstartcol = 4
ElseIf Check1.Value = 1 And (OptWithClient.Value = True Or OptDPIncome.Value = True) Then
   'MSHFlexGrid1.Cols = 18
   MSHFlexGrid1.Cols = 19
   MSHFlexGrid1.TextMatrix(1, 0) = "RM"
   MSHFlexGrid1.TextMatrix(1, 1) = "AM"
   MSHFlexGrid1.TextMatrix(1, 2) = "State"
   MSHFlexGrid1.TextMatrix(1, 3) = "Location"
   MSHFlexGrid1.TextMatrix(1, 4) = "Client Name"
   MSHFlexGrid1.TextMatrix(1, 5) = "Month-Year"
   MSHFlexGrid1.TextMatrix(1, 6) = "7 Days Income"
   MSHFlexGrid1.TextMatrix(1, 7) = "NNNR- Non NCDEX Non Reservation Income"
   MSHFlexGrid1.TextMatrix(1, 8) = "Other Income"
   MSHFlexGrid1.TextMatrix(1, 9) = "NCDEX Rejected Withdrawal Income"
   MSHFlexGrid1.TextMatrix(1, 10) = "Reservation Income"
   MSHFlexGrid1.TextMatrix(1, 11) = "Spot Income"
   MSHFlexGrid1.TextMatrix(1, 12) = "DP Income"
   MSHFlexGrid1.TextMatrix(1, 13) = "Spillage Income"
   MSHFlexGrid1.TextMatrix(1, 14) = "Cheque Return Charges"
   MSHFlexGrid1.TextMatrix(1, 15) = "SCM Income"
   MSHFlexGrid1.TextMatrix(1, 16) = "Commtrack"
   MSHFlexGrid1.TextMatrix(1, 17) = "Total"
   MSHFlexGrid1.TextMatrix(1, 18) = "ClientIDRow"

   MSHFlexGrid1.ColWidth(0) = 2800
   MSHFlexGrid1.ColWidth(1) = 2800
   MSHFlexGrid1.ColWidth(2) = 2200
   MSHFlexGrid1.ColWidth(3) = 2200
   MSHFlexGrid1.ColWidth(4) = 2200
   MSHFlexGrid1.ColWidth(5) = 2200
   mstartcol = 5

Else
   MSHFlexGrid1.Cols = 17
   MSHFlexGrid1.TextMatrix(1, 0) = "Month-Year"
   MSHFlexGrid1.TextMatrix(1, 1) = "7 Days Income"
   MSHFlexGrid1.TextMatrix(1, 2) = "NNNR- Non NCDEX Non Reservation Income"
   MSHFlexGrid1.TextMatrix(1, 3) = "Other Income"
   MSHFlexGrid1.TextMatrix(1, 4) = "NCDEX Rejected Withdrawal Income"
   MSHFlexGrid1.TextMatrix(1, 5) = "Reservation Income"
   MSHFlexGrid1.TextMatrix(1, 6) = "Spot Income"
   MSHFlexGrid1.TextMatrix(1, 7) = "DP Income"
   MSHFlexGrid1.TextMatrix(1, 8) = "Spillage Income"
   MSHFlexGrid1.TextMatrix(1, 9) = "Cheque Return Charges"
   MSHFlexGrid1.TextMatrix(1, 10) = "SCM Income"
   MSHFlexGrid1.TextMatrix(1, 11) = "Commtrack"
   MSHFlexGrid1.TextMatrix(1, 12) = "NNNR Yearly"
   MSHFlexGrid1.TextMatrix(1, 13) = "PWR"
   MSHFlexGrid1.TextMatrix(1, 14) = "RRN Withdrawal"
   MSHFlexGrid1.TextMatrix(1, 15) = "Space Occupacy Charges"
   
   MSHFlexGrid1.TextMatrix(1, 16) = "Total"
   mstartcol = 0
End If

MSHFlexGrid1.TextMatrix(0, 0) = "Total"
If Check2.Value = 1 Then
   MSHFlexGrid1.ColWidth(0 + mstartcol) = 0
Else
   MSHFlexGrid1.ColWidth(0 + mstartcol) = 1500
End If

MSHFlexGrid1.ColWidth(1 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(2 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(3 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(4 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(5 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(6 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(7 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(8 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(9 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(10 + mstartcol) = 1600
MSHFlexGrid1.ColWidth(11 + mstartcol) = 1600

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(1) = 1200

End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame11.Enabled = False
   Exit Sub
End If

Frame11.Enabled = True
Call CmdClear_Click

End Sub


Private Sub AddLocationDetails()
'--- Location,
 If OptWithClient.Value = False And OptDPIncome.Value = False Then
    tmp = "Select ML.LocationID,MS.StateName,ML.LocationName,EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM'"
    tmp = tmp & " From Mst_Location ML"
    tmp = tmp & " inner  join Mst_State MS on MS.StateID=ML.StateID"
    tmp = tmp & " left Join Mst_Employee EMP On ML.AM_EmployeeID = EMP.EmployeeID "
    tmp = tmp & " left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID "
    tmp = tmp & " order by MS.StateName,ML.LocationName"
 Else
    
    tmp = "Select Distinct A.LocationID, A.StateName, A.LocationName, A.ClientIDRow, A.ClientName, A.AM, A.RM,A.AM1, A.RM1"
    tmp = tmp & " From"
    tmp = tmp & " ("
    tmp = tmp & " Select TI.LocationID, MS.StateName,  ML.LocationName, TI.ClientIDRow, MC.ClientName, EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM',  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM1',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM1'"
    tmp = tmp & " From Txn_TempInvoice TI"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TI.LocationID"
    tmp = tmp & " Inner join Mst_State MS on MS.StateID=ML.StateID"
    tmp = tmp & " Left Join Mst_Employee EMP On ML.AM_EmployeeID = EMP.EmployeeID"
    tmp = tmp & " left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
    tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = TI.ClientIDRow"
    tmp = tmp & " Where TI.ClientIDRow is not  NULL and TI.InvoiceDate Between '" & TxtFromDate & "' and '" & TxtToDate & "'"
    tmp = tmp & " Group by TI.LocationID, MS.StateName, ML.LocationName, TI.ClientIDRow, MC.ClientName, EMP.EmployeeName ,EMP.EmployeeNo,EMP1.EmployeeName,EMP1.EmployeeNo"
    tmp = tmp & " Union All"
    tmp = tmp & " Select ML.LocationID, MS.StateName, ML.LocationName, TIG.ClientIDRow, MC.ClientName,  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM',  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM1',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM1'"
    tmp = tmp & " From Txn_InvMonGSLDetail TIG"
    tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TIG.SM_Prefix"
    tmp = tmp & " Inner join Mst_State MS on MS.StateID=ML.StateID"
    tmp = tmp & " Left Join Mst_Employee EMP On ML.AM_EmployeeID = EMP.EmployeeID"
    tmp = tmp & " left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
    tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = TIG.ClientIDRow"
    tmp = tmp & " Where TIG.ClientIDRow is not  NULL and TIG.InvoiceDate Between '" & TxtFromDate & "' and '" & TxtToDate & "'"
    tmp = tmp & " Group by ML.LocationID, MS.StateName, ML.LocationName, TIG.ClientIDRow, MC.ClientName,  EMP.EmployeeName ,EMP.EmployeeNo,EMP1.EmployeeName,EMP1.EmployeeNo"
    tmp = tmp & " Union All"
    tmp = tmp & " Select TRI.LocationID, MS.StateName, ML.LocationName, MSC.ClientIDRow, MC.ClientName, EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM',  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM1',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM1'"
    tmp = tmp & " From Txn_ReservationInvoice TRI"
    tmp = tmp & " Inner Join Mst_StorageContract MSC on MSC.ContractID = TRI.ContractID"
    tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = MSC.ClientIDRow"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TRI.LocationID"
    tmp = tmp & " Inner join Mst_State MS on MS.StateID=ML.StateID"
    tmp = tmp & " Left Join Mst_Employee EMP On ML.AM_EmployeeID = EMP.EmployeeID"
    tmp = tmp & " left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
    tmp = tmp & " Where TRI.YearMonth Between '" & TxtFromDate & "' and '" & TxtToDate & "'"
    tmp = tmp & " Group by TRI.LocationID, MS.StateName, ML.LocationName, MSC.ClientIDRow, MC.ClientName,  EMP.EmployeeName ,EMP.EmployeeNo,EMP1.EmployeeName,EMP1.EmployeeNo"
    tmp = tmp & " Union All"
    tmp = tmp & " Select TC.LocationID, MS.StateName, ML.LocationName, TC.ClientIDRow, MC.ClientName,  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM',  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM1',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM1'"
    tmp = tmp & " From Txn_CDTFBOE TC"
    tmp = tmp & " Inner Join Mst_Client MC on MC.ClientIDRow = TC.ClientIDRow"
    tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TC.SM_Prefix"
    tmp = tmp & " Inner join Mst_State MS on MS.StateID=ML.StateID"
    tmp = tmp & " Left Join Mst_Employee EMP On ML.AM_EmployeeID = EMP.EmployeeID"
    tmp = tmp & " left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
    tmp = tmp & " Where TC.BillingMonth between   '" & TxtFromDate & "' and '" & TxtToDate & "'"
    tmp = tmp & " Group by TC.LocationID, MS.StateName, ML.LocationName, TC.ClientIDRow, MC.ClientName, EMP.EmployeeName ,EMP.EmployeeNo,EMP1.EmployeeName,EMP1.EmployeeNo"
    tmp = tmp & " Union All"
    tmp = tmp & " Select  MDP.LocationID,MS.StateName, ML.LocationName, 0 , MD.DepositoryName,  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM',  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM1',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM1'"
    tmp = tmp & " from Mst_DPInvoiceUpload MDP"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID=MDP.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " Inner Join Mst_Depository MD on MD.NavCode = MDP.DPNavCode and MD.ExchangeTypeID = 2"
    tmp = tmp & " Left Join Mst_Employee EMP On ML.AM_EmployeeID = EMP.EmployeeID"
    tmp = tmp & " Left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
    tmp = tmp & " Where MDP.MonthYear Between '" & TxtFromDate & "' and '" & TxtToDate & "' And isnull(MDP.ReverseFlag,'')<>'Y'"
    tmp = tmp & " Group By MDP.LocationID,MS.StateName, ML.LocationName, MD.DepositoryID, MD.DepositoryName,  EMP.EmployeeName ,EMP.EmployeeNo,EMP1.EmployeeName,EMP1.EmployeeNo"
    tmp = tmp & " Union All"
    tmp = tmp & " Select MDP1.LocationID,MS.StateName, ML.LocationName, 0 , MD.DepositoryName,  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM',  EMP.EmployeeName + '~' + EMP.EmployeeNo as 'AM1',EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'RM1'"
    tmp = tmp & " From Txn_MonthlyCommtrackInvoice MDP1"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID=MDP1.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " Inner Join Mst_Depository MD on MD.NavCode = MDP1.CPNavCode and MD.ExchangeTypeID = 2"
    tmp = tmp & " Left Join Mst_Employee EMP On ML.AM_EmployeeID = EMP.EmployeeID"
    tmp = tmp & " Left Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
    tmp = tmp & " Where MDP1.MonthYear Between '" & TxtFromDate & "' and '" & TxtToDate & "' And Isnull(MDP1.ReverseFlag,'')<>'Y'"
    tmp = tmp & " Group By MDP1.LocationID,MS.StateName, ML.LocationName, MD.DepositoryID, MD.DepositoryName,  EMP.EmployeeName ,EMP.EmployeeNo,EMP1.EmployeeName,EMP1.EmployeeNo"
    tmp = tmp & " )A"
End If
 
Call TmpTable

If Not TmpRs.EOF Then mFlag = True

   tmp = "Select ML.LocationID,ML.LocationName,EMP1.EmployeeName + '~' + EMP1.EmployeeNo as 'CC/AM', " & _
        " EMP2.EmployeeName + '~'+ EMP2.EmployeeNo as 'AM/RM' , " & _
        " EMP3.EmployeeName + '~'+ EMP3.EmployeeNo as 'RM/Managment' , " & _
        " EMP4.EmployeeName + '~'+ EMP4.EmployeeNo as Managment " & _
        " From Mst_EMPCMPMapping ECM " & _
        " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID " & _
        " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID And EMP1.FLag = 'Y' " & _
        " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID And EMP2.FLag = 'Y'  " & _
        " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID And EMP3.FLag = 'Y'  " & _
        " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID And EMP4.FLag = 'Y' " & _
        " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID " & _
        " Inner Join Mst_Designation MD1 On EMP2.DesignationID = MD1.DesignationID " & _
        " Inner Join Mst_Designation MD2 On EMP3.DesignationID = MD2.DesignationID " & _
        " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
        " Where EMP.Flag = 'Y' " & _
        " Group By ML.LocationID,ML.LocationName,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,  " & _
        " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation " & _
        " Order By ML.LocationName"

Call Tmp1Table

If Check2.Value = 1 Then
    For I = 2 To TmpRs.RecordCount + 1
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        If (OptWithClient.Value = True Or OptDPIncome.Value = True) Then
           MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
           MSHFlexGrid1.TextMatrix(I, 18) = TmpRs!ClientIDRow
'           MSHFlexGrid1.TextMatrix(r, 5) = Format(mdate, "dd-MMM-yyyy")
'        Else
'           MSHFlexGrid1.TextMatrix(r, 4) = Format(mdate, "dd-MMM-yyyy")
        End If
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!AM), "", TmpRs!AM)
        
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
        TmpRs.MoveNext
    Next
    '--- For  AM/RM and CC/AM
    
    If TmpRs1.RecordCount > 0 Then
       For I = 2 To MSHFlexGrid1.Rows - 2
           TmpRs1.MoveFirst
           TmpRs1.Find "LocationName='" & MSHFlexGrid1.TextMatrix(I, 3) & "'"
           If TmpRs1.EOF = False Then
              MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1![AM/RM]), "", TmpRs1![AM/RM])
              MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1![CC/AM]), "", TmpRs1![CC/AM])
           End If
       Next
    End If
Else
   For rn = 0 To mTotalMonth
       If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
          For I = 1 To TmpRs.RecordCount
              r = MSHFlexGrid1.Rows - 1
              If (OptWithClient.Value = True Or OptDPIncome.Value = True) Then
                 MSHFlexGrid1.TextMatrix(r, 4) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
                 MSHFlexGrid1.TextMatrix(r, 5) = Format(mDate, "dd-MMM-yyyy")
                 MSHFlexGrid1.TextMatrix(r, 18) = TmpRs!ClientIDRow
              Else
                 MSHFlexGrid1.TextMatrix(r, 4) = Format(mDate, "dd-MMM-yyyy")
              End If
              MSHFlexGrid1.TextMatrix(r, 0) = IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
              MSHFlexGrid1.TextMatrix(r, 1) = IIf(IsNull(TmpRs!AM), "", TmpRs!AM)
              MSHFlexGrid1.TextMatrix(r, 2) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
              MSHFlexGrid1.TextMatrix(r, 3) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
              MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
              TmpRs.MoveNext
          Next
          mDate = DateAdd("m", 1, mDate)
   Next
    '--- For  AM/RM and CC/AM
    If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
    If TmpRs1.RecordCount > 0 Then
       For I = 2 To MSHFlexGrid1.Rows - 2
           TmpRs1.MoveFirst
           TmpRs1.Find "LocationName='" & MSHFlexGrid1.TextMatrix(I, 3) & "'"
           If TmpRs1.EOF = False Then
              MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1![AM/RM]), "", TmpRs1![AM/RM])
              MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1![CC/AM]), "", TmpRs1![CC/AM])
           End If
       Next
    End If
End If
End Sub

