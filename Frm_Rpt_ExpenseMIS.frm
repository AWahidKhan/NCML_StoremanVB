VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_ExpenseMIS 
   Caption         =   "Expense MIS"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   9150
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   " Select Date Range "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton CmdExcelReport 
         Caption         =   "Excel Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   1320
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
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
         CustomFormat    =   "MMM- yyyy"
         Format          =   20709379
         CurrentDate     =   40098
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   480
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
         CustomFormat    =   "MMM- yyyy"
         Format          =   20709379
         CurrentDate     =   40098
      End
      Begin VB.Label Label2 
         Caption         =   "To"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "From"
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
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmRpt_ExpenseMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FromMonth As Integer, ToMonth As Integer, FromYear As Integer, ToYear As Integer
Dim BindAtEnd As Boolean, StartFromFirst As Boolean
Dim MonthList(12) As String
Private Sub SetMonth()
    MonthList(1) = "JAN"
    MonthList(2) = "FEB"
    MonthList(3) = "MAR"
    MonthList(4) = "APR"
    MonthList(5) = "MAY"
    MonthList(6) = "JUN"
    MonthList(7) = "JUL"
    MonthList(8) = "AUG"
    MonthList(9) = "SEP"
    MonthList(10) = "OCT"
    MonthList(11) = "NOV"
    MonthList(12) = "DEC"
End Sub

Private Sub CmdExcelReport_Click()

If TxtToDate.Value < TxtFromDate.Value Then
    MsgBox "Please Select Proper Date Range!", vbInformation
    TxtFromDate.SetFocus
    Exit Sub
End If
BindAtEnd = False
FromMonth = Month(TxtFromDate.Value)
ToMonth = Month(TxtToDate.Value)
FromYear = Year(TxtFromDate.Value)
ToYear = Year(TxtToDate.Value)
If FromYear <> ToYear Then
    ToMonth = 12 + ToMonth
Else
    If FromMonth <= 3 Then
        BindAtEnd = True
    End If
End If

'Filename Format-- Rpt_ExpenseMIS_By_Username_Time.xls
Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "Expense MIS.xls")
'Call Gen_Create_Xls("RptChart_ExpenseListing.xls")
Call Gen_Xls_Report
End Sub
Public Sub Create_Xls(FileName As String)
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FLName = FileName
oBook.SaveAs (Pat & "\Excel\" & FileName)

Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)

oExcel.Quit
'oExcel.Close
End Sub
Private Sub Gen_Xls_Report()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer, mCol As Integer
Dim RowCount As Integer
Dim ExpenseTypeCount As Integer, LocationID() As Integer, ExpenseTypeID() As Integer
Dim LocationName() As String
Dim Range1 As Excel.Range
'MsgBox ("Wait till Next Message")
Screen.MousePointer = vbHourglass
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path

'Set oWB = oXL.Workbooks.Open(Pat & "\Excel\RptChart_ExpenseListing.xls")
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Expense MIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
oXL.Visible = True         ' Show the excel sheet
Call SetMonth
With oXL.ActiveSheet
        ' Print the heading and columns first
        .Cells(1, 3).Value = "EXPENCE LISTING FOR YEAR 2009"
        .Cells(1, 3).Font.Name = "Verdana"
        .Cells(1, 3).Font.Size = 12
        .Cells(1, 3).Font.Bold = True:
        .Cells(4, 1).Value = "LOCATION":        .Cells(4, 2).Value = "EXPENSE TYPE"
        .Cells(4, 3).Value = "TOTAL AMOUNT"
End With
Dim HeadSet As Integer
HeadSet = 4
For i = FromMonth To ToMonth
    If i > 12 Then
       
        oXL.ActiveSheet.Cells(4, HeadSet) = MonthList(i - 12) & "-" & ToYear
    Else
        
        oXL.ActiveSheet.Cells(4, HeadSet) = MonthList(i) & "-" & FromYear
        
    End If
    HeadSet = HeadSet + 1
Next
'If BindAtEnd = False Then
'    StartFromFirst = False
'    With oXL.ActiveSheet
'        .Cells(4, 4).Value = "APR" & "-" & FromYear
'        .Cells(4, 5).Value = "MAY" & "-" & FromYear:            .Cells(4, 6).Value = "JUN" & "-" & FromYear
'        .Cells(4, 7).Value = "JUL" & "-" & FromYear:            .Cells(4, 8).Value = "AUG" & "-" & FromYear
'        .Cells(4, 9).Value = "SEP" & "-" & FromYear:            .Cells(4, 10).Value = "OCT" & "-" & FromYear
'        .Cells(4, 11).Value = "NOV" & "-" & FromYear:           .Cells(4, 12).Value = "DEC" & "-" & FromYear
'        .Cells(4, 13).Value = "JAN" & "-" & ToYear:           .Cells(4, 14).Value = "FEB" & "-" & ToYear
'        .Cells(4, 15).Value = "MAR" & "-" & ToYear
'    End With
'Else
'    StartFromFirst = True
'    With oXL.ActiveSheet
'        .Cells(4, 4).Value = "JAN" & "-" & FromYear
'        .Cells(4, 5).Value = "FEB" & "-" & FromYear:            .Cells(4, 6).Value = "MAR" & "-" & FromYear
'        .Cells(4, 7).Value = "APR" & "-" & FromYear:            .Cells(4, 8).Value = "MAY" & "-" & FromYear
'        .Cells(4, 9).Value = "JUN" & "-" & FromYear:            .Cells(4, 10).Value = "JUL" & "-" & FromYear
'        .Cells(4, 11).Value = "AUG" & "-" & FromYear:           .Cells(4, 12).Value = "SEP" & "-" & FromYear
'        .Cells(4, 13).Value = "OCT" & "-" & FromYear:           .Cells(4, 14).Value = "NOV" & "-" & FromYear
'        .Cells(4, 15).Value = "DEC" & "-" & FromYear
'    End With
'End If
' Merge Columns
Set Range1 = oWS.Range("A1:O1")
Range1.Select
Selection.Merge

'Set Column Width
oXL.Columns("A:A").ColumnWidth = 14.43
oXL.Columns("B:B").ColumnWidth = 25.57
oXL.Columns("C:C").ColumnWidth = 16.43

'Setting Column Head Font Property
For j = 1 To 15
oXL.ActiveSheet.Cells(4, j).Font.Name = "Verdana"
oXL.ActiveSheet.Cells(4, j).Font.Bold = True
Next


mRow = 5
'Find All Location
tmp = "Select Distinct L.LocationName,L.LocationID From Txn_ClaimSubmission C,Mst_Location L where C.LocationID=L.LocationID order by L.LocationName "
Call Tmp0Table
Dim L As Integer
L = TmpRs0.RecordCount
ReDim LocationName(0 To L) As String
ReDim LocationID(0 To L) As Integer

For a = 1 To L
    LocationName(a) = TmpRs0!LocationName
    LocationID(a) = TmpRs0!LocationID
    TmpRs0.MoveNext
Next

'Find All Expense Type
tmp = "Select Distinct E.ExpenseType,E.ExpenseTypeID"
tmp = tmp & " from Txn_ClaimSubmissionDetail C,Mst_ExpenseType E where C.ExpenseTypeID=E.ExpenseTypeID"
tmp = tmp & " Order by E.ExpenseTypeID"
Call TmpTable
RowCount = L * TmpRs.RecordCount
ExpenseTypeCount = TmpRs.RecordCount
ReDim ExpenseTypeID(0 To ExpenseTypeCount) As Integer

'Write All Location And Expense Type To Excel Sheet
For b = 1 To L
    TmpRs.MoveFirst
    For E = 1 To TmpRs.RecordCount
        oXL.ActiveSheet.Cells(mRow, 1).Value = LocationName(b) ' Location
        oXL.ActiveSheet.Cells(mRow, 2).Value = TmpRs!ExpenseType ' Expense Type
        ExpenseTypeID(E) = TmpRs!ExpenseTypeID
        TmpRs.MoveNext
        mRow = mRow + 1
    Next
    
Next

'Calculating Sum of Amount For Per Month
mRow = 5
'For c = 1 To RowCount
For D = 1 To L
    For f = 1 To ExpenseTypeCount
    HeadSet = 4
        For g = FromMonth To ToMonth
            If g > 12 Then
                vMonth = ToMonth - 12
                ToYear = Year(TxtToDate.Value)
            Else
                vMonth = g
                ToYear = Year(TxtFromDate.Value)
            End If
            tmp = "Select Sum(M.Amount) as Amount"
            tmp = tmp & " From Txn_ClaimSubmission M ,Txn_ClaimSubmissionDetail D"
            tmp = tmp & " Where M.LocationID = " & LocationID(D) & " And D.ExpenseTypeID = " & ExpenseTypeID(f) & " And Month(m.ExpenseFrom) = " & vMonth & " and Year(M.ExpenseFrom)=" & ToYear & ""
            Call TmpTable
'            If g <= 3 Or BindAtEnd = True Then
'                If StartFromFirst = True Then
'                    oXL.ActiveSheet.Cells(mRow, g + 3).Value = IIf(IsNull(TmpRs!Amount), "0", TmpRs!Amount)
'                Else
'                    oXL.ActiveSheet.Cells(mRow, g + 12).Value = IIf(IsNull(TmpRs!Amount), "0", TmpRs!Amount)
'                End If
'            Else
                oXL.ActiveSheet.Cells(mRow, HeadSet).Value = IIf(IsNull(TmpRs!Amount), "0", TmpRs!Amount)
'            End If
            'mRow = mRow + 1
            HeadSet = HeadSet + 1
        Next
        With Worksheets.Application.ActiveSheet
        Set myrange = .Range(.Cells(mRow, 4), .Cells(mRow, 15))
        End With
        oXL.ActiveSheet.Cells(mRow, 3).Value = WorksheetFunction.Sum(myrange)
        mRow = mRow + 1
    Next
Next
'Next
BindAtEnd = False
StartFromFirst = False
'Find Total Amount
'Range1 = Nothing
'For h = 5 To RowCount + 4
'    With Worksheets.Application.ActiveSheet
'    Set myrange = .Range(.Cells(h, 4), .Cells(h, 15))
'    End With
'    oXL.ActiveSheet.Cells(h, 3).Value = WorksheetFunction.Sum(myrange)
'Next
'oXL.Visible = True
'oWB.Save
Screen.MousePointer = 0
oWB.Save
MsgBox ("Excel File Created !!! ")



''''''For i = 1 To 12 'For Per Month Total Amout
''''''
''''''    tmp = "SELECT  L.LocationName,E.ExpenseType,sum(M.Amount) as Amount"
''''''    tmp = tmp & "  From Txn_ClaimSubmission M INNER JOIN Mst_Location L ON M.LocationID=L.LocationID ,Txn_ClaimSubmissionDetail D INNER JOIN Mst_ExpenseType E ON D.ExpenseTypeID=E.ExpenseTypeID"
''''''    tmp = tmp & "  Where Month(m.ExpenseFrom) = " & i & ""
''''''    tmp = tmp & "  Group By L.LocationName,E.ExpenseType"
''''''    tmp = tmp & "  Order By L.LocationName"
''''''
''''''    Call TmpTable
''''''    If TmpRs.EOF <> True Then
''''''        mRow = 5
''''''        For rs = 1 To TmpRs.RecordCount
'''''''        If oXL.ActiveSheet.Cells(mRow, 1).Value = TmpRs!LocationName Then
''''''
''''''        oXL.ActiveSheet.Cells(mRow, 1).Value = TmpRs!LocationName ' Location
''''''        oXL.ActiveSheet.Cells(mRow, 2).Value = TmpRs!ExpenseType ' Expense Type
''''''        oXL.ActiveSheet.Cells(mRow, i + 3).Value = TmpRs!Amount ' Amount
''''''        mRow = mRow + 1
''''''        TmpRs.MoveNext
''''''        Next
''''''    End If
''''''Next
'oXL.Visible = True
'oXL.SaveAs "C:/Test.xls"

' Now we will export data into the sheet
'Do Until rsGetAllData.EOF
'    For lc = 1 To MSHFlexGrid1.Rows - 1
'        oXL.ActiveSheet.Cells(mRow, lc).Value = MSHFlexGrid1.TextMatrix(lc, 5) ' Location
'        oXL.ActiveSheet.Cells(mRow, lc + 1).Value = MSHFlexGrid1.TextMatrix(lc, 9) 'Expense Type
'        oXL.ActiveSheet.Cells(mRow, lc + 1).Value = MSHFlexGrid1.TextMatrix(lc, 9)
'        mRow = mRow + 1
'    Next
'    rsGetAllData.MoveNext
'Loop
'Cleanup:
'    Set oWS = Nothing
'    If Not oWB Is Nothing Then oWB.Close
'    Set oWB = Nothing
    'oXL.Quit
    'Set oExcel = Nothing

End Sub
