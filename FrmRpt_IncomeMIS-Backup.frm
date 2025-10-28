VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_IncomeMIS 
   Caption         =   "Income MIS Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   135
      TabIndex        =   2
      Top             =   0
      Width           =   15015
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   855
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_IncomeMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GunVer As String
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
MSHFlexGrid1.Cols = 11

MSHFlexGrid1.TextMatrix(0, 0) = "RM"
MSHFlexGrid1.TextMatrix(0, 1) = "AM"
MSHFlexGrid1.TextMatrix(0, 2) = "State"
MSHFlexGrid1.TextMatrix(0, 3) = "LocationID"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "Capacity"
MSHFlexGrid1.TextMatrix(0, 6) = "Total Income"
MSHFlexGrid1.TextMatrix(0, 7) = "Direct Expense"
MSHFlexGrid1.TextMatrix(0, 8) = "Salary Expense "
MSHFlexGrid1.TextMatrix(0, 9) = "Telephone Expense"
MSHFlexGrid1.TextMatrix(0, 10) = "Corporate Office Expense"

MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(1) = 1800
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 0
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 2600
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 2200
MSHFlexGrid1.ColWidth(10) = 2200
'MSHFlexGrid1.ColWidth(8) = 1600

End Sub

Private Sub CmdGetReport_Click()
GunVer = ""
MsgBox "Wait till next msg"

tmp = "Select * from Mst_Employee " & _
      " Where EmployeeNo Not In (Select EmployeeCode From Txn_EmpCTC) " & _
      " Order By EmployeeNo"

Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   MsgBox "Some Employee data from Store-Man Missing in Employee CTC Data!!"
   Call Grid_Head
   Exit Sub
End If

tmp = "Select MS.StateName,ML.LocationID,ML.LocationName From Mst_Location ML " & _
      " Inner Join Mst_State MS On ML.StateID = MS.StateID " & _
      " Order By ML.LocationName "
'Call TableMst_Location(" Order By LocationName ")
Call TmpTable
Call displocationdata

MsgBox "Done!!!"

End Sub
Private Sub GetRMAMforLocation(mLocationName_ As Variant, X1 As Double)
Dim strSplitString() As String
Dim EmpNo_, EmpName_, EmpDesig_ As String

TmpRs3.MoveFirst
TmpRs3.Filter = "LocationName = '" & mLocationName_ & "'"
If TmpRs3.EOF = False Then
   If IsNull(TmpRs3.Fields(2)) = False Then
      strSplitString = Split(TmpRs3.Fields(2), "~")
      EmpName_ = Trim(strSplitString(0))
      EmpNo_ = Trim(strSplitString(1))
      EmpDesig_ = Trim(strSplitString(2))

      If LCase(EmpDesig_) = "am" Then
         MSHFlexGrid1.TextMatrix(X1, 1) = IIf(IsNull(TmpRs3.Fields(2)), "", TmpRs3.Fields(2))
      End If
   End If
   
   If IsNull(TmpRs3.Fields(3)) = False Then
      strSplitString = Split(TmpRs3.Fields(3), "~")
      EmpName_ = Trim(strSplitString(0))
      EmpNo_ = Trim(strSplitString(1))
      EmpDesig_ = Trim(strSplitString(2))
      If LCase(EmpDesig_) = "rm" Then
         MSHFlexGrid1.TextMatrix(X1, 0) = IIf(IsNull(TmpRs3.Fields(3)), "", TmpRs3.Fields(3))
      ElseIf LCase(EmpDesig_) = "am" Then
         MSHFlexGrid1.TextMatrix(X1, 1) = IIf(IsNull(TmpRs3.Fields(3)), "", TmpRs3.Fields(3))
      End If
   End If
      
   If IsNull(TmpRs3.Fields(4)) = False Then
      strSplitString = Split(TmpRs3.Fields(4), "~")
      EmpName_ = Trim(strSplitString(0))
      EmpNo_ = Trim(strSplitString(1))
      EmpDesig_ = Trim(strSplitString(2))
      If LCase(EmpDesig_) = "rm" Then
         MSHFlexGrid1.TextMatrix(X1, 0) = IIf(IsNull(TmpRs3.Fields(4)), "", TmpRs3.Fields(4))
      End If
   End If
   
   If TmpRs3.RecordCount > 1 Then
      MSHFlexGrid1.TextMatrix(X1, 0) = "*" & MSHFlexGrid1.TextMatrix(X1, 0)
   End If
End If

TmpRs3.Filter = ""

End Sub

Private Sub displocationdata()
Dim x As Double


tmp = "Select ML.LocationID,ML.LocationName,EMP1.EmployeeName +'~'+ EMP1.EmployeeNo +'~'+ MD.Designation as 'CC/AM', " & _
      " EMP2.EmployeeName+'~'+EMP2.EmployeeNo +'~'+ MD1.Designation as 'AM/RM' , " & _
      " EMP3.EmployeeName +'~'+EMP3.EmployeeNo +'~'+ MD2.Designation as 'RM/Managment' , " & _
      " EMP4.EmployeeName +'~'+EMP4.EmployeeNo as Managment " & _
      " From Mst_EMPCMPMapping ECM " & _
      " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID " & _
      " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID " & _
      " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID " & _
      " Inner Join Mst_Designation MD1 On EMP2.DesignationID = MD1.DesignationID " & _
      " Inner Join Mst_Designation MD2 On EMP3.DesignationID = MD2.DesignationID " & _
      " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
      " Group By ML.LocationID,ML.LocationName,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,  " & _
      " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation " & _
      " Order By ML.LocationName"

Call Tmp3Table


x = 1
If TmpRs.RecordCount > 0 Then
   For i = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows '- 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(x, c + 2) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       Call GetRMAMforLocation(TmpRs!Locationname, x)
       Call CapacityData(TmpRs!Locationname, x)
       Call TotalIncomeLocWise(TmpRs!Locationname, x)
       x = MSHFlexGrid1.Rows '- 1
       TmpRs.MoveNext
   Next
   Call DirectIncomeLocWise

   
   Call TotalSalaryLocWise
   
   Call TotalTelephoneLocWise
   
   
   CmdExcel.Enabled = True

Else
   Call Grid_Head
End If

If MSHFlexGrid1.Rows > 1 Then
   MSHFlexGrid1.FixedRows = 1
   MSHFlexGrid1.FixedCols = 4
End If

End Sub

Private Sub CapacityData(mLocationName_ As Variant, X1 As Double)

tmp = "Select ML.LocationName,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML " & _
      " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
      " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
      " Where ML.LocationName = '" & mLocationName_ & "' And MG.CloseDate is Null " & _
      " Group By ML.LocationName "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(X1, 5) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
End If

TmpRs1.Close
End Sub

Private Sub TotalIncomeLocWise(mLocationName_ As Variant, X1 As Double)

tmp = "Select ML.LocationName ,Sum(TI.AmtInclOfSerTax) As TotalIncome From Txn_IncomeDetails TI " & _
      " Right Outer Join Mst_Location ML On TI.LocationID = ML.LocationID " & _
      " Where LocationName = '" & mLocationName_ & "' " & _
      " Group By LocationName "

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(X1, 6) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
End If

TmpRs1.Close

End Sub
Private Sub DirectIncomeLocWise()

tmp = "Select ME.EmployeeID,ME.EmployeeNo + '~' + ME.EmployeeName as Employee,MD.Designation, Sum(TCS.Amount) As TotalExpenseAmount,  Me.AllocateExpense " & _
      " From Txn_ClaimSubmission TCS " & _
      " Inner Join Mst_Employee ME On TCS.EmployeeID = ME.EmployeeID " & _
      " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID " & _
      " Where TCS.ExpenseFrom <= '2009-09-30' " & _
      " Group By ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,ME.AllocateExpense,MD.Designation "

Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
    Call CalDirectIncomeLocWise("DirectIncome", 7, TmpRs1!EmployeeID, TmpRs1!Designation, TmpRs1!AllocateExpense, TmpRs1!TotalExpenseAmount)
    TmpRs1.MoveNext
Next

End Sub
Private Sub CalDirectIncomeLocWise(mType_ As Variant, mRow_ As Variant, mEMPID_ As Variant, mEMPDesignation_ As Variant, mAllocateExpense_ As Variant, mExpenseAmount_ As Variant)
Dim mAmtAfterDeduction, mCorporateExpense, mCorporateExpenseAfterDed, mEMPIDDummy_ As Variant

mEMPIDDummy_ = mEMPID_
mCorporateExpense = (30000000 / 2)
'mAmtAfterDeduction = mExpenseAmount_ * (Gen_Percentage / 100)

mAllocateExpense_ = IIf(IsNull(mAllocateExpense_), 0, mAllocateExpense_)

If LCase(mEMPDesignation_) = "managment" Or LCase(mEMPDesignation_) = "rag" Then
   tmp = "Select Sum(MG.GodownCapacity) as TotalLocationCapacity From Mst_Location ML " & _
         " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
         " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
         " Where MG.CloseDate is Null"
Else
   If LCase(mEMPDesignation_) = "qqs" Or LCase(mEMPDesignation_) = "drc" Then
      tmp = "Select DirectReportingTo From Mst_Employee Where EmployeeID = " & mEMPIDDummy_ & ""
      Call Tmp2Table
      mEMPIDDummy_ = IIf(IsNull(TmpRs2!DirectReportingTo), 0, TmpRs2!DirectReportingTo)
   End If
   tmp = "Select Sum(MG.GodownCapacity) as TotalLocationCapacity " & _
         " From Mst_EMPCMPMapping ECM " & _
         " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID " & _
         " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP6 On EMP.ConnectedToRAG = EMP6.EmployeeID " & _
         " Left Outer Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
         " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
         " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
         " Where (EMP.EmployeeID = " & mEMPIDDummy_ & " Or EMP1.EmployeeID = " & mEMPIDDummy_ & " Or EMP2.EmployeeID = " & mEMPIDDummy_ & " Or EMP3.EmployeeID = " & mEMPIDDummy_ & " Or EMP4.EmployeeID = " & mEMPIDDummy_ & " Or EMP6.EmployeeID = " & mEMPIDDummy_ & ") " & _
         " And MG.CloseDate is Null"

End If

Call Tmp2Table

If LCase(mEMPDesignation_) = "managment" Or LCase(mEMPDesignation_) = "rag" Then
   tmp = "Select ML.LocationName,Sum(MG.GodownCapacity) as TotalLocationCapacity From Mst_Location ML " & _
         " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
         " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
         " Where MG.CLoseDate is Null " & _
         " Group By ML.LocationName "
Else
'   If LCase(mEMPDesignation_) = "qqs" Or LCase(mEMPDesignation_) = "drc" Then
'      tmp = "Select DirectReportingTo From Mst_Employee Where EmployeeID = " & mEMPIDDummy_ & ""
'      Call TmpTable
'      If TmpRs.RecordCount <> 0 Then
'         mEMPIDDummy_ = IIf(IsNull(TmpRs!DirectReportingTo), 0, TmpRs!DirectReportingTo)
'      End If
'   End If
   tmp = "Select ECM.LocationID,ML.LocationName,Sum(MG.GodownCapacity) as TotalLocationCapacity " & _
         " From Mst_EMPCMPMapping ECM " & _
         " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID " & _
         " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP6 On EMP.ConnectedToRAG = EMP6.EmployeeID " & _
         " Left Outer Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
         " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
         " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
         " Where (EMP.EmployeeID = " & mEMPIDDummy_ & " Or EMP1.EmployeeID = " & mEMPIDDummy_ & " Or EMP2.EmployeeID = " & mEMPIDDummy_ & " Or EMP3.EmployeeID = " & mEMPIDDummy_ & " Or EMP4.EmployeeID = " & mEMPIDDummy_ & " Or EMP6.EmployeeID = " & mEMPIDDummy_ & ") And " & _
         " MG.CloseDate is Null " & _
         " Group By ECM.LocationID,ML.LocationName " & _
         " Order By ECM.LocationID "
End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
   If LCase(mType_) = "salary" Then
      mExpenseAmount_ = mExpenseAmount_ / 2
   End If
   If mAllocateExpense_ = 1 Then
      mAmtAfterDeduction = mExpenseAmount_ * (Gen_Percentage / 100)
   Else
      mAmtAfterDeduction = mExpenseAmount_
   End If
   mCorporateExpenseAfterDed = mCorporateExpense / TmpRs2!TotalLocationCapacity
   mAmtAfterDeduction = mAmtAfterDeduction / TmpRs2!TotalLocationCapacity
   For i = 1 To TmpRs.RecordCount
       For ix = 1 To MSHFlexGrid1.Rows - 1
           If MSHFlexGrid1.TextMatrix(ix, 4) = TmpRs!Locationname Then
              If TmpRs!Locationname = "Gulbarga" Then
                
                 If LCase(mType_) = "salary" Then
                    If mEMPID_ = 247 Or mEMPID_ = 31 Or mEMPID_ = 16 Or mEMPID_ = 103 Then
                        tmp = tmp
                    End If
                    GunVer = GunVer & TmpRs2!TotalLocationCapacity & "," & mEMPID_ & " , " & (MSHFlexGrid1.TextMatrix(ix, 5)) & "," & mAmtAfterDeduction & "," & (Val(MSHFlexGrid1.TextMatrix(ix, mRow_)) * mAmtAfterDeduction) & Chr(13)
                 End If
              End If
              MSHFlexGrid1.TextMatrix(ix, mRow_) = Val(MSHFlexGrid1.TextMatrix(ix, mRow_)) + (Val(MSHFlexGrid1.TextMatrix(ix, 5)) * mAmtAfterDeduction)
              MSHFlexGrid1.TextMatrix(ix, mRow_) = Round(MSHFlexGrid1.TextMatrix(ix, mRow_), 3)
              If LCase(mEMPDesignation_) = "managment" And LCase(mType_) = "directincome" Then
                 MSHFlexGrid1.TextMatrix(ix, 10) = (Val(MSHFlexGrid1.TextMatrix(ix, 5)) * mCorporateExpenseAfterDed)
                 MSHFlexGrid1.TextMatrix(ix, 10) = Round(MSHFlexGrid1.TextMatrix(ix, 10), 3)
                 Exit For
              End If
           End If
       Next
       TmpRs.MoveNext
   Next
End If




TmpRs2.Close

End Sub
Private Sub TotalSalaryLocWise()

tmp = " Select ME.EmployeeID,TC.EmployeeCode,MD.Designation,ME.AllocateExpense,TC.CTC_Yearly " & _
      " From Txn_EmpCTC TC " & _
      " Inner Join Mst_Employee ME On TC.EmployeeCode = ME.EmployeeNo " & _
      " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID " & _
      " Where EmployeeCode In (Select EmployeeNo From Mst_Employee ) " & _
      " Order By EmployeeCode"

Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
    Call CalDirectIncomeLocWise("Salary", 8, TmpRs1!EmployeeID, TmpRs1!Designation, TmpRs1!AllocateExpense, TmpRs1!CTC_YEARLY)
    TmpRs1.MoveNext
Next
MsgBox GunVer

End Sub
Private Sub TotalTelephoneLocWise()

tmp = "Select ME.EmployeeID,MD.Designation,ME.AllocateExpense,TT.EmpNo,Sum(TT.MonthlyBillAmt) as MonthlyBillAmt,Sum(TT.Excess) as Excess,Sum(TT.MonthlyBillAmt - TT.Excess) as NCMSLBillAmt " & _
      " From Txn_TelephoneDetails TT " & _
      " Inner Join Mst_Employee ME On TT.EmpNo = ME.EmployeeNo " & _
      " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID " & _
      " Where EmpNo In (Select EmployeeNo From Mst_Employee ) " & _
      " Group By ME.EmployeeID,EmpNo,MD.Designation,ME.AllocateExpense " & _
      " Order By EmpNo"

Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
    Call CalDirectIncomeLocWise("Tele", 9, TmpRs1!EmployeeID, TmpRs1!Designation, TmpRs1!AllocateExpense, TmpRs1!MonthlyBillAmt)
    TmpRs1.MoveNext
Next


End Sub
'Private Sub TotalCorporateExpenseLocWise()
'
'tmp = "Select ME.EmployeeID,MD.Designation,ME.AllocateExpense,TT.EmpNo,Sum(TT.MonthlyBillAmt) as MonthlyBillAmt,Sum(TT.Excess) as Excess,Sum(TT.MonthlyBillAmt - TT.Excess) as NCMSLBillAmt " & _
'      " From Txn_TelephoneDetails TT " & _
'      " Inner Join Mst_Employee ME On TT.EmpNo = ME.EmployeeNo " & _
'      " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID " & _
'      " Where EmpNo In (Select EmployeeNo From Mst_Employee ) " & _
'      " Group By ME.EmployeeID,EmpNo,MD.Designation,ME.AllocateExpense " & _
'      " Order By EmpNo"
'
'Call Tmp1Table
'
'For i1 = 1 To TmpRs1.RecordCount
'    Call CalDirectIncomeLocWise("Tele", 6, TmpRs1!EmployeeID, TmpRs1!Designation, TmpRs1!AllocateExpense, TmpRs1!MonthlyBillAmt)
'    TmpRs1.MoveNext
'Next
'
'
'End Sub

Private Sub Form_Load()

CmdExcel.Enabled = False
Gen_Percentage = 100
Call Grid_Head

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Rpt_IncomeMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Rpt_IncomeMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
