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
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   10398
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   6240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4048
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
Dim gmTotCap As Double
Dim TotExp As Double
Dim mMonYr As Double

Private Sub Grid_HeadEmp()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 1
MSHFlexGrid2.Cols = 11

MSHFlexGrid2.TextMatrix(0, 0) = "Emp ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Emp No"
MSHFlexGrid2.TextMatrix(0, 2) = "Employee Name"
MSHFlexGrid2.TextMatrix(0, 3) = "Designation ID"
MSHFlexGrid2.TextMatrix(0, 4) = "Designation"
MSHFlexGrid2.TextMatrix(0, 5) = "Allocation"
MSHFlexGrid2.TextMatrix(0, 6) = "Direct Reporting"
MSHFlexGrid2.TextMatrix(0, 7) = "Total Capacity"
MSHFlexGrid2.TextMatrix(0, 8) = "Expense"
MSHFlexGrid2.TextMatrix(0, 9) = "Salary"
MSHFlexGrid2.TextMatrix(0, 10) = "Telephone"

MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 1200
MSHFlexGrid2.ColWidth(2) = 1500
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 1200
MSHFlexGrid2.ColWidth(5) = 1000
MSHFlexGrid2.ColWidth(7) = 1500

tmp = "select a.EmployeeID,a.EmployeeNo,a.EmployeeName,a.DesignationID,b.Designation,a.AllocateExpense,a.DirectReportingTo " & _
"from mst_employee a inner join mst_Designation b on a.DesignationID = b.DesignationID " & _
"order by b.Designation,a.EmployeeName"
'MSHFlexGrid1.ColWidth(8) = 1600
Call TmpTable
MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For c = 0 To 6 'MSHFlexGrid2.Cols - 2
        MSHFlexGrid2.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    
    ' Old Code
'    If LCase(Trim(TmpRs!Designation)) = "am" Then
'       tmp = "select sum(GodownCapacity) From " & _
'             " Mst_Godown MG where CMPID in( select ecm.CMPID from Mst_Employee emp " & _
'             "inner join Mst_EMPCMPMapping ecm on emp.EmployeeID = ecm.EmployeeID " & _
'             "inner Join Mst_Location ML On ecm.LocationID = ML.LocationID " & _
'             "inner Join Mst_CMP MC On ecm.CMPID = MC.CMPID " & _
'             "Where emp.EmployeeID in(select employeeid  from " & _
'             "Mst_Employee  where DirectReportingTo in(" & TmpRs!EmployeeID & ")))  And MG.CloseDate is Null "
'        Call Tmp0Table
'        MSHFlexGrid2.TextMatrix(i, 7) = IIf(IsNull(TmpRs0.Fields(0)), 0, TmpRs0.Fields(0))
'    End If
'    If LCase(Trim(TmpRs!Designation)) = "cmpm" Then
'        tmp = "select sum(GodownCapacity)  From Mst_Employee emp " & _
'              "inner join Mst_EMPCMPMapping ecm on emp.EmployeeID = ecm.EmployeeID " & _
'              "inner Join Mst_Location ML On ecm.LocationID = ML.LocationID " & _
'              "inner Join Mst_CMP MC On ecm.CMPID = MC.CMPID " & _
'              "Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
'              "Where emp.EmployeeID in(" & TmpRs!EmployeeID & ")  And MG.CloseDate is Null "
'        Call Tmp0Table
'        MSHFlexGrid2.TextMatrix(i, 7) = IIf(IsNull(TmpRs0.Fields(0)), 0, TmpRs0.Fields(0))
'    End If
    
    If LCase(Trim(TmpRs!Designation)) = "am" Then
       tmp = "select sum(MG.ClosingWeight) From " & _
             " mst_stockstatus MG where CMPID in( select ecm.CMPID from Mst_Employee emp " & _
             "inner join Mst_EMPCMPMapping ecm on emp.EmployeeID = ecm.EmployeeID " & _
             "inner Join Mst_Location ML On ecm.LocationID = ML.LocationID " & _
             "inner Join Mst_CMP MC On ecm.CMPID = MC.CMPID " & _
             "Where emp.EmployeeID in(select employeeid  from " & _
             "Mst_Employee  where DirectReportingTo in(" & TmpRs!EmployeeID & ")))   "
        Call Tmp0Table
        MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs0.Fields(0)), 0, TmpRs0.Fields(0))
    End If
    If LCase(Trim(TmpRs!Designation)) = "cmpm" Then
        tmp = "select sum(MG.ClosingWeight)  From Mst_Employee emp " & _
              "inner join Mst_EMPCMPMapping ecm on emp.EmployeeID = ecm.EmployeeID " & _
              "inner Join Mst_Location ML On ecm.LocationID = ML.LocationID " & _
              "inner Join Mst_CMP MC On ecm.CMPID = MC.CMPID " & _
              "Inner Join mst_stockstatus MG On MC.CMPID = MG.CMPID " & _
              "Where emp.EmployeeID in(" & TmpRs!EmployeeID & ")   "
        Call Tmp0Table
        MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs0.Fields(0)), 0, TmpRs0.Fields(0))
    End If

    TmpRs.MoveNext
Next
Dim mCapacity  As Double
Dim mDesig As String

For I = 1 To MSHFlexGrid2.Rows - 1
    mDesig = LCase(Trim(MSHFlexGrid2.TextMatrix(I, 4)))
    TmpRs.MoveFirst
    If mDesig = "cmpm" Then
    
    End If
    
  
    If mDesig = "am" Or mDesig = "rm" Then
       TmpRs.Find "EmployeeID = " & MSHFlexGrid2.TextMatrix(I, 0)
       If Not TmpRs.EOF Then
          Call FixCapacity(TmpRs!DirectReportingTo, Val(MSHFlexGrid2.TextMatrix(I, 7)))
       End If
    End If
    If mDesig = "cc" Or mDesig = "drc" Or mDesig = "bdo" Or mDesig = "qqs" Then
       TmpRs.Find "EmployeeID = " & MSHFlexGrid2.TextMatrix(I, 0)
       If Not TmpRs.EOF Then
          mCapacity = GetCapacity(IIf(IsNull(TmpRs!DirectReportingTo), 0, TmpRs!DirectReportingTo))
           Call FixCapacity(MSHFlexGrid2.TextMatrix(I, 0), mCapacity)
       End If
    End If
    
Next

If MSHFlexGrid2.Rows > 1 Then
   MSHFlexGrid2.FixedRows = 1
End If
End Sub
Private Sub FixCapacity(EmpID_ As Double, Cap_ As Double)
For o = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(o, 0) = EmpID_ Then
       MSHFlexGrid2.TextMatrix(o, 7) = Val(MSHFlexGrid2.TextMatrix(o, 7)) + Cap_
       Exit For
    End If
Next
End Sub
Private Sub FixExpense(EmpID_ As Double, Exp_ As Double, Col_ As Double)
For o = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(o, 0) = EmpID_ Then
       MSHFlexGrid2.TextMatrix(o, Col_) = Val(MSHFlexGrid2.TextMatrix(o, Col_)) + Exp_
       Exit For
    End If
Next
End Sub


Function GetCapacity(EmpID_ As Double) As Double
Dim ValRet As Double
ValRet = 0
For o = 1 To MSHFlexGrid2.Rows - 1
    If Val(MSHFlexGrid2.TextMatrix(o, 0)) = EmpID_ Then
       ValRet = Val(MSHFlexGrid2.TextMatrix(o, 7))
       Exit For
    End If
Next
GetCapacity = ValRet
End Function

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
MSHFlexGrid1.Cols = 13

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
MSHFlexGrid1.TextMatrix(0, 10) = "Gross Prift"
MSHFlexGrid1.TextMatrix(0, 11) = "Corporate Office Expense"
MSHFlexGrid1.TextMatrix(0, 12) = "Net Prift"

MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 0
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 1500
MSHFlexGrid1.ColWidth(10) = 1500
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub CmdGetReport_Click()
GunVer = ""
Call TempDataTableDelete
Call TempDataTable
MsgBox "Importing Employee wise Data May take some time, Wait till next msg!!!"

gmTotCap = 0

tmp = "Select * from Mst_Employee " & _
      " Where EmployeeNo Not In (Select EmployeeCode From Txn_EmpCTC) " & _
      " Order By EmployeeNo"

Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   MsgBox "Some Employee data from Store-Man Missing in Employee CTC Data!!"
'   Call Grid_Head
'   Exit Sub
End If

tmp = "Select MS.StateName,ML.LocationID,ML.LocationName From Mst_Location ML " & _
      " Inner Join Mst_State MS On ML.StateID = MS.StateID " & _
      " Order By ML.LocationName "
'Call TableMst_Location(" Order By LocationName ")
Call TmpTable
Call displocationdata

For I = 1 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(I, 10) = Val(MSHFlexGrid1.TextMatrix(I, 6)) - (Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 9))) '"Gross Prift"
    MSHFlexGrid1.TextMatrix(I, 12) = Val(MSHFlexGrid1.TextMatrix(I, 10)) - Val(MSHFlexGrid1.TextMatrix(I, 11))  ' "Net Prift"
    MSHFlexGrid1.TextMatrix(I, 10) = Format(MSHFlexGrid1.TextMatrix(I, 10), "###########0.00")
    MSHFlexGrid1.TextMatrix(I, 12) = Format(MSHFlexGrid1.TextMatrix(I, 12), "###########0.00")
Next

'MsgBox GunVer


MsgBox "Done!!!"
'RsTemp.Close
End Sub
Private Sub GetRMAMforLocation(mLocationName_ As Variant, x1 As Double)
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
         MSHFlexGrid1.TextMatrix(x1, 1) = IIf(IsNull(TmpRs3.Fields(2)), "", TmpRs3.Fields(2))
      End If
   End If
   
   If IsNull(TmpRs3.Fields(3)) = False Then
      strSplitString = Split(TmpRs3.Fields(3), "~")
      EmpName_ = Trim(strSplitString(0))
      EmpNo_ = Trim(strSplitString(1))
      EmpDesig_ = Trim(strSplitString(2))
      If LCase(EmpDesig_) = "rm" Then
         MSHFlexGrid1.TextMatrix(x1, 0) = IIf(IsNull(TmpRs3.Fields(3)), "", TmpRs3.Fields(3))
      ElseIf LCase(EmpDesig_) = "am" Then
         MSHFlexGrid1.TextMatrix(x1, 1) = IIf(IsNull(TmpRs3.Fields(3)), "", TmpRs3.Fields(3))
      End If
   End If
      
   If IsNull(TmpRs3.Fields(4)) = False Then
      strSplitString = Split(TmpRs3.Fields(4), "~")
      EmpName_ = Trim(strSplitString(0))
      EmpNo_ = Trim(strSplitString(1))
      EmpDesig_ = Trim(strSplitString(2))
      If LCase(EmpDesig_) = "rm" Then
         MSHFlexGrid1.TextMatrix(x1, 0) = IIf(IsNull(TmpRs3.Fields(4)), "", TmpRs3.Fields(4))
      End If
   End If
   
   If TmpRs3.RecordCount > 1 Then
      MSHFlexGrid1.TextMatrix(x1, 0) = "*" & MSHFlexGrid1.TextMatrix(x1, 0)
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
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows '- 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(x, c + 2) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       Call GetRMAMforLocation(TmpRs!LocationName, x)
       Call CapacityData(TmpRs!LocationName, x)
       Call TotalIncomeLocWise(TmpRs!LocationName, x)
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

Private Sub CapacityData(mLocationName_ As Variant, x1 As Double)
tmp = "Select ML.LocationName,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML " & _
      " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
      " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
      " Where ML.LocationName = '" & mLocationName_ & "' And MG.CloseDate is Null " & _
      " Group By ML.LocationName "
Call Tmp1Table

'tmp = "Select ML.LocationName,Sum(MC.ClosingWeight) as TotalCapacity From Mst_Location ML " & _
'      " Left Outer Join mst_stockstatus MC On ML.LocationID = MC.LocationID " & _
'      " Where ML.LocationName = '" & mLocationName_ & "' " & _
'      " Group By ML.LocationName "
'Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x1, 5) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
End If
gmTotCap = gmTotCap + Val(MSHFlexGrid1.TextMatrix(x1, 5))
TmpRs1.Close
End Sub

Private Sub TotalIncomeLocWise(mLocationName_ As Variant, x1 As Double)

tmp = "Select ML.LocationName ,Sum(TI.AmtInclOfSerTax) As TotalIncome From Txn_IncomeDetails TI " & _
      " Right Outer Join Mst_Location ML On TI.LocationID = ML.LocationID " & _
      " Where LocationName = '" & mLocationName_ & "' " & _
      " Group By LocationName "

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x1, 6) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
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
    Call FixExpense(TmpRs1!EmployeeID, TmpRs1!TotalExpenseAmount, 8)
    TmpRs1.MoveNext
Next
'MsgBox GunVer
'GunVer = ""
End Sub
Private Sub CalDirectIncomeLocWise(mType_ As Variant, mROw_ As Variant, mEmpID_ As Variant, mEMPDesignation_ As Variant, mAllocateExpense_ As Variant, mExpenseAmount_ As Variant)
Dim mAmtAfterDeduction, mCorporateExpense, mCorporateExpenseAfterDed As Variant
Dim mEMPIDDummy_ As Double
Dim mTotCap As Double

mEMPIDDummy_ = mEmpID_
mCorporateExpense = (30000000 / mMonYr)
'mAmtAfterDeduction = mExpenseAmount_ * (Gen_Percentage / 100)


mAllocateExpense_ = IIf(IsNull(mAllocateExpense_), 0, mAllocateExpense_)

If LCase(mEMPDesignation_) = "rag" Then Exit Sub

If LCase(mEMPDesignation_) = "managment" Or LCase(mEMPDesignation_) = "rag" Then
   tmp = "Select ML.LocationName,Sum(MG.GodownCapacity) as TotalLocationCapacity From Mst_Location ML " & _
         " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
         " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
         " Where MG.CLoseDate is Null " & _
         " Group By ML.LocationName "
   mTotCap = gmTotCap
Else
   If LCase(mEMPDesignation_) = "qqs" Or LCase(mEMPDesignation_) = "drc" Then
      tmp = "Select DirectReportingTo From Mst_Employee Where EmployeeID = " & mEMPIDDummy_ & ""
      Call TmpTable
      If TmpRs.RecordCount <> 0 Then
         mEMPIDDummy_ = IIf(IsNull(TmpRs!DirectReportingTo), 0, TmpRs!DirectReportingTo)
      End If
   End If
   tmp = "Select distinct ECM.LocationID,ML.LocationName " & _
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
    mTotCap = GetCapacity(mEMPIDDummy_)
'    If mTotCap = 0 Then
'       mTotCap = 1
'    End If

End If

Call TmpTable
If TmpRs.RecordCount > 0 Then
   If LCase(mType_) = "salary" Then
      mExpenseAmount_ = mExpenseAmount_ / mMonYr
   End If
   If mAllocateExpense_ = 1 Then
      mAmtAfterDeduction = mExpenseAmount_ * (Gen_Percentage / 100)
   Else
      mAmtAfterDeduction = mExpenseAmount_
   End If
   
   If mTotCap <> 0 Then
      mCorporateExpenseAfterDed = mCorporateExpense / mTotCap 'TmpRs2!TotalLocationCapacity
      mAmtAfterDeduction = mAmtAfterDeduction / mTotCap 'TmpRs2!TotalLocationCapacity
   End If
   TotExp = mAmtAfterDeduction
   For I = 1 To TmpRs.RecordCount
       For ix = 1 To MSHFlexGrid1.Rows - 1
           If MSHFlexGrid1.TextMatrix(ix, 4) = TmpRs!LocationName Then
'              If TmpRs!Locationname = "Kolkata" Then '"Warangal" Then
''                 If LCase(mType_) = "salary" Then
'                    If mEMPIDDummy_ = 260 Or mEMPIDDummy_ = 261 Or mEMPIDDummy_ = 262 Then
'                       GunVer = GunVer & mTotCap & "," & mEMPID_ & " , " & (MSHFlexGrid1.TextMatrix(ix, 5)) & "," & mAmtAfterDeduction & "," & (Val(MSHFlexGrid1.TextMatrix(ix, mRow_)) * mAmtAfterDeduction) & Chr(13)
'
'                       tmp = tmp
'                    End If
''                 End If
'                  'GunVer = GunVer & mTotCap & "," & mEMPID_ & " , " & (MSHFlexGrid1.TextMatrix(ix, 5)) & "," & mAmtAfterDeduction & "," & (Val(MSHFlexGrid1.TextMatrix(ix, mRow_)) * mAmtAfterDeduction) & Chr(13)
'              End If

              
              If mTotCap >= Val(MSHFlexGrid1.TextMatrix(ix, 5)) Then
                 MSHFlexGrid1.TextMatrix(ix, mROw_) = Val(MSHFlexGrid1.TextMatrix(ix, mROw_)) + (Val(MSHFlexGrid1.TextMatrix(ix, 5)) * mAmtAfterDeduction)
                 MSHFlexGrid1.TextMatrix(ix, mROw_) = Round(MSHFlexGrid1.TextMatrix(ix, mROw_), 3)
                 TotExp = TotExp - (Val(MSHFlexGrid1.TextMatrix(ix, 5)) * mAmtAfterDeduction)
                 
                 rstemp.AddNew
                 GunVer = LCase(mType_) & "," & TmpRs!LocationName & "," & mTotCap & "," & mEmpID_ & " , " & (MSHFlexGrid1.TextMatrix(ix, 5)) & "," & mAmtAfterDeduction & "," & (Val(MSHFlexGrid1.TextMatrix(ix, 5)) * mAmtAfterDeduction) '& Chr(13)
                 rstemp!TempData = GunVer
                 rstemp.Update
                
                 If LCase(mEMPDesignation_) = "managment" And LCase(mType_) = "directincome" Then
                    MSHFlexGrid1.TextMatrix(ix, 11) = (Val(MSHFlexGrid1.TextMatrix(ix, 5)) * mCorporateExpenseAfterDed)
                    MSHFlexGrid1.TextMatrix(ix, 11) = Round(MSHFlexGrid1.TextMatrix(ix, 11), 3)
                    Exit For
                 End If
              Else
                 MSHFlexGrid1.TextMatrix(ix, mROw_) = Val(MSHFlexGrid1.TextMatrix(ix, mROw_)) + (TotExp)
                 MSHFlexGrid1.TextMatrix(ix, mROw_) = Round(MSHFlexGrid1.TextMatrix(ix, mROw_), 3)
                 rstemp.AddNew
                 GunVer = LCase(mType_) & "," & TmpRs!LocationName & "," & mTotCap & "," & mEmpID_ & " , " & (MSHFlexGrid1.TextMatrix(ix, 5)) & "," & mAmtAfterDeduction & "," & (TotExp) & ", **" '& Chr(13)
                 rstemp!TempData = GunVer
                 rstemp.Update
                 TotExp = 0
                 If LCase(mEMPDesignation_) = "managment" And LCase(mType_) = "directincome" Then
                    MSHFlexGrid1.TextMatrix(ix, 11) = (Val(MSHFlexGrid1.TextMatrix(ix, 5)) * mCorporateExpenseAfterDed)
                    MSHFlexGrid1.TextMatrix(ix, 11) = Round(MSHFlexGrid1.TextMatrix(ix, 11), 3)
                   Exit Sub
                 End If
              End If
           End If
       Next
       TmpRs.MoveNext
   Next
End If




'TmpRs2.Close

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
    Call FixExpense(TmpRs1!EmployeeID, TmpRs1!CTC_YEARLY / mMonYr, 9)
    TmpRs1.MoveNext
Next
'MsgBox GunVer
'GunVer = ""
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
    Call FixExpense(TmpRs1!EmployeeID, TmpRs1!MonthlyBillAmt, 10)
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

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   lblRecordCount.Caption = ""
   Exit Sub
End If

'CmdExcel.Enabled = False
mMonYr = 1
'mMonYr = 2  for 6 month costing
Gen_Percentage = 60
Call Grid_Head
Call Grid_HeadEmp

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_IncomeMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_IncomeMIS.xls")
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


Set oWS = oWB.Worksheets("Sheet2")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid2.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
        End If
    Next
Next


oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
