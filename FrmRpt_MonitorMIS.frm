VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_MonitorMIS 
   Caption         =   "Monitor MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   6
      Top             =   975
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   7
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
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   135
      TabIndex        =   5
      Top             =   0
      Width           =   15015
      Begin VB.TextBox TxtEmpName 
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
         Left            =   9120
         TabIndex        =   11
         Top             =   480
         Width           =   3105
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2325
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
         Left            =   4380
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   2280
      End
      Begin VB.ComboBox CmbReportType 
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
         ItemData        =   "FrmRpt_MonitorMIS.frx":0000
         Left            =   120
         List            =   "FrmRpt_MonitorMIS.frx":0002
         TabIndex        =   0
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
         Height          =   375
         Left            =   13560
         TabIndex        =   4
         Top             =   480
         Width           =   1095
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
         Left            =   12360
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Total Record Found :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12180
         TabIndex        =   14
         Top             =   105
         Width           =   2295
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   14550
         TabIndex        =   13
         Top             =   120
         Width           =   90
      End
      Begin VB.Label LblEmpName 
         Caption         =   "Employee Name"
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
         Left            =   9960
         TabIndex        =   12
         Top             =   150
         Width           =   1815
      End
      Begin VB.Label LblCMPName 
         Caption         =   "CMP"
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
         Left            =   7560
         TabIndex        =   10
         Top             =   150
         Width           =   600
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
         Left            =   5160
         TabIndex        =   9
         Top             =   150
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Report Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   150
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmRpt_MonitorMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mReportType As Variant
Dim mLocationID, mCMPID As Variant

Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location("Where " & Gen_WcLocation)
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
Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP & " And LocationID= " & mLocationID
Else
   tmp1 = "And LocationID = " & mLocationID
End If
Call TableMst_CMP(" where CloseDate is null " & tmp1)

CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
'----
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 7
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.TextMatrix(0, 0) = "CMP ID"
MSHFlexGrid1.TextMatrix(0, 1) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Alias"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "State"
MSHFlexGrid1.TextMatrix(0, 5) = "Start Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Close Date"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 2500
MSHFlexGrid1.ColWidth(3) = 2500
MSHFlexGrid1.ColWidth(4) = 2500
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1500

End Sub

Private Sub Grid_Head2()

End Sub
Private Sub Grid_Head3()

End Sub

Private Sub Grid_Head4()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 6
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.TextMatrix(0, 0) = "CMP ID"
MSHFlexGrid1.TextMatrix(0, 1) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Designation"
MSHFlexGrid1.TextMatrix(0, 5) = "Status"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 2500
MSHFlexGrid1.ColWidth(3) = 2500
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1200
End Sub

Private Sub Grid_Head()


If mReportType = 1 Then
   Call Grid_Head1
End If
If mReportType = 2 Then

End If

If mReportType = 3 Then

End If

If mReportType = 4 Then
   Call Grid_Head4
End If




End Sub

Private Sub CmbReportType_GotFocus()
CmbReportType.Clear
CmbReportType.AddItem "1. CMP not connected to any CMPM"
CmbReportType.AddItem "2. Employee not connected to any location"
CmbReportType.AddItem "3. Location connected to more than one AM"
CmbReportType.AddItem "4. CMP has more than one CMPM"

End Sub


Private Sub CmbReportType_LostFocus()

tmp = CmbReportType.Text
If tmp = "1. CMP not connected to any CMPM" Or _
   tmp = "2. Employee not connected to any location" Or _
   tmp = "3. Location connected to more than one AM" Or _
   tmp = "4. CMP has more than one CMPM" Then
   mReportType = Mid(tmp, 1, 1)
   If mReportType = 1 Then
      TxtEmpName.Visible = False
      LblEmpName.Visible = False
   ElseIf mReportType = 4 Then
      TxtEmpName.Visible = True
      LblEmpName.Visible = True
   End If
   Exit Sub
End If

MsgBox "Invalid Selection!!!"
CmbReportType.SetFocus
End Sub

Private Sub CmdGetReport_Click()
Dim Wc As String
Dim tmp1 As String
Wc = ""
If CmbReportType.Text = "" Then
   MsgBox "Please select report type"
   CmbReportType.SetFocus
   Call Grid_Head1
   Exit Sub
End If

If CmbLocationID.Text <> "" Then
   Wc = GenWc(Wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      Wc = GenWc(Wc, "", tmp1, "N", "I")
   End If
End If

If CmbCMPName.Text <> "" Then
   Wc = GenWc(Wc, mCMPID, "MC.CMPID", "N", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp1 = Replace(Gen_WcCMP, "CMPID", "MC.CMPID")
      Wc = GenWc(Wc, "", tmp1, "N", "I")
   End If
End If


If mReportType = 1 Then
   TxtEmpName.Text = ""
   tmp = "Select MC.CMPID,MC.CMPName,MC.CMPAlias,ML.LocationName,MS.StateName,MC.StartDate,MC.CloseDate"
   tmp = tmp & " From Mst_CMP MC"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
   tmp = tmp & " Where MC.CMPID Not In (Select CMPID From Mst_EMPCMPMapping) And MC.CloseDate is NUll Order By MC.CMPName"
   
   Call TmpTable
   Call Grid_Head1
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount + 1
      For I = 1 To TmpRs.RecordCount
          For c = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
          Next
'          MSHFlexGrid1.TextMatrix(i, 5) = IIf(MSHFlexGrid1.TextMatrix(i, 5) = "Y", "Active", "De-Active")
          TmpRs.MoveNext
      Next
   End If
End If

If mReportType = 2 Then
   MsgBox "Not working!!!"
   Exit Sub
End If

If mReportType = 3 Then
   MsgBox "Not working!!!"
   Exit Sub
End If

If mReportType = 4 Then
   If TxtEmpName.Text <> "" Then
      Wc = GenWc(Wc, TxtEmpName.Text, "EMP.EmployeeName", "S", "L")
   End If
   
   tmp = "Select MECM.CMPID,MC.CMPName,ML.LocationName,EMP.EmployeeName + '~' + EMP.EmployeeNo As CMPM, "
   tmp = tmp & " MD.Designation ,EMP.Flag As Status"
   tmp = tmp & " From Mst_EMPCMPMapping MECM"
   tmp = tmp & " Inner Join Mst_CMP MC On MECM.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Employee EMP On MECm.EmployeeID = EMP.EmployeeID"
   tmp = tmp & " Inner Join Mst_Designation MD On EMP.DesignationID = MD.DesignationID "
   tmp = tmp & Wc
   tmp = tmp & " And MECM.CMPID In"
   tmp = tmp & " (Select a.CMPID From Mst_EMPCMPMapping a Inner Join Mst_Employee b On a.EmployeeID = b.EmployeeID And b.Flag = 'Y' "
   tmp = tmp & " Group By CMPID Having Count(a.EmployeeID) > 1 ) "
   tmp = tmp & " And EMP.Flag = 'Y' Order By MC.CMPName"
   Call TmpTable
   Call Grid_Head4
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount + 1
      For I = 1 To TmpRs.RecordCount
          'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount + 1
          For c = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(I, c) = TmpRs.Fields(c)
          Next
          MSHFlexGrid1.TextMatrix(I, 5) = IIf(MSHFlexGrid1.TextMatrix(I, 5) = "Y", "Active", "De-Active")
          TmpRs.MoveNext
      Next
   End If
End If

CmdExcel.Enabled = True
MsgBox "Done!!!"

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

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x1, 5) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
End If

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
    TmpRs1.MoveNext
Next

End Sub
Private Sub CalDirectIncomeLocWise(mType_ As Variant, mrow_ As Variant, mEmpID_ As Variant, mEMPDesignation_ As Variant, mAllocateExpense_ As Variant, mExpenseAmount_ As Variant)
Dim mAmtAfterDeduction, mCorporateExpense, mCorporateExpenseAfterDed, mEMPIDDummy_ As Variant

mEMPIDDummy_ = mEmpID_
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
   For I = 1 To TmpRs.RecordCount
       If mAllocateExpense_ = 1 Then
          mAmtAfterDeduction = mExpenseAmount_ * (Gen_Percentage / 100)
       Else
          mAmtAfterDeduction = mExpenseAmount_
       End If
       mCorporateExpenseAfterDed = mCorporateExpense / TmpRs2!TotalLocationCapacity
       mAmtAfterDeduction = mAmtAfterDeduction / TmpRs2!TotalLocationCapacity
       For ix = 1 To MSHFlexGrid1.Rows - 1
           If MSHFlexGrid1.TextMatrix(ix, 4) = TmpRs!LocationName Then
              MSHFlexGrid1.TextMatrix(ix, mrow_) = Val(MSHFlexGrid1.TextMatrix(ix, mrow_)) + (IIf(IsNull(TmpRs!TotalLocationCapacity), 0, TmpRs!TotalLocationCapacity) * mAmtAfterDeduction)
              MSHFlexGrid1.TextMatrix(ix, mrow_) = Round(MSHFlexGrid1.TextMatrix(ix, mrow_), 3)
              If LCase(mEMPDesignation_) = "managment" And LCase(mType_) = "directincome" Then
                 MSHFlexGrid1.TextMatrix(ix, 10) = (IIf(IsNull(TmpRs!TotalLocationCapacity), 0, TmpRs!TotalLocationCapacity) * mCorporateExpenseAfterDed)
                 MSHFlexGrid1.TextMatrix(ix, 10) = Round(MSHFlexGrid1.TextMatrix(ix, 10), 3)
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

Frame1.Enabled = False
Frame2.Enabled = False

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   
'   CmdGetReport.Enabled = False
'   CmdExcel.Enabled = False
'   CmdClear.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Call Grid_Head
CmdExcel.Enabled = False


CmdExcel.Enabled = False
Gen_Percentage = 60
Call Grid_Head

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
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


