VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_NewDailyMIS 
   Caption         =   "New Daily MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7815
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
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
         Left            =   2640
         TabIndex        =   4
         Top             =   173
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Go"
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
         Left            =   1440
         TabIndex        =   3
         Top             =   173
         Width           =   1215
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
         Left            =   3855
         TabIndex        =   2
         Top             =   173
         Width           =   1215
      End
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Process"
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
         Left            =   240
         TabIndex        =   1
         Top             =   173
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Please ensure that you have clicked on the 'Process' button before running this report by clicking on the 'Go' button!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   12330
      End
      Begin VB.Label lblRec 
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
         TabIndex        =   6
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblRecordCount 
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
         Left            =   12240
         TabIndex        =   5
         Top             =   240
         Width           =   210
      End
   End
End
Attribute VB_Name = "FrmRpt_NewDailyMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdProcess_Click()

MsgBox "Please Wait!!!"
tmp = "update Mst_GSL set BagEquivalent = 1,WtQty=0 where Status <> 'B' and BalanceBags<>0 "
Call TmpTable
Call TableMst_BagEquivalent

For I = 1 To RsMst_BagEquivalent.RecordCount
    If RsMst_BagEquivalent!Equivalent <> 1 Then
       tmp = "update Mst_GSL set BagEquivalent = " & RsMst_BagEquivalent!Equivalent & " where Status <> 'B' and CommodityID = " & RsMst_BagEquivalent!CommodityID & " And GodownID in (Select Distinct GodownID from Mst_Godown Where UnitTypeID <> 7 And UnitTypeID <> 8 )"
       Call TmpTable
    End If
        
    'For Owned Godown (Project Locations)
    If RsMst_BagEquivalent!ProjectEquivalent <> 1 Then
       tmp = "update Mst_GSL set BagEquivalent = " & RsMst_BagEquivalent!ProjectEquivalent & " Where Status <> 'B' and CommodityID = " & RsMst_BagEquivalent!CommodityID & " And GodownID in (Select Distinct GodownID from Mst_Godown Where (UnitTypeID = 7 Or UnitTypeID = 8) )"
       Call TmpTable
    End If
        
    RsMst_BagEquivalent.MoveNext
Next

'tmp = "update Mst_GSL set BagSize = BalanceWeight/BalanceBags,WtQty = BagEquivalent*BalanceWeight where flag <> 'B' and BalanceBags<>0"

tmp = "update Mst_GSL set WtQty = BagEquivalent*BalanceWeight where Status <> 'B' and BalanceBags<>0"
Call TmpTable

tmp = "update mst_gsl set ContractID = Null where ContractID in(select ContractID from Mst_StorageContract where ExtendedPeriod < '" & Date & "')" 'GetDate())"
Call TmpTable


'If RsMst_BagEquivalent.RecordCount > 0 Then
'   RsMst_BagEquivalent.MoveFirst
'   For I = 1 To RsMst_BagEquivalent.RecordCount
''       If RsMst_BagEquivalent!Equivalent <> 1 Then
'          tmp = "Update Mst_StorageContract set WtQty = BookedQty * " & Val(RsMst_BagEquivalent!Equivalent) & " Where ExtendedPeriod>= '" & Date & "' and CommodityID = " & RsMst_BagEquivalent!CommodityID & ""
'          Call TmpTable
''       End If
'       RsMst_BagEquivalent.MoveNext
'   Next
'End If

'tmp = "Update Mst_StorageContract set WtQty = BookedQty Where ExtendedPeriod>= '" & Date & "' and CommodityID is null "
'Call TmpTable

'Commented on 22 Jan 2013 433 PM

'tmp = " Select MSCD.SCBDID,MSCD.ResQtyforMIS,MSC.CommodityID From Mst_SContractBillingDetail MSCD"
'tmp = tmp & " Inner Join Mst_StorageContract MSC On MSCD.ContractID = MSC.ContractID"
'tmp = tmp & " Where MSC.Flag = 'A' And MSC.ExtendedPeriod >= '" & Date & "' "

tmp = " Select TSRGMD.SRGodownMapDetailsID,TSGM.ContractID,MSC.ContractName,TSRGMD.SRGodownMappingID,"
tmp = tmp & " TSRGMD.ReservationQty,TSRGMD.WtQty,MSC.CommodityID,TSRGMD.GodownID from Txn_SRGodownMapDetails TSRGMD"
tmp = tmp & " Inner Join  Txn_SRGodownMapping TSGM On TSRGMD.SRGodownMappingID = TSGM.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_StorageContract MSC On TSGM.ContractID = MSC.ContractID"
tmp = tmp & " Where MSC.Flag = 'A' And MSC.ExtendedPeriod >= '" & Date & "'"

Call TmpTable

tmp = "Select GodownID from Mst_Godown Where (UnitTypeID = 7 Or UnitTypeID = 8) "
Call Tmp1Table

Call TableTxn_SRGodownMapDetails

If RsTxn_SRGodownMapDetails.RecordCount > 0 Then
   If TmpRs.RecordCount > 0 Then
      For I = 1 To TmpRs.RecordCount
          If RsMst_BagEquivalent.RecordCount > 0 Then
             RsMst_BagEquivalent.MoveFirst
             RsTxn_SRGodownMapDetails.MoveFirst
             If IsNull(TmpRs!CommodityID) = False Then
                RsMst_BagEquivalent.Find "CommodityID = " & TmpRs!CommodityID & ""
                If RsMst_BagEquivalent.EOF = False Then
                   RsTxn_SRGodownMapDetails.Find "SRGodownMapDetailsID = " & TmpRs!SRGodownMapDetailsID & ""
                   If RsTxn_SRGodownMapDetails.EOF = False Then
                      If TmpRs1.RecordCount > 0 Then
                         TmpRs1.MoveFirst
                         TmpRs1.Find "GodownID = " & TmpRs!GodownID & ""
                         If TmpRs1.EOF = False Then
                            RsTxn_SRGodownMapDetails!WtQty = TmpRs!ReservationQty * RsMst_BagEquivalent!ProjectEquivalent
                            RsTxn_SRGodownMapDetails.Update
                            RsTxn_SRGodownMapDetails.Requery
                         Else
                            RsTxn_SRGodownMapDetails!WtQty = TmpRs!ReservationQty * RsMst_BagEquivalent!Equivalent
                            RsTxn_SRGodownMapDetails.Update
                            RsTxn_SRGodownMapDetails.Requery
                         End If
                      Else
                         RsTxn_SRGodownMapDetails!WtQty = TmpRs!ReservationQty * RsMst_BagEquivalent!Equivalent
                         RsTxn_SRGodownMapDetails.Update
                         RsTxn_SRGodownMapDetails.Requery
                      End If
                   End If
                End If
             Else
                RsTxn_SRGodownMapDetails.Find "SRGodownMapDetailsID = " & TmpRs!SRGodownMapDetailsID & ""
                If RsTxn_SRGodownMapDetails.EOF = False Then
                   RsTxn_SRGodownMapDetails!WtQty = TmpRs!ReservationQty
                   RsTxn_SRGodownMapDetails.Update
                   RsTxn_SRGodownMapDetails.Requery
                End If
             End If
          End If
          TmpRs.MoveNext
      Next
   End If
End If
MsgBox "Done!!!"
CmdProcess.Enabled = False

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
LblRecordCount.Caption = ""
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NewDailyStockMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NewDailyStockMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1).WrapText = True
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()
Dim mVal As Variant
CmdExcel.Enabled = True
Call Grid_Head

'--- Location, Min Date And Start Date

tmp = "Select distinct Mst_CMP.LocationID,StateName,LocationName,min(CurrentDate)as MinDate,max(StartDate) as StartDate"
tmp = tmp & " From Mst_CMP left outer join Mst_Location  on Mst_Location.LocationID=Mst_CMP.LocationID"
tmp = tmp & " inner  join Mst_State on Mst_State.StateID=Mst_Location.StateID"
tmp = tmp & " where CloseDate is null group by Mst_CMP.LocationID,StateName,LocationName"
tmp = tmp & " order by Mst_State.StateName,Mst_Location.LocationName"
Call TmpTable
For I = 2 To TmpRs.RecordCount + 1
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
    MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
    MSHFlexGrid1.TextMatrix(I, 22) = Format(IIf(IsNull(TmpRs!MinDate), TmpRs!StartDate, TmpRs!MinDate), "dd-MMM-yyyy")
    MSHFlexGrid1.TextMatrix(I, 23) = Format(TmpRs!StartDate, "dd-MMM-yyyy")
    TmpRs.MoveNext
Next
LblRecordCount.Caption = TmpRs.RecordCount

'---For AM/RM and CC/AM

tmp = "Select ML.LocationID,ML.LocationName,EMP1.EmployeeName as 'CC/AM', " & _
     " EMP2.EmployeeName as 'AM/RM' , " & _
     " EMP3.EmployeeName as 'RM/Managment' , " & _
     " EMP4.EmployeeName as Managment " & _
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
     " Where EMP.Flag = 'Y' " & _
     " Group By ML.LocationID,ML.LocationName,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,  " & _
     " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation " & _
     " Order By ML.LocationName"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For j = 0 To TmpRs.RecordCount - 1
              Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 1, IIf(IsNull(TmpRs![AM/RM]), "", TmpRs![AM/RM]))
              Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 2, IIf(IsNull(TmpRs![CC/AM]), "", TmpRs![CC/AM]))
          Next
       End If
   Next
End If
'--- For Capacity

''comment by rakesh, treated owned godown as leased godown
'tmp = "Select LocationID,UnitTypeID,Capacity,AccreQty from LeaseFranchCapacity Where UnitTypeID=1"
tmp = "Select LocationID,UnitTypeID,Capacity,AccreQty from LeaseFranchCapacity Where UnitTypeID in (1, 7,8)"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 4, Format(Val(IIf(IsNull(TmpRs!Capacity), "0", TmpRs!Capacity)), "0"))
       Else
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 4, "0")
       End If
   Next
End If
''''''''''For Adding Capacity in Leased Capacity column and LeasedUpto Date is < = Today Date''''''''''''''''
tmp = " SELECT SUM(MG.GodownCapacity) AS Capacity, MCMP.LocationID,"
tmp = tmp & " SUM(ISNULL(MG.NCDEXAccreditedQty, 0)) AS AccreQty,MG.UnitTypeID"
tmp = tmp & " FROM Mst_Godown MG"
tmp = tmp & " INNER JOIN Mst_CMP MCMP ON MCMP.CMPID = MG.CMPID"
tmp = tmp & " INNER JOIN Mst_Location ML ON ML.LocationID = MCMP.LocationID"
tmp = tmp & " Where MG.UnitTypeID = 6 And LeasedUptoDate >= '" & Date & "' And MG.CloseDate is Null "
tmp = tmp & " GROUP BY MCMP.LocationID, MG.UnitTypeID"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 4, Val(MSHFlexGrid1.TextMatrix(I, 4)) + Format(Val(IIf(IsNull(TmpRs!Capacity), "0", TmpRs!Capacity)), "0"))
       End If
   Next
End If

'
'tmp = " SELECT SUM(s1.BalanceWeight) AS Capacity, l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, SUM(s1.WtQty) AS WtQty"
'tmp = tmp & " FROM Mst_GSL s1"
'tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
'tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z') And Mst_godown.UnitTypeID = 6 And Mst_godown.leasedUptoDate < '" & Date & "'"
'tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate"
'tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)"
'tmp = tmp & " order by l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID"
'

tmp = " SELECT SUM(s1.BalanceWeight) AS Capacity, l1.LocationID, Mst_Godown.UnitTypeID, SUM(s1.WtQty) AS WtQty"
tmp = tmp & " FROM Mst_GSL s1"
tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z') And Mst_godown.UnitTypeID = 6 And Mst_godown.leasedUptoDate < '" & Date & "'"
tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, Mst_Godown.CloseDate"
tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)"
tmp = tmp & " order by l1.LocationID, Mst_Godown.UnitTypeID"

Call TmpTable
If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          'Commented on 22 Jan 2013
          'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 4, Val(MSHFlexGrid1.TextMatrix(I, 4)) + Format(Val(IIf(IsNull(TmpRs!Capacity), "0", TmpRs!Capacity)), "0"))
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 15, Val(MSHFlexGrid1.TextMatrix(I, 15)) + Format(Val(IIf(IsNull(TmpRs!Capacity), "0", TmpRs!Capacity)), "0"))
       End If
   Next
End If
''''''''''For Adding Capacity in Leased Capacity column and LeasedUpto Date is < = Today Date''''''''''''''''

'tmp = "select MSC.LocationID,MG.UnitTypeID,sum(ReservationQty) as ResCapacity,MSC.WtQty"
'tmp = tmp & " from Mst_StorageContract MSC"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " where  MSC.BookedTo>= '" & Date & "' And MG.UnitTypeID in (1,2)"
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.WtQty Order By MSC.LocationID"

'1st Qry
'Commented on 07 Jan 2013 320PM for effect of new reservation screen
'tmp = "select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,MSC.WtQty"
'tmp = tmp & " from Mst_StorageContract MSC"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " where  MSC.BookedTo>= '" & Date & "' And MG.UnitTypeID in (1,2)"
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID,MSC.WtQty Order By MSC.LocationID"


'Commented on 22 Jan 2013 After discussion with Sankar Sir
'tmp = "Select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,MSCD.WtQty"
'tmp = tmp & " from Mst_StorageContract MSC"
'tmp = tmp & " Inner Join Mst_SContractBillingDetail MSCD On MSCD.ContractID = MSC.ContractID"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " where  MSC.Flag = 'A' And MSC.ExtendedPeriod >= '" & Date & "' And MG.UnitTypeID in (1,2) And '" & Date & "' Between  MSCD.FromDate And MSCD.ToDate And MG.CloseDate is null "
'tmp = tmp & " Group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID,MSCD.WtQty"
'tmp = tmp & " Order By MSC.LocationID"


tmp = "Select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,Sum(TSMD.WtQty) ResWtQty"
tmp = tmp & " from Mst_StorageContract MSC"
tmp = tmp & " Inner Join Mst_SContractBillingDetail MSCD On MSCD.ContractID = MSC.ContractID"
tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
''comment by rakesh, treated owned godown as leased godown
'tmp = tmp & " where  MSC.Flag = 'A' And MSC.ExtendedPeriod >= '" & Date & "' And MG.UnitTypeID in (1,2) And '" & Date & "' Between  MSCD.FromDate And MSCD.ToDate And MG.CloseDate is null "
tmp = tmp & " where  MSC.Flag = 'A' And MSC.ExtendedPeriod >= '" & Date & "' And MG.UnitTypeID in (1,2,7,8) And '" & Date & "' Between  MSCD.FromDate And MSCD.ToDate And MG.CloseDate is null "
tmp = tmp & " Group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID"
tmp = tmp & " Order By MSC.LocationID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For trc = 1 To TmpRs.RecordCount
              ''comment by rakesh, treated owned godown type as leased godown
              'If TmpRs!UnitTypeID = 1 Then
              If TmpRs!UnitTypeID = 1 Or TmpRs!UnitTypeID = 7 Or TmpRs!UnitTypeID = 8 Then
                 'Commented on 14 Dec 2012 after discussion with Mayur Sir and in Res Qty WtQty is added
                 'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
                 
                 'Commented on 22 Jan 2013 after discussion with Sankar Sir
                 'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!ResWtQty), "0", TmpRs!ResWtQty)), "0"))
                 
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 24, Format(Val(MSHFlexGrid1.TextMatrix(I, 24)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
              
              Else
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 15, Format(Val(MSHFlexGrid1.TextMatrix(I, 15)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
              End If
              TmpRs.MoveNext
          Next
       End If
   Next
End If

''''''''''For Adding Res Capacity in Leased Res Capacity column and LeasedUpto Date is < = Today Date''''''''''''''''
'tmp = " select MSC.LocationID,MG.UnitTypeID,sum(ReservationQty) as ResCapacity,MSC.WtQty from Mst_StorageContract MSC"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " Where MSC.BookedTo >= '" & Date & "' And MG.UnitTypeID = 6 And MG.leasedUptoDate >= '" & Date & "' And MG.CloseDate is null "
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.WtQty"
'tmp = tmp & " Order By MSC.LocationID"

'--2nd Qry
'Commented on 07 Jan 2013 320PM for effect of new reservation screen
'tmp = " select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,MSC.WtQty from Mst_StorageContract MSC"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " Where MSC.BookedTo >= '" & Date & "' And MG.UnitTypeID = 6 And MG.leasedUptoDate >= '" & Date & "' And MG.CloseDate is null "
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID,MSC.WtQty"
'tmp = tmp & " Order By MSC.LocationID"

'Commented on 22 Jan 2013 After discussion with Sankar Sir
'tmp = "Select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,MSCD.WtQty"
'tmp = tmp & " From Mst_StorageContract MSC"
'tmp = tmp & " Inner Join Mst_SContractBillingDetail MSCD On MSCD.ContractID = MSC.ContractID"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " Where MSC.Flag = 'A' And MSC.ExtendedPeriod >= '" & Date & "' And MG.UnitTypeID = 6 And MG.leasedUptoDate >= '" & Date & "' And MG.CloseDate is null"
'tmp = tmp & " And '" & Date & "' Between  MSCD.FromDate And MSCD.ToDate"
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID,MSCD.WtQty"
'tmp = tmp & " Order By MSC.LocationID"


tmp = "Select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,Sum(TSMD.WtQty) 'ResWtQty'"
tmp = tmp & " From Mst_StorageContract MSC"
tmp = tmp & " Inner Join Mst_SContractBillingDetail MSCD On MSCD.ContractID = MSC.ContractID"
tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
tmp = tmp & " Where MSC.Flag = 'A' And MSC.ExtendedPeriod >= '" & Date & "' And MG.UnitTypeID = 6 And MG.leasedUptoDate >= '" & Date & "' And MG.CloseDate is null"
tmp = tmp & " And '" & Date & "' Between  MSCD.FromDate And MSCD.ToDate"
tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID"
tmp = tmp & " Order By MSC.LocationID"


Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For trc = 1 To TmpRs.RecordCount
             'Commented on 14 Dec 2012 after discussion with Mayur Sir and in Res Qty WtQty is added
             'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
             
             'Commented on 22 Jan 2013 after discussion with Sankar Sir
             'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
             
             Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!ResWtQty), "0", TmpRs!ResWtQty)), "0"))
             
             Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 24, Format(Val(MSHFlexGrid1.TextMatrix(I, 24)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
             TmpRs.MoveNext
          Next
       End If
   Next
End If

'--3rd Qry
'Commented on 07 Jan 2013 320PM for effect of new reservation screen
'tmp = "Select A.LocationID,A.UnitTypeID,A.ContractID,A.WtQty,B.ResCapacity,A.Stock,B.ResWtQty"
'tmp = tmp & " from (SELECT l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID,SUM(s1.WtQty) AS WtQty,SUM(s1.BalanceWeight) AS Stock"
'tmp = tmp & " FROM Mst_StorageContract MSC"
'tmp = tmp & " inner join  Mst_GSL s1 on MSC.ContractID=s1.ContractID"
'tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
'tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z') And Mst_Godown.leasedUptoDate < '" & Date & "' And Mst_Godown.UnitTypeID = 6"
'tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate, Mst_Godown.leasedUptoDate"
'tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)) As A"
'tmp = tmp & " Inner Join"
'tmp = tmp & " (select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,MSC.WtQty as ResWtQty"
'tmp = tmp & " from Mst_StorageContract MSC"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " where  MSC.BookedTo>= '" & Date & "' And MG.leasedUptoDate < '" & Date & "' And MG.UnitTypeID = 6 "
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID, MG.leasedUptoDate,MSC.WtQty)"
'tmp = tmp & " As B On A.LocationID=B.LocationID And A.UnitTypeId=B.UnitTypeID And A.ContractID=B.ContractID"
'tmp = tmp & " Order By A.LocationID,A.UnitTypeID,A.ContractID"

'Commented on 22 Jan 2013 after discussion with Sankar Sir
'tmp = " Select A.LocationID,A.UnitTypeID,A.ContractID,A.WtQty,B.ResCapacity,A.Stock,B.ResWtQty"
'tmp = tmp & " from (SELECT l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID,SUM(s1.WtQty) AS WtQty,SUM(s1.BalanceWeight) AS Stock"
'tmp = tmp & " FROM Mst_StorageContract MSC"
'tmp = tmp & " inner join  Mst_GSL s1 on MSC.ContractID=s1.ContractID"
'tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
'tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " WHERE MSC.Flag = 'A' And (s1.Flag <> 'E') AND (s1.Flag <> 'Z') And Mst_Godown.leasedUptoDate < '" & Date & "' And Mst_Godown.UnitTypeID = 6"
'tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate, Mst_Godown.leasedUptoDate"
'tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)) As A"
'tmp = tmp & " Inner Join"
'tmp = tmp & " (select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,MSCD.WtQty as ResWtQty"
'tmp = tmp & " from Mst_StorageContract MSC"
'tmp = tmp & " Inner Join Mst_SContractBillingDetail MSCD On MSCD.ContractID = MSC.ContractID"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " where MSC.Flag = 'A' And  MSC.ExtendedPeriod>= '" & Date & "' And MG.leasedUptoDate < '" & Date & "' And MG.UnitTypeID = 6"
'tmp = tmp & " And '" & Date & "' Between  MSCD.FromDate And MSCD.ToDate"
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID, MG.leasedUptoDate,MSCD.WtQty)"
'tmp = tmp & " As B On A.LocationID=B.LocationID And A.UnitTypeId=B.UnitTypeID And A.ContractID=B.ContractID"
'tmp = tmp & " Order By A.LocationID,A.UnitTypeID,A.ContractID"



tmp = " Select A.LocationID,A.UnitTypeID,A.ContractID,A.WtQty,B.ResCapacity,A.Stock,B.ResWtQty"
tmp = tmp & " from (SELECT l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID,SUM(s1.WtQty) AS WtQty,SUM(s1.BalanceWeight) AS Stock"
tmp = tmp & " FROM Mst_StorageContract MSC"
tmp = tmp & " inner join  Mst_GSL s1 on MSC.ContractID=s1.ContractID"
tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
tmp = tmp & " WHERE MSC.Flag = 'A' And (s1.Flag <> 'E') AND (s1.Flag <> 'Z') And Mst_Godown.leasedUptoDate < '" & Date & "' And Mst_Godown.UnitTypeID = 6"
tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate, Mst_Godown.leasedUptoDate"
tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)) As A"
tmp = tmp & " Inner Join"
tmp = tmp & " (select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,Sum(TSMD.WtQty) as ResWtQty"
tmp = tmp & " from Mst_StorageContract MSC"
tmp = tmp & " Inner Join Mst_SContractBillingDetail MSCD On MSCD.ContractID = MSC.ContractID"
tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
tmp = tmp & " where MSC.Flag = 'A' And  MSC.ExtendedPeriod>= '" & Date & "' And MG.leasedUptoDate < '" & Date & "' And MG.UnitTypeID = 6"
tmp = tmp & " And '" & Date & "' Between  MSCD.FromDate And MSCD.ToDate"
tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID, MG.leasedUptoDate)"
tmp = tmp & " As B On A.LocationID=B.LocationID And A.UnitTypeId=B.UnitTypeID And A.ContractID=B.ContractID"
tmp = tmp & " Order By A.LocationID,A.UnitTypeID,A.ContractID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For trc = 1 To TmpRs.RecordCount
'             Commented on 14 Dec 2012 after discussion with Mayur Sir and in Res Qty WtQty is added
'             If TmpRs!WtQty < TmpRs!ResCapacity Then
'                Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
'             Else
'                Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
'             End If
          
             'Commented on 22 Jan 2013 after discussion with Sankar Sir
             'If TmpRs!WtQty < TmpRs!ResWtQty Then
             '   Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!ResWtQty), "0", TmpRs!ResWtQty)), "0"))
             'Else
             '   Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 6, Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
             'End If
          
          
             If TmpRs!WtQty < TmpRs!ResWtQty Then
                Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 15, Format(Val(MSHFlexGrid1.TextMatrix(I, 15)) + Val(IIf(IsNull(TmpRs!ResWtQty), "0", TmpRs!ResWtQty)), "0"))
             Else
                Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 15, Format(Val(MSHFlexGrid1.TextMatrix(I, 15)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
             End If
          
          TmpRs.MoveNext
          Next
       End If
   Next
End If
''''''''''For Adding Res Capacity in Leased Res Capacity column and LeasedUpto Date is < = Today Date''''''''''''''''

'Commented on 22 Jan 2013 after discussion with Sankar Sir
'tmp = "Select A.LocationID,A.UnitTypeID,A.ContractID,A.WtQty,B.ResCapacity,A.Stock,B.ResWtQty"
'tmp = tmp & " from (SELECT l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID,SUM(s1.WtQty) AS WtQty,SUM(s1.BalanceWeight) AS Stock"
'tmp = tmp & " FROM Mst_StorageContract MSC"
'tmp = tmp & " inner join  Mst_GSL s1 on MSC.ContractID=s1.ContractID"
'tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
'tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " WHERE MSC.Flag = 'A' And (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
'tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate"
'tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)) As A"
'tmp = tmp & " Inner Join"
'tmp = tmp & " (select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,MSC.WtQty as ResWtQty"
'tmp = tmp & " from Mst_StorageContract MSC"
'tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
'tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
'tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
'tmp = tmp & " where MSC.Flag = 'A' And  MSC.BookedTo>='" & Date & "' And MG.UnitTypeID in (1,2,6)"
'tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID,MSC.WtQty)"
'tmp = tmp & " As B On A.LocationID=B.LocationID And A.UnitTypeId=B.UnitTypeID And A.ContractID=B.ContractID"
'tmp = tmp & " Order By A.LocationID,A.UnitTypeID,A.ContractID"


tmp = "Select A.LocationID,A.UnitTypeID,A.ContractID,A.WtQty,B.ResCapacity,A.Stock,B.ResWtQty,B.LeasedUptoDate "
tmp = tmp & " from (SELECT l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID,SUM(s1.WtQty) AS WtQty,SUM(s1.BalanceWeight) AS Stock"
tmp = tmp & " FROM Mst_StorageContract MSC"
tmp = tmp & " inner join  Mst_GSL s1 on MSC.ContractID=s1.ContractID"
tmp = tmp & " INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
tmp = tmp & " WHERE MSC.Flag = 'A' And (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate"
tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)) As A"
tmp = tmp & " Inner Join"
tmp = tmp & " (select MSC.LocationID,MG.UnitTypeID,MSC.ContractID,sum(ReservationQty) as ResCapacity,Sum(TSMD.WtQty) as ResWtQty,MG.LeasedUptoDate"
tmp = tmp & " from Mst_StorageContract MSC"
tmp = tmp & " inner join Txn_SRGodownMapping TSM on MSC.ContractID=TSM.ContractID"
tmp = tmp & " inner  Join Txn_SRGodownMapDetails TSMD on TSM.SRGodownMappingID=TSMD.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_Godown MG on TSMD.GodownID=MG.GodownID"
''comment by rakesh, treated owned godown type as leased godown
'tmp = tmp & " where MSC.Flag = 'A' And  MSC.ExtendedPeriod>='" & Date & "' And MG.UnitTypeID in (1,2,6)"
tmp = tmp & " where MSC.Flag = 'A' And  MSC.ExtendedPeriod>='" & Date & "' And MG.UnitTypeID in (1,2,6,7,8) And MG.CloseDate is null "
tmp = tmp & " group by MSC.LocationID,MG.UnitTypeID,MSC.ContractID,MG.LeasedUptoDate)"
tmp = tmp & " As B On A.LocationID=B.LocationID And A.UnitTypeId=B.UnitTypeID And A.ContractID=B.ContractID"
tmp = tmp & " Order By A.LocationID,A.UnitTypeID,A.ContractID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For trc = 1 To TmpRs.RecordCount
              ''comment by rakesh, treated owned godown type as leased godown
              'If TmpRs!UnitTypeID = 1 Then 'Or (TmpRs!UnitTypeID = 6 And TmpRs!LeasedUptoDate > Date) Then  'If TmpRs!UnitTypeID = 1 Then
              If TmpRs!UnitTypeID = 1 Or TmpRs!UnitTypeID = 7 Or TmpRs!UnitTypeID = 8 Then 'Or (TmpRs!UnitTypeID = 6 And TmpRs!LeasedUptoDate > Date) Then  'If TmpRs!UnitTypeID = 1 Then
                 'Commented on 22 Jan 2013 after discussion with Mayur Sir and in Res Qty WtQty is added
                 'If TmpRs!WtQty < TmpRs!ResCapacity Then
                    'Commented on 22 Jan 2013 after discussion with Mayur Sir and in Res Qty WtQty is added
                    'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 7, Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 If TmpRs!WtQty < TmpRs!ResWtQty Then
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 7, Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 Else
                    'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 7, Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 7, Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(IIf(IsNull(TmpRs!ResWtQty), "0", TmpRs!ResWtQty)), "0"))
                 End If
                 If TmpRs!Stock < TmpRs!ResCapacity Then
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 11, Format(Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 Else
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 11, Format(Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
                 End If
              ElseIf TmpRs!UnitTypeID = 6 And TmpRs!LeasedUptoDate > Date Then
                 If TmpRs!WtQty < TmpRs!ResWtQty Then
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 7, Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 Else
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 7, Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
                    
                 End If
                 If TmpRs!Stock < TmpRs!ResCapacity Then
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 11, Format(Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 Else
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 11, Format(Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
                 End If
              Else
                 If TmpRs!Stock < TmpRs!ResCapacity Then
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 16, Format(Val(MSHFlexGrid1.TextMatrix(I, 16)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!Stock)), "0"))
                 Else
                    'Commented on 13 Dec 2012 after discussion with Mayur Sir and in Res Qty WtQty is added
                    Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 16, Format(Val(MSHFlexGrid1.TextMatrix(I, 16)) + Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
                 
                    'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 16, Format(Val(MSHFlexGrid1.TextMatrix(I, 16)) + (Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty))), "0"))
                    'Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 16, Format(Val(MSHFlexGrid1.TextMatrix(I, 16)) + Val(IIf(IsNull(TmpRs!ResWtQty), "0", TmpRs!ResWtQty)), "0"))
                 End If
              End If
              TmpRs.MoveNext
          Next
       End If
   Next
End If

tmp = "SELECT  SUM(s1.BalanceWeight) AS Stock, l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID,SUM(s1.WtQty) AS WtQty,Mst_Godown.LeasedUptoDate"
tmp = tmp & " FROM Mst_GSL s1 INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate,Mst_Godown.LeasedUptoDate"
tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)"
tmp = tmp & " order by l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID"

Call TmpTable
If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For trc = 1 To TmpRs.RecordCount
              ''comment by rakesh, treated owned godown type as leased godown
              'If TmpRs!UnitTypeID = 1 Then 'Or TmpRs!UnitTypeID = 6 Then ' If TmpRs!UnitTypeID = 1 Then
              If TmpRs!UnitTypeID = 1 Or TmpRs!UnitTypeID = 7 Or TmpRs!UnitTypeID = 8 Then 'Or TmpRs!UnitTypeID = 6 Then ' If TmpRs!UnitTypeID = 1 Then
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 5, Format(Val(MSHFlexGrid1.TextMatrix(I, 5)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 10, Format(Val(MSHFlexGrid1.TextMatrix(I, 10)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              ElseIf TmpRs!UnitTypeID = 6 And TmpRs!LeasedUptoDate > Date Then
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 5, Format(Val(MSHFlexGrid1.TextMatrix(I, 5)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 10, Format(Val(MSHFlexGrid1.TextMatrix(I, 10)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              ElseIf TmpRs!UnitTypeID = 2 Then
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 14, Format(Val(MSHFlexGrid1.TextMatrix(I, 14)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              ElseIf TmpRs!UnitTypeID = 6 And TmpRs!LeasedUptoDate < Date Then
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 14, Format(Val(MSHFlexGrid1.TextMatrix(I, 14)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              ElseIf TmpRs!UnitTypeID = 3 Then
                 'Total Associate Stock including Reserve stock
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 19, Format(Val(MSHFlexGrid1.TextMatrix(I, 19)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              ElseIf TmpRs!UnitTypeID = 4 Then
                 ' Total Stock Mgmt including Reserve stock
                 Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 20, Format(Val(MSHFlexGrid1.TextMatrix(I, 20)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              End If
              'Total Other Stock
              TmpRs.MoveNext
          Next
        End If
   Next
End If
Dim mCnt As Double
mCnt = 1

For x = 2 To MSHFlexGrid1.Rows - 1
    
    MSHFlexGrid1.TextMatrix(x, 8) = Val(MSHFlexGrid1.TextMatrix(x, 5)) - Val(MSHFlexGrid1.TextMatrix(x, 7))
    MSHFlexGrid1.TextMatrix(x, 9) = Format((Val(MSHFlexGrid1.TextMatrix(x, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 6))) / IIf(Val(MSHFlexGrid1.TextMatrix(x, 4)) = 0, 1, Val(MSHFlexGrid1.TextMatrix(x, 4))), "#####0.00") * 100
    MSHFlexGrid1.TextMatrix(x, 12) = Val(MSHFlexGrid1.TextMatrix(x, 10)) - Val(MSHFlexGrid1.TextMatrix(x, 11))
    
    'Commented on 22 Jan 2013 after discussion with Sankar Sir
    'MSHFlexGrid1.TextMatrix(x, 13) = Val(MSHFlexGrid1.TextMatrix(x, 12)) + Val(MSHFlexGrid1.TextMatrix(x, 6))
    
    MSHFlexGrid1.TextMatrix(x, 13) = Val(MSHFlexGrid1.TextMatrix(x, 12)) + Val(MSHFlexGrid1.TextMatrix(x, 24))
    
    MSHFlexGrid1.TextMatrix(x, 17) = Val(MSHFlexGrid1.TextMatrix(x, 14)) - Val(MSHFlexGrid1.TextMatrix(x, 16))
    MSHFlexGrid1.TextMatrix(x, 18) = Val(MSHFlexGrid1.TextMatrix(x, 17)) + Val(MSHFlexGrid1.TextMatrix(x, 15))
    
    MSHFlexGrid1.TextMatrix(x, 21) = Val(MSHFlexGrid1.TextMatrix(x, 13)) + Val(MSHFlexGrid1.TextMatrix(x, 18)) + Val(MSHFlexGrid1.TextMatrix(x, 19))

    MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(x, 4))
    MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5))
    MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(x, 6))
    MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(x, 7))
    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 8))
    
    MSHFlexGrid1.TextMatrix(0, 9) = Format((Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(0, 6))) / IIf(Val(MSHFlexGrid1.TextMatrix(0, 4)) = 0, 1, Val(MSHFlexGrid1.TextMatrix(0, 4))), "#####0.00") * 100

    MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(x, 10))
    MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(x, 11))
    MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(x, 12))
    MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(x, 13))
    MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(x, 14))
    MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(x, 15))
    MSHFlexGrid1.TextMatrix(0, 16) = Val(MSHFlexGrid1.TextMatrix(0, 16)) + Val(MSHFlexGrid1.TextMatrix(x, 16))

    MSHFlexGrid1.TextMatrix(0, 17) = Val(MSHFlexGrid1.TextMatrix(0, 17)) + Val(MSHFlexGrid1.TextMatrix(x, 17))
    MSHFlexGrid1.TextMatrix(0, 18) = Val(MSHFlexGrid1.TextMatrix(0, 18)) + Val(MSHFlexGrid1.TextMatrix(x, 18))
    MSHFlexGrid1.TextMatrix(0, 19) = Val(MSHFlexGrid1.TextMatrix(0, 19)) + Val(MSHFlexGrid1.TextMatrix(x, 19))
    MSHFlexGrid1.TextMatrix(0, 20) = Val(MSHFlexGrid1.TextMatrix(0, 20)) + Val(MSHFlexGrid1.TextMatrix(x, 20))
    MSHFlexGrid1.TextMatrix(0, 21) = Val(MSHFlexGrid1.TextMatrix(0, 21)) + Val(MSHFlexGrid1.TextMatrix(x, 21))
    MSHFlexGrid1.TextMatrix(0, 24) = Val(MSHFlexGrid1.TextMatrix(0, 24)) + Val(MSHFlexGrid1.TextMatrix(x, 24))
    MSHFlexGrid1.TextMatrix(x, 0) = mCnt
    mCnt = mCnt + 1
Next
tmp = ""

MsgBox " Done!!!"

End Sub

Private Sub dispdata(L_ID As Variant, mCol As Variant, Rec As Variant)
For k = 2 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(k, 0) = L_ID Then
       MSHFlexGrid1.TextMatrix(k, mCol) = Rec
       Exit For
    End If
Next
End Sub
Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
CmdExcel.Enabled = False
Call Grid_Head

End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .Cols = 25
    .FixedRows = 2
    .WordWrap = True
    .RowHeight(1) = 800
    .TextMatrix(1, 0) = "Sr No."  'Location ID
    .TextMatrix(1, 1) = "RM"
    .TextMatrix(1, 2) = "AM"
    .TextMatrix(1, 3) = "Location"
    .TextMatrix(1, 4) = "Leased Capacity in MT"
    .TextMatrix(1, 5) = "Leased Stock (W-Qty)"
    .TextMatrix(1, 6) = "Leased Reserved Capacity (W-Qty)" '"Leased Reserved Capacity (Qty)"
    .TextMatrix(1, 7) = "Leased Reserved Stock (Actual Stock not exceeding Reserved capacity for each reserved godown (W-Qty)" '"Leased Reserved Stock (Actual Stock not exceeding Reserved capacity for each reservation id(W-Qty)"
    .TextMatrix(1, 8) = "Non Reserved Leased Stock (W-Qty)"
    .TextMatrix(1, 9) = "Capacity Utilisation "
    .TextMatrix(1, 10) = "Leased Stock (Qty in MT)"
    .TextMatrix(1, 11) = "Leased Reservation Stock (Actual stock not exceeding reserved capacity for each reserved godown (Qty))" '"Leased Reservation Stock (Actual stock not exceeding reserved capacity for each reservation id (Qty))"
    .TextMatrix(1, 12) = "Non Reserved Leased Stock  (Qty)"
    .TextMatrix(1, 13) = "Billable Leased Stock(Qty)"
    .TextMatrix(1, 14) = "Franchisee Stock (Qty)"
    .TextMatrix(1, 15) = "Franchisee Reserved Capacity (Qty)"
    .TextMatrix(1, 16) = "Franchisee Reserved Stock (Actual stock not exceeding reserved capacity for each reserved godown(Qty))" '"Franchisee Reserved Stock (Actual stock not exceeding reserved capacity for each reservation id (Qty))"
    .TextMatrix(1, 17) = "Non Reserved Frachnisee Stock (Qty)"
    .TextMatrix(1, 18) = "Billable Franchisee Stock(Qty)"
    .TextMatrix(1, 19) = "Total Stock in Associate Locations (Qty)"
    .TextMatrix(1, 20) = "Total Stock in Stock Monitoring Locations (Qty)"
    .TextMatrix(1, 21) = "Total Billable Qty(Excluding Stock Management Qty)"
    .TextMatrix(1, 22) = "Entered UpTo Date"
    .TextMatrix(1, 23) = "Location Start Date"
    .TextMatrix(1, 24) = "Leased Reserved Capacity (Qty)"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 2200
    .ColWidth(2) = 2200
    .ColWidth(3) = 1600
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
    .ColWidth(7) = 1600
    .ColWidth(8) = 1600
    .ColWidth(9) = 1600
    .ColWidth(10) = 1400
    .ColWidth(11) = 1400
    .ColWidth(12) = 1400
    .ColWidth(13) = 1400
    .ColWidth(14) = 1800
    .ColWidth(15) = 1800
    .ColWidth(16) = 1800
    .ColWidth(17) = 1800
    .ColWidth(18) = 1800
    .ColWidth(19) = 1800
    .ColWidth(20) = 1800
    .ColWidth(21) = 1800
    .ColWidth(22) = 1400
    .ColWidth(23) = 1400
End With

End Sub


