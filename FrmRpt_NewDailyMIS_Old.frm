VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_NewDailyMIS 
   Caption         =   "New Daily MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15000
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
         Left            =   4320
         TabIndex        =   12
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
         Left            =   7935
         TabIndex        =   9
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
         Left            =   5520
         TabIndex        =   8
         Top             =   173
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
         Left            =   6720
         TabIndex        =   7
         Top             =   173
         Width           =   1215
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   233
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optSpot 
         Caption         =   "Spot"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   233
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optNonNcdex 
         Caption         =   "Non NCDEX"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   233
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton optNCDEX 
         Caption         =   "NCDEX"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   233
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   11
         Top             =   240
         Width           =   210
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
         TabIndex        =   10
         Top             =   240
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13996
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
tmp = "update Mst_GSL set BagEquivalent = 1,WtQty=0 where flag <> 'B'  "
Call TmpTable



Call TableMst_BagEquivalent
For i = 1 To RsMst_BagEquivalent.RecordCount
    If RsMst_BagEquivalent!Equivalent <> 1 Then
       tmp = "update Mst_GSL set BagEquivalent = " & RsMst_BagEquivalent!Equivalent & " where flag <> 'B' and CommodityID = " & RsMst_BagEquivalent!CommodityID & ""
       Call TmpTable
    End If
    RsMst_BagEquivalent.MoveNext
Next


tmp = "update Mst_GSL set BagSize=BalanceWeight/BalanceBags,WtQty = BagEquivalent*BalanceWeight where flag <> 'B' and BalanceBags<>0"
Call TmpTable

tmp = "update mst_gsl set ContractID = 0 where ContractID in(select ContractID from Mst_StorageContract where bookedto < getdate())"
Call TmpTable


MsgBox "Done!!!"
CmdProcess.Enabled = False
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
lblRecordCount.Caption = ""
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

For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, c + 1).WrapText = True
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
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
For i = 2 To TmpRs.RecordCount + 1
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
    MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
    MSHFlexGrid1.TextMatrix(i, 15) = IIf(IsNull(TmpRs!MinDate), TmpRs!StartDate, TmpRs!MinDate)
    MSHFlexGrid1.TextMatrix(i, 16) = TmpRs!StartDate
    TmpRs.MoveNext
Next
lblRecordCount.Caption = TmpRs.RecordCount

'--- For  AM/RM and CC/AM

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
     " Group By ML.LocationID,ML.LocationName,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,  " & _
     " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation " & _
     " Order By ML.LocationName"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(i, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For j = 0 To TmpRs.RecordCount - 1
              Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 1, IIf(IsNull(TmpRs![AM/RM]), "", TmpRs![AM/RM]))
              Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 2, IIf(IsNull(TmpRs![CC/AM]), "", TmpRs![CC/AM]))
          Next
       End If
   Next
End If
'--- For Capacity

tmp = "select LocationID,UnitTypeID,Capacity,AccreQty from LeaseFranchCapacity Where UnitTypeID=1"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(i, 0) & ""
       If TmpRs.RecordCount > 0 Then
          Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 4, Format(Val(IIf(IsNull(TmpRs!Capacity), "0", TmpRs!Capacity)), "0"))
       Else
          Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 4, "0")
       End If
   Next
End If

'--- For Reservation Capacity

tmp = "select LocationID,sum(BookedQty) as ResCapacity from Mst_StorageContract" & _
      " where  BookedTo>= '" & Format(Date, Gen_DatFormat) & "'  group by LocationID"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(i, 0) & ""
       If TmpRs.RecordCount > 0 Then
          Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 6, Format(Val(IIf(IsNull(TmpRs!ResCapacity), "0", TmpRs!ResCapacity)), "0"))
       Else
          Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 6, "0")
       End If
   Next
End If

'--- Lease Stock
'LS 5,RS 7, FS 20
'Written by kishor and update with new query given by Mayur sir
'tmp = "select LocationID,UnitTypeID,Stock,ContractID from LeaseFranchStock Order By  LocationID,UnitTypeID,ContractID"


tmp = "SELECT  SUM(s1.BalanceWeight) AS Stock, l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID,SUM(s1.WtQty) AS WtQty"
tmp = tmp & " FROM Mst_GSL s1 INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate"
tmp = tmp & " Having (Mst_Godown.CloseDate Is Null)"
tmp = tmp & " order by l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID"


Call TmpTable
If TmpRs.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(i, 0) & ""
       If TmpRs.RecordCount > 0 Then
          For trc = 1 To TmpRs.RecordCount
              If TmpRs!UnitTypeID = 1 And (TmpRs!ContractID = 0 Or IsNull(TmpRs!ContractID) = True) Then
                 'Lease Not Reserve Stock
                 
                 Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 5, Format(Val(MSHFlexGrid1.TextMatrix(i, 5)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 20, Format(Val(MSHFlexGrid1.TextMatrix(i, 20)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              ElseIf TmpRs!UnitTypeID = 1 And IsNull(TmpRs!ContractID) = False Then
                 'Lease  Reserve Stock
                 MSHFlexGrid1.TextMatrix(i, 21) = Format(Val(MSHFlexGrid1.TextMatrix(i, 21)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0")
                 Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 7, Format(Val(MSHFlexGrid1.TextMatrix(i, 7)) + Val(IIf(IsNull(TmpRs!WtQty), "0", TmpRs!WtQty)), "0"))
                 Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 20, Format(Val(MSHFlexGrid1.TextMatrix(i, 20)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
              
              ElseIf TmpRs!UnitTypeID = 2 Then
                 'Total Franchiese Stock including Reserve stock
                 Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 11, Format(Val(MSHFlexGrid1.TextMatrix(i, 11)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 If IsNull(TmpRs!ContractID) = False Then
                    'Fanchisee reserve stock
                    Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 17, Format(Val(MSHFlexGrid1.TextMatrix(i, 17)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 End If
              ElseIf TmpRs!UnitTypeID = 3 Then
                 'Total Associate Stock including Reserve stock
                 Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 12, Format(Val(MSHFlexGrid1.TextMatrix(i, 12)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 If IsNull(TmpRs!ContractID) = False Then
                    'Associate Reserve Stock
                    Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 18, Format(Val(MSHFlexGrid1.TextMatrix(i, 18)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 End If
              ElseIf TmpRs!UnitTypeID = 4 Then
                 ' Total Stock Mgmt including Reserve stock
                 Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 13, Format(Val(MSHFlexGrid1.TextMatrix(i, 13)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 If IsNull(TmpRs!ContractID) = False Then
                    'Stock mgmt Reserve Stock
                    Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 19, Format(Val(MSHFlexGrid1.TextMatrix(i, 19)) + Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
                 End If
              End If
              'Total Other Stock

              mVal = Val(MSHFlexGrid1.TextMatrix(i, 11)) + Val(MSHFlexGrid1.TextMatrix(i, 12)) + Val(MSHFlexGrid1.TextMatrix(i, 13)) + Val(MSHFlexGrid1.TextMatrix(i, 20))
              MSHFlexGrid1.TextMatrix(i, 14) = mVal
              MSHFlexGrid1.TextMatrix(i, 8) = ""
              MSHFlexGrid1.TextMatrix(i, 8) = IIf(Val(MSHFlexGrid1.TextMatrix(i, 6)) > Val(MSHFlexGrid1.TextMatrix(i, 7)), MSHFlexGrid1.TextMatrix(i, 6), MSHFlexGrid1.TextMatrix(i, 7))
              
              TmpRs.MoveNext
          Next
    '       Else
    '          Call DispData(MSHFlexGrid1.TextMatrix(i, 0), 4, "0")
        End If
   Next
End If

tmp = "select Mst_CMP.LocationID,g.UnitTypeID,sum(a.ReservationQty) as 'ResQty'"
tmp = tmp & " From  Txn_SRGodownMapDetails a"
tmp = tmp & " inner join Txn_SRGodownMapping  b on a.SRGodownMappingID = b.SRGodownMappingID"
tmp = tmp & " inner join Mst_StorageContract c on b.ContractID = c.ContractID  and BookedTo>= getdate() "
tmp = tmp & " inner join Mst_Godown g on a.GodownID = g.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = g.CMPID"
tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
tmp = tmp & " Group by Mst_CMP.LocationID,g.UnitTypeID"
tmp = tmp & " Having g.UnitTypeID <> 1"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.Filter = ""
       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(i, 0) & ""
       If TmpRs.RecordCount > 0 Then
          Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 10, Format(Val(IIf(IsNull(TmpRs!ResQty), "0", TmpRs!ResQty)), "0"))
          mVal = (Val(IIf(MSHFlexGrid1.TextMatrix(i, 6) = "", "0", MSHFlexGrid1.TextMatrix(i, 6)) - Format(Val(IIf(IsNull(TmpRs!ResQty), "0", TmpRs!ResQty)), "0")))
          Call dispdata(MSHFlexGrid1.TextMatrix(i, 0), 6, mVal)
       End If
   Next
End If


Dim ct As Variant
ct = 0
For x = 2 To MSHFlexGrid1.Rows - 1
    ' If (MSHFlexGrid1.TextMatrix(x, 5)) <> "" And (MSHFlexGrid1.TextMatrix(x, 4)) <> "" And (MSHFlexGrid1.TextMatrix(x, 8)) = "" Then
    
    MSHFlexGrid1.TextMatrix(x, 8) = ""
    MSHFlexGrid1.TextMatrix(x, 8) = IIf(Val(MSHFlexGrid1.TextMatrix(x, 6)) > Val(MSHFlexGrid1.TextMatrix(x, 7)), MSHFlexGrid1.TextMatrix(x, 6), MSHFlexGrid1.TextMatrix(x, 7))
   
    ct = ct + 1
    MSHFlexGrid1.TextMatrix(x, 0) = ct
    If Val(MSHFlexGrid1.TextMatrix(x, 4)) <> 0 Then
       MSHFlexGrid1.TextMatrix(x, 9) = ((Val(MSHFlexGrid1.TextMatrix(x, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 8))) * 100) / Val(MSHFlexGrid1.TextMatrix(x, 4))
       MSHFlexGrid1.TextMatrix(x, 9) = Format(MSHFlexGrid1.TextMatrix(x, 9), "######0")
    End If
    
    MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(x, 4))
    MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5))
    MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(x, 6))
    MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(x, 7))
    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 8))
    
    
    MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(x, 10))
    MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(x, 11))
    MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(x, 12))
    MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(x, 13))
    MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(x, 14))
    
    MSHFlexGrid1.TextMatrix(0, 17) = Val(MSHFlexGrid1.TextMatrix(0, 17)) + Val(MSHFlexGrid1.TextMatrix(x, 17))
    MSHFlexGrid1.TextMatrix(0, 18) = Val(MSHFlexGrid1.TextMatrix(0, 18)) + Val(MSHFlexGrid1.TextMatrix(x, 18))
    MSHFlexGrid1.TextMatrix(0, 19) = Val(MSHFlexGrid1.TextMatrix(0, 19)) + Val(MSHFlexGrid1.TextMatrix(x, 19))
    MSHFlexGrid1.TextMatrix(0, 20) = Val(MSHFlexGrid1.TextMatrix(0, 20)) + Val(MSHFlexGrid1.TextMatrix(x, 20))
    MSHFlexGrid1.TextMatrix(0, 21) = Val(MSHFlexGrid1.TextMatrix(0, 21)) + Val(MSHFlexGrid1.TextMatrix(x, 21))
    
    MSHFlexGrid1.TextMatrix(0, 9) = ((Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(0, 8))) * 100) / Val(MSHFlexGrid1.TextMatrix(0, 4))
    MSHFlexGrid1.TextMatrix(0, 9) = Format(MSHFlexGrid1.TextMatrix(0, 9), "######0")
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
    .Cols = 22
    .FixedRows = 2
    .WordWrap = True
    .RowHeight(1) = 800
    .TextMatrix(1, 0) = "Sr No."  'Location ID
    .TextMatrix(1, 1) = "RM"
    .TextMatrix(1, 2) = "AM"
    .TextMatrix(1, 3) = "Location"
    .TextMatrix(1, 4) = "Leased Capacity in MT"
    .TextMatrix(1, 5) = "Leased Stock excluding reserved stock (Weighted Qty in MT)"
    .TextMatrix(1, 6) = "Reserved Capacity in MT (Leased)"
    .TextMatrix(1, 7) = "Reserved Stock (Weighted Qty in MT)"
    .TextMatrix(1, 8) = "Reserved Qty used in Capacity Utilization"
    .TextMatrix(1, 9) = "Capacity Utilisation"
    .TextMatrix(1, 10) = "Other Reserved Capacity in MT"
    .TextMatrix(1, 11) = "Total Franchisee Stock  (Qty in MT)"
    .TextMatrix(1, 12) = "Total Associated Stock  (Qty in MT)"
    .TextMatrix(1, 13) = "Total Stock Management Stock  (Qty in MT)"
    .TextMatrix(1, 14) = "Total Stock (Qty in MT)"
    .TextMatrix(1, 15) = "Entered UpTo  Date"
    .TextMatrix(1, 16) = "Location Start Date"
    .TextMatrix(1, 17) = "Frachnisee Reserved Stock"
    .TextMatrix(1, 18) = "Associate Reserved Stock"
    .TextMatrix(1, 19) = "Stock Management Reserved Stock"
    .TextMatrix(1, 20) = "Total Leased Stock (Qty in MT)"
    .TextMatrix(1, 21) = "Reserved Stock (Qty in MT)"
    
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
End With
End Sub

