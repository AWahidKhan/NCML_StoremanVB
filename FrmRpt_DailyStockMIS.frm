VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_DailyStockMIS 
   Caption         =   "Daily Stock MIS"
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
      Height          =   8295
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   5
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
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.OptionButton optNCDEX 
         Caption         =   "NCDEX"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   233
         Width           =   1215
      End
      Begin VB.OptionButton optNonNcdex 
         Caption         =   "Non NCDEX"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   233
         Width           =   1215
      End
      Begin VB.OptionButton optSpot 
         Caption         =   "Spot"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   233
         Width           =   1095
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   233
         Value           =   -1  'True
         Width           =   1095
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   173
         Width           =   1215
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
         Left            =   9360
         TabIndex        =   7
         Top             =   233
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
         Left            =   11760
         TabIndex        =   6
         Top             =   240
         Width           =   210
      End
   End
End
Attribute VB_Name = "FrmRpt_DailyStockMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer
Private Sub CmdClear_Click()
Call Grid_Head
lblRecordCount.Caption = ""
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_DailyStockMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_DailyStockMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()
Dim wc As Variant

Call Grid_Head
MsgBox "Wait till Next Message."

'tmp = "Select distinct Mst_CMP.LocationID,StateName,LocationName" & _
'      " From Mst_CMP left outer join Mst_Location  on Mst_Location.LocationID=Mst_CMP.LocationID" & _
'      " inner  join Mst_State on Mst_State.StateID=Mst_Location.StateID order by Mst_State.StateName,Mst_Location.LocationName"

tmp = "Select distinct Mst_CMP.LocationID,StateName,LocationName,min(CurrentDate)as MinDate,max(StartDate) as StartDate" & _
       " From Mst_CMP left outer join Mst_Location  on Mst_Location.LocationID=Mst_CMP.LocationID" & _
       " inner  join Mst_State on Mst_State.StateID=Mst_Location.StateID" & _
       " where CloseDate is null group by Mst_CMP.LocationID,StateName,LocationName" & _
       " order by Mst_State.StateName,Mst_Location.LocationName"

Call TmpTable
Call dispdata
MsgBox "Done !!!"
End Sub
Public Sub dispdata()

If TmpRs.RecordCount > 0 Then
    For I = 2 To TmpRs.RecordCount + 1
    
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       
 '-- For Location ID ,State , Location & MinDate
 
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs("LocationID")), "0", TmpRs("LocationID"))
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs("StateName")), "", TmpRs("StateName"))
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs("LocationName")), "", TmpRs("LocationName"))
       If OptAll.Value = True Then
          MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs("MinDate")), TmpRs("StartDate"), TmpRs("MinDate"))
       Else
          MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs("MinDate")), TmpRs("StartDate"), TmpRs("MinDate"))
       End If
 '-- For RM & AM
 
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
            " Where ML.LocationID=" & MSHFlexGrid1.TextMatrix(I, c) & "" & _
            " Group By ML.LocationID,ML.LocationName,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,  " & _
            " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation " & _
            " Order By ML.LocationName"
            
        Call Tmp0Table
        If TmpRs0.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs0("AM/RM")), "", TmpRs0("AM/RM"))
           MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs0("CC/AM")), "", TmpRs0("CC/AM"))
        End If
 '-- For Capacity, Accreditrd Qty
   
        tmp = "select LocationID,UnitTypeID,Capacity,AccreQty from LeaseFranchCapacity where LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & ""
        Call Tmp0Table
        If TmpRs0.RecordCount > 0 Then
           For k = 0 To TmpRs0.RecordCount - 1
               If TmpRs0("UnitTypeID") = 1 Then    ' For Lease
                  MSHFlexGrid1.TextMatrix(I, 5) = Format(Val(MSHFlexGrid1.TextMatrix(I, 5)) + Val(IIf(IsNull(TmpRs0("Capacity")), "0", TmpRs0("Capacity"))), "0.000")
                  MSHFlexGrid1.TextMatrix(I, 6) = Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(IIf(IsNull(TmpRs0("AccreQty")), "0", TmpRs0("AccreQty"))), "0.000")
               ElseIf TmpRs0("UnitTypeID") <> 1 Then  ' For Franchise
                  MSHFlexGrid1.TextMatrix(I, 8) = Format(Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(IIf(IsNull(TmpRs0("Capacity")), "0", TmpRs0("Capacity"))), "0.000")
                  MSHFlexGrid1.TextMatrix(I, 9) = Format(Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(IIf(IsNull(TmpRs0("AccreQty")), "0", TmpRs0("AccreQty"))), "0.000")
               End If
               TmpRs0.MoveNext
           Next
        End If
        
 '-- For Lease Stock & Franchise Stock
   
        tmp = "select LocationID,UnitTypeID,Stock from LeaseFranchStock where LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & " And ContractID IS NULL"
        Call Tmp0Table
        If TmpRs0.RecordCount > 0 Then
              For k = 0 To TmpRs0.RecordCount - 1
                If TmpRs0("UnitTypeID") = 1 Then    ' For Lease
                    MSHFlexGrid1.TextMatrix(I, 7) = Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(IIf(IsNull(TmpRs0("Stock")), "0", TmpRs0("Stock"))), "0.000")
                ElseIf TmpRs0("UnitTypeID") <> 1 Then  ' For Franchise
                    MSHFlexGrid1.TextMatrix(I, 10) = Format(Val(MSHFlexGrid1.TextMatrix(I, 10)) + Val(IIf(IsNull(TmpRs0("Stock")), "0", TmpRs0("Stock"))), "0.000")
                End If
                TmpRs0.MoveNext
              Next
        End If
          
          
  '-- For LeaseNCDEX, LeaseNonNCDEX, LeaseSpot,FranchiseNCDEX, FranchiseNonNCDEX & FrinciseSpot
  
        tmp = "select stock,UnitTypeID,ExchangeTypeID from ExchageStock where LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & " And ContractID IS NULL"
        Call Tmp0Table
        If TmpRs0.RecordCount > 0 Then
           For k = 0 To TmpRs0.RecordCount - 1
                If TmpRs0("UnitTypeID") = 1 Then ' for Lease
                    If TmpRs0("ExchangeTypeID") = 2 Then  'For NCDEX
                       If OptAll.Value = True Or optNCDEX.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                          
                       End If
                    ElseIf TmpRs0("ExchangeTypeID") = 1 Or IsNull(TmpRs0("ExchangeTypeID")) = True Then    'For NonNCDEX
                       If OptAll.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 13) = Format(Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       ElseIf optNonNcdex.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       End If
                    ElseIf TmpRs0("ExchangeTypeID") = 3 Then  'For Spot
                       If OptAll.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 14) = Format(Val(MSHFlexGrid1.TextMatrix(I, 14)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       ElseIf optSpot.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       End If
                    End If
                 ElseIf TmpRs0("UnitTypeID") <> 1 Then   ' For franchisee
                    If TmpRs0("ExchangeTypeID") = 2 Then  'For NCDEX
                       If OptAll.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 15) = Format(Val(MSHFlexGrid1.TextMatrix(I, 15)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       ElseIf optNCDEX.Value = True Then
                           MSHFlexGrid1.TextMatrix(I, 13) = Format(Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       End If
                    ElseIf TmpRs0("ExchangeTypeID") = 1 Or IsNull(TmpRs0("ExchangeTypeID")) = True Then    'For NonNCDEX
                       If OptAll.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 16) = Format(Val(MSHFlexGrid1.TextMatrix(I, 16)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       ElseIf optNonNcdex.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 13) = Format(Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       End If
                    ElseIf TmpRs0("ExchangeTypeID") = 3 Then  'For Spot
                       If OptAll.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 17) = Format(Val(MSHFlexGrid1.TextMatrix(I, 17)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       ElseIf optSpot.Value = True Then
                          MSHFlexGrid1.TextMatrix(I, 13) = Format(Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(IIf(IsNull(TmpRs0("stock")), "0", TmpRs0("stock"))), "0.000")
                       End If
                    End If
                 End If
              TmpRs0.MoveNext
           Next
                
        End If
        tmp = "select LocationID,sum(BookedQty) as ResCapacity from Mst_StorageContract" & _
              " where LocationID=" & MSHFlexGrid1.TextMatrix(I, c) & " and (BookedFrom <= '" & Format(Date, Gen_DatFormat) & "' and BookedTo>= '" & Format(Date, Gen_DatFormat) & "')  group by LocationID"
              '
        Call Tmp0Table
        If TmpRs0.RecordCount > 0 Then
           If OptAll.Value = True Then
              MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs0("ResCapacity")), "0", TmpRs0("ResCapacity"))
           Else
              MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs0("ResCapacity")), "0", TmpRs0("ResCapacity"))
           End If
        End If
        
        tmp = "select LocationID,UnitTypeID,Stock from LeaseFranchStock where LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & " and NOT (ContractID IS NULL)"
        Call Tmp0Table
        If TmpRs0.RecordCount > 0 Then
              For k = 0 To TmpRs0.RecordCount - 1
                  If OptAll.Value = True Then
                    MSHFlexGrid1.TextMatrix(I, 19) = Format(Val(MSHFlexGrid1.TextMatrix(I, 19)) + Val(IIf(IsNull(TmpRs0("Stock")), "0", TmpRs0("Stock"))), "0.000")
                  Else
                    MSHFlexGrid1.TextMatrix(I, 15) = Format(Val(MSHFlexGrid1.TextMatrix(I, 15)) + Val(IIf(IsNull(TmpRs0("Stock")), "0", TmpRs0("Stock"))), "0.000")
                  End If
                TmpRs0.MoveNext
              Next
        End If
        
        '-- For Total Stock
        tmp = "select LocationID,sum(Stock) as TotalStock from LeaseFranchStock where LocationID=" & MSHFlexGrid1.TextMatrix(I, 0) & "  group by LocationID"
        Call Tmp0Table
        If TmpRs0.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs0("TotalStock")), "0", TmpRs0("TotalStock"))
        End If
        
        For k = 5 To MSHFlexGrid1.Cols - 2
           MSHFlexGrid1.TextMatrix(0, k) = Format(Val(MSHFlexGrid1.TextMatrix(0, k)) + Val(MSHFlexGrid1.TextMatrix(I, k)), "0.000")
           MSHFlexGrid1.TextMatrix(I, k) = IIf(MSHFlexGrid1.TextMatrix(I, k) = "", "0", MSHFlexGrid1.TextMatrix(I, k))
        Next
        
        TmpRs.MoveNext
    Next
    CmdExcel.Enabled = True
Else
    MsgBox "No Record Found!!!"
End If
lblRecordCount.Refresh
lblRecordCount.Caption = TmpRs.RecordCount
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
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(1) = 500

    If OptAll.Value = True Then
        MSHFlexGrid1.Cols = 21
        MSHFlexGrid1.TextMatrix(1, 1) = " Region"
        MSHFlexGrid1.TextMatrix(1, 2) = " State"
        MSHFlexGrid1.TextMatrix(1, 3) = " Location"
        MSHFlexGrid1.TextMatrix(1, 4) = " Area Manager"
        MSHFlexGrid1.TextMatrix(1, 5) = " Lease Capacity"
        MSHFlexGrid1.TextMatrix(1, 6) = " Lease Accre Qty"
        MSHFlexGrid1.TextMatrix(1, 7) = " Lease Stock"
        MSHFlexGrid1.TextMatrix(1, 8) = " Franchise Capacity"
        MSHFlexGrid1.TextMatrix(1, 9) = " Franchise Accre Qty"
        MSHFlexGrid1.TextMatrix(1, 10) = " Franchise Stock"
        MSHFlexGrid1.TextMatrix(1, 11) = " Total Stock"
      
        MSHFlexGrid1.TextMatrix(1, 12) = " Lease NCDEX"
        MSHFlexGrid1.TextMatrix(1, 13) = " Lease NonNCDEX"
        MSHFlexGrid1.TextMatrix(1, 14) = " Lease Spot"
        MSHFlexGrid1.TextMatrix(1, 15) = " Franch NCDEX"
        MSHFlexGrid1.TextMatrix(1, 16) = " Franch NonNCDEX"
        MSHFlexGrid1.TextMatrix(1, 17) = " Franch Spot"
        MSHFlexGrid1.TextMatrix(1, 18) = " Reserved Capacity"
        MSHFlexGrid1.TextMatrix(1, 19) = " Reserved Stock"
        MSHFlexGrid1.TextMatrix(1, 20) = " Entered UpTo"
        
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 2200
        MSHFlexGrid1.ColWidth(2) = 1400
        MSHFlexGrid1.ColWidth(3) = 1400
        MSHFlexGrid1.ColWidth(4) = 2200
        MSHFlexGrid1.ColWidth(5) = 1400
        MSHFlexGrid1.ColWidth(6) = 1600
        MSHFlexGrid1.ColWidth(7) = 1600
        MSHFlexGrid1.ColWidth(8) = 1600
        MSHFlexGrid1.ColWidth(9) = 1400
        MSHFlexGrid1.ColWidth(10) = 1400
        MSHFlexGrid1.ColWidth(11) = 1600
        MSHFlexGrid1.ColWidth(12) = 1400
        MSHFlexGrid1.ColWidth(13) = 1400
        MSHFlexGrid1.ColWidth(14) = 1600
        MSHFlexGrid1.ColWidth(15) = 1400
        MSHFlexGrid1.ColWidth(16) = 1600
        MSHFlexGrid1.ColWidth(17) = 1400
        MSHFlexGrid1.ColWidth(18) = 1400
        MSHFlexGrid1.ColWidth(19) = 1200
    Else
        MSHFlexGrid1.Cols = 17
        MSHFlexGrid1.TextMatrix(1, 1) = " Region"
        MSHFlexGrid1.TextMatrix(1, 2) = " State"
        MSHFlexGrid1.TextMatrix(1, 3) = " Location"
        MSHFlexGrid1.TextMatrix(1, 4) = " Area Manager"
        MSHFlexGrid1.TextMatrix(1, 5) = " Lease Capacity"
        MSHFlexGrid1.TextMatrix(1, 6) = " Lease Accre Qty"
        MSHFlexGrid1.TextMatrix(1, 7) = " Lease Stock"
        MSHFlexGrid1.TextMatrix(1, 8) = " Franchise Capacity"
        MSHFlexGrid1.TextMatrix(1, 9) = " Franchise Accre Qty"
        MSHFlexGrid1.TextMatrix(1, 10) = " Franchise Stock"
        MSHFlexGrid1.TextMatrix(1, 11) = " Total Stock"
        MSHFlexGrid1.TextMatrix(1, 14) = " Reserved Capacity"
        MSHFlexGrid1.TextMatrix(1, 15) = " Reserved Stock"
        MSHFlexGrid1.TextMatrix(1, 16) = " Entered UpTo"
       If optNCDEX.Value = True Then
           MSHFlexGrid1.TextMatrix(1, 12) = " Lease NCDEX"
           MSHFlexGrid1.TextMatrix(1, 13) = " Franch NCDEX"
        ElseIf optNonNcdex.Value = True Then
            MSHFlexGrid1.TextMatrix(1, 12) = " Lease NonNCDEX"
            MSHFlexGrid1.TextMatrix(1, 13) = " Franch NonNCDEX"
        ElseIf optSpot.Value = True Then
            MSHFlexGrid1.TextMatrix(1, 12) = " Lease Spot"
            MSHFlexGrid1.TextMatrix(1, 13) = " Franch Spot"
        End If
               
        MSHFlexGrid1.ColWidth(0) = 0
        MSHFlexGrid1.ColWidth(1) = 2200
        MSHFlexGrid1.ColWidth(2) = 1400
        MSHFlexGrid1.ColWidth(3) = 1400
        MSHFlexGrid1.ColWidth(4) = 2200
        MSHFlexGrid1.ColWidth(5) = 1400
        MSHFlexGrid1.ColWidth(6) = 1600
        MSHFlexGrid1.ColWidth(7) = 1600
        MSHFlexGrid1.ColWidth(8) = 1600
        MSHFlexGrid1.ColWidth(9) = 1400
        MSHFlexGrid1.ColWidth(10) = 1400
        MSHFlexGrid1.ColWidth(11) = 1400
        MSHFlexGrid1.ColWidth(12) = 1400
        MSHFlexGrid1.ColWidth(13) = 1400
        MSHFlexGrid1.ColWidth(14) = 1400
        MSHFlexGrid1.ColWidth(15) = 1200
    End If
End Sub
