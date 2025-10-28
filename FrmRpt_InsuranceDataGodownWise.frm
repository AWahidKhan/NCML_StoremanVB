VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_InsuranceDataGodownWise 
   Caption         =   "Monthly Insurance Data Godown wise"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   15015
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
         Left            =   11925
         TabIndex        =   4
         Top             =   255
         Width           =   1455
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
         Left            =   10470
         TabIndex        =   3
         Top             =   255
         Width           =   1455
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
         Left            =   13380
         TabIndex        =   5
         Top             =   255
         Width           =   1455
      End
      Begin VB.ComboBox CmbMonth 
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
         ItemData        =   "FrmRpt_InsuranceDataGodownWise.frx":0000
         Left            =   855
         List            =   "FrmRpt_InsuranceDataGodownWise.frx":0002
         TabIndex        =   1
         Top             =   930
         Width           =   1680
      End
      Begin VB.ComboBox CmbYear 
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
         ItemData        =   "FrmRpt_InsuranceDataGodownWise.frx":0004
         Left            =   3330
         List            =   "FrmRpt_InsuranceDataGodownWise.frx":0006
         TabIndex        =   2
         Top             =   930
         Width           =   1560
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   345
         Left            =   1830
         TabIndex        =   15
         Top             =   255
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
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
         Format          =   111607809
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Date for Report"
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
         Left            =   240
         TabIndex        =   16
         Top             =   307
         Width           =   1365
      End
      Begin VB.Label Label2 
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
         Left            =   6375
         TabIndex        =   14
         Top             =   315
         Width           =   90
      End
      Begin VB.Label Label1 
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
         Left            =   3975
         TabIndex        =   13
         Top             =   315
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   195
         TabIndex        =   12
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         Left            =   2790
         TabIndex        =   11
         Top             =   975
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   15030
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   7560
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   7560
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4215
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7435
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5106
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_InsuranceDataGodownWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTotDays As Variant
Dim mDate As Variant
Dim mYear As Variant
Dim mMonthNo As Variant
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11

MSHFlexGrid1.TextMatrix(0, 0) = "State"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 5) = "Stock Bags"
MSHFlexGrid1.TextMatrix(0, 6) = "Stock Weight"
MSHFlexGrid1.TextMatrix(0, 7) = "Date"
MSHFlexGrid1.TextMatrix(0, 8) = "Price"
MSHFlexGrid1.TextMatrix(0, 9) = "Amount"
MSHFlexGrid1.TextMatrix(0, 10) = "Gain / Loss"

MSHFlexGrid1.ColWidth(0) = 2200
MSHFlexGrid1.ColWidth(1) = 2200
MSHFlexGrid1.ColWidth(2) = 2200
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400

End Sub
Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 13

MSHFlexGrid3.TextMatrix(0, 0) = "State"
MSHFlexGrid3.TextMatrix(0, 1) = "Location"
MSHFlexGrid3.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid3.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid3.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid3.TextMatrix(0, 5) = "Stock Bags"
MSHFlexGrid3.TextMatrix(0, 6) = "Stock Weight"
MSHFlexGrid3.TextMatrix(0, 7) = "Date"
MSHFlexGrid3.TextMatrix(0, 8) = "DepBags"
MSHFlexGrid3.TextMatrix(0, 9) = "DepWt"
MSHFlexGrid3.TextMatrix(0, 10) = "WitBags"
MSHFlexGrid3.TextMatrix(0, 11) = "WitWt"
MSHFlexGrid3.TextMatrix(0, 12) = "Weight Gain/Loss"

MSHFlexGrid3.ColWidth(0) = 2200
MSHFlexGrid3.ColWidth(1) = 2200
MSHFlexGrid3.ColWidth(2) = 2200
MSHFlexGrid3.ColWidth(3) = 2200
MSHFlexGrid3.ColWidth(4) = 1400
MSHFlexGrid3.ColWidth(5) = 1400
MSHFlexGrid3.ColWidth(6) = 1400

End Sub
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 11

MSHFlexGrid2.TextMatrix(0, 0) = "State"
MSHFlexGrid2.TextMatrix(0, 1) = "Location"
MSHFlexGrid2.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid2.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid2.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid2.TextMatrix(0, 5) = "Balance Bags"
MSHFlexGrid2.TextMatrix(0, 6) = "Balance Weight"
MSHFlexGrid2.TextMatrix(0, 7) = "Average Bags"
MSHFlexGrid2.TextMatrix(0, 8) = "Average Weight"
MSHFlexGrid2.TextMatrix(0, 9) = "Total Amount"
MSHFlexGrid2.TextMatrix(0, 10) = "Average Amount"

MSHFlexGrid2.ColWidth(0) = 2200
MSHFlexGrid2.ColWidth(1) = 2200
MSHFlexGrid2.ColWidth(2) = 2200
MSHFlexGrid2.ColWidth(3) = 2200
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 1400
MSHFlexGrid2.ColWidth(6) = 1400
MSHFlexGrid2.ColWidth(7) = 1400
MSHFlexGrid2.ColWidth(8) = 1400
MSHFlexGrid2.ColWidth(9) = 1400

End Sub

Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I)
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then
   Exit Sub
End If
End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2009 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If CmbYear.Text = "" Then
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Call Grid_Head1
Call Grid_Head2
Label2.Caption = ""
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_AvgMonthStockMIS_" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_AvgMonthStockMIS_" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next


Set oWS = oWB.Worksheets("Sheet2")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
        End If
    Next
Next


oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub CmdGetReport_Click()
Dim mDayNo As Variant

'If CmbMonth.Text = "" Then
'   MsgBox "Please Select Month."
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox "Please Select Year."
'   Exit Sub
'End If

If TxtFromDate.Value > Date Then
   MsgBox "Future Date not allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If


Call Grid_Head
Call Grid_Head1


MsgBox "Wait till next message !!!"
Label2.Caption = ""
mYear = Year(TxtFromDate.Value) 'CmbYear.Text
mMonthNo = Month(TxtFromDate.Value) 'Month(mYear & "-" & CmbMonth.Text & "-" & 1)
mTotDays = Day(TxtFromDate.Value) '  Day(DateSerial(mYear, mMonthNo + 1, 0))

'tmp = "Select ML.LocationName,MCm.Commodity"
'tmp = tmp & " From Txn_CxTF_Details TC"
'tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.CxTFDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity"
'tmp = tmp & " order by ML.LocationName,MCm.Commodity"

'Commented on 11 Oct 2012 After mail from Sandeep Mohite regarding -ve Balance Bags
'tmp = "Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
'tmp = tmp & " From Txn_CxTF_Details TC"
'tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.CxTFDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'tmp = tmp & " order by MS.StateName,ML.LocationName,MCm.Commodity"
''



'tmp = " Select StateName,LocationName,EQZone,Commodity From"
'tmp = tmp & " (("
'tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
'tmp = tmp & " From Txn_CxTF_Details TC"
'tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.CxTFDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'tmp = tmp & " )"
''==========================added transaction details on 19 apr 2013 ==========================
'tmp = tmp & " Union All"
'tmp = tmp & " ("
'tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
'tmp = tmp & " From Txn_AdjTransaction TAD"
'tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TAD.TxnDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'tmp = tmp & " )"
''====================================================
'tmp = tmp & " Union All"
'tmp = tmp & " ("
'tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
'tmp = tmp & " From Txn_Spillage TC"
'tmp = tmp & " inner Join Mst_GSL MG On TC.Mst_GSL_ID = MG.Mst_GSL_ID And TC.SM_Prefix = MG.SM_Prefix And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
'tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.SpillageDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'tmp = tmp & " )"
'tmp = tmp & " Union All"
'tmp = tmp & " ("
'tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
'tmp = tmp & " From Txn_Spillage TC"
'tmp = tmp & " inner Join Mst_GSL MG On TC.NewMst_GSL_ID = MG.Mst_GSL_ID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
'tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.SpillageDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'tmp = tmp & " )) TempTable"
'tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity"


tmp = " Select StateName,LocationName,EQZone,Commodity ,GodownNo"
tmp = tmp & " From (("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,MGod.GodownNo"
tmp = tmp & " From Txn_CxTF_Details TC"
tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_GSL MGSL on TG.Mst_GSL_ID = MGSL.Mst_GSL_ID ANd TG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " inner Join Mst_Godown MGod on MGSL.GodownID = MGod.GodownID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.CxTFDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "' " '" & mYear & "-" & mMonthNo & "-" & mTotDays & "'"
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )"
tmp = tmp & " Union All ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity ,MGod.GodownNo"
tmp = tmp & " From Txn_AdjTransaction TAD"
tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
tmp = tmp & " inner Join Txn_GSLAdjTransaction TGAD on TAD.TxnID = TGAD.TxnID"
tmp = tmp & " inner Join Mst_GSL MGSL on TGAD.Mst_GSL_ID = MGSL.Mst_GSL_ID And TAD.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " inner Join Mst_Godown MGod on MGSL.GodownID = MGod.GodownID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
tmp = tmp & " where TAD.TxnDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )"
tmp = tmp & " Union All ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity ,MGod.GodownNo"
tmp = tmp & " From Txn_Spillage TC"
tmp = tmp & " inner Join Mst_GSL MG On TC.Mst_GSL_ID = MG.Mst_GSL_ID And TC.SM_Prefix = MG.SM_Prefix And TC.Flag = 'S'"
tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner Join Mst_Godown MGod on MG.GodownID = MGod.GodownID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.SpillageDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )"
tmp = tmp & " Union All ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity ,MGod.GodownNo"
tmp = tmp & " From Txn_Spillage TC"
tmp = tmp & " inner Join Mst_GSL MG On TC.NewMst_GSL_ID = MG.Mst_GSL_ID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner Join Mst_Godown MGod on MG.GodownID = MGod.GodownID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.SpillageDate<='" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )) TempTable"
tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity,GodownNo"

Call TmpTable

MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 1
For N = 1 To TmpRs.RecordCount
    MSHFlexGrid2.TextMatrix(N, 0) = TmpRs!StateName
    MSHFlexGrid2.TextMatrix(N, 1) = TmpRs!Locationname
    MSHFlexGrid2.TextMatrix(N, 2) = IIf(IsNull(TmpRs!EQZone), "", TmpRs!EQZone)
    MSHFlexGrid2.TextMatrix(N, 3) = TmpRs!Commodity
    MSHFlexGrid2.TextMatrix(N, 4) = TmpRs!GodownNo
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
TmpRs.MoveFirst
    
ProgressBar1.Value = 0
ProgressBar1.Max = mTotDays + 1


For r1 = mTotDays To mTotDays
    
    Call Grid_Head2
    'r1 = 30
    
    mDate = Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-" & r1

'    tmp = "Select ML.LocationName,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' "
'    tmp = tmp & " group by ML.LocationName,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by ML.LocationName,MCm.Commodity,TC.TxnType"
    
    
'Commented on 11 Oct 2012 After mail from Sandeep Mohite regarding -ve Balance Bags
'    tmp = "Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'"
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by MS.StateName,ML.LocationName,MCm.Commodity,TC.TxnType"
    
    

'    tmp = " Select StateName,LocationName,EQZone,Commodity,TxnType,Sum(NoOfBags)'NoOfBags',Sum(GrossWt)'GrossWt',Sum(GainLoss) 'GainLoss' From"
'    tmp = tmp & " (("
'    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight) + Sum(isnull(TG.WeightGainLossGSL,0)) 'GrossWt' , Sum(isnull(TG.WeightGainLossGSL,0)) 'GainLoss'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " )"
'
'    '==========================added transaction details on 19 apr 2013 ==========================
'    tmp = tmp & " Union All"
'    tmp = tmp & " ("
'    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'D'"
'    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
'    tmp = tmp & " From Txn_AdjTransaction TAD"
'    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
'    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType = 'AR' "
'    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'    tmp = tmp & " )"
'
'    tmp = tmp & " Union All"
'    tmp = tmp & " ("
'    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'W'"
'    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
'    tmp = tmp & " From Txn_AdjTransaction TAD"
'    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
'    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType <> 'AR' "
'    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'    tmp = tmp & " )"
''====================================================
'
'    tmp = tmp & " Union All"
'    tmp = tmp & " ("
'    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
'    tmp = tmp & " From Txn_Spillage TC"
'    tmp = tmp & " inner Join Mst_GSL MG On TC.GSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix  And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
'    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'    tmp = tmp & " )"
'    tmp = tmp & " Union All"
'    tmp = tmp & " ("
'    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
'    tmp = tmp & " From Txn_Spillage TC"
'    tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
'    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'    tmp = tmp & " )) TempTable"
'    tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity,TxnType"
    
        
        
    tmp = " Select StateName,LocationName,EQZone,Commodity,TxnType,GodownNo,Sum(NoOfBags)'NoOfBags',Sum(GrossWt)'GrossWt',"
    tmp = tmp & " Sum (GainLoss) 'GainLoss'"
    tmp = tmp & " From (("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType,MGod.GodownNo ,sum(TG.NoOfBags) 'NoOfBags',"
    tmp = tmp & " Sum (TG.GrossWeight) + Sum(IsNull(TG.WeightGainLossGSL, 0)) 'GrossWt' , Sum(isnull(TG.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_CxTF_Details TC"
    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTF_DetailsID=TG.CxTF_DetailsID and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_GSL MGSL on TG.Mst_GSL_ID = MGSL.Mst_GSL_ID ANd TG.SM_Prefix = MGSL.SM_Prefix"
    tmp = tmp & " inner Join Mst_Godown MGod on MGSL.GodownID = MGod.GodownID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'"
    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity,MGod.GodownNo )"
    tmp = tmp & " Union All ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'D',MGod.GodownNo ,sum(TGAD.NoOfBags) 'NoOfBags',"
    tmp = tmp & " Sum (TGAD.Quantity) + Sum(IsNull(TGAD.WeightGainLossGSL, 0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_GSL MGSL on TGAD.Mst_GSL_ID = MGSL.Mst_GSL_ID And TAD.SM_Prefix = MGSL.SM_Prefix"
    tmp = tmp & " inner Join Mst_Godown MGod on MGSL.GodownID = MGod.GodownID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType = 'AR'"
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )"
    tmp = tmp & " Union All ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'W',MGod.GodownNo ,sum(TGAD.NoOfBags) 'NoOfBags',"
    tmp = tmp & " Sum (TGAD.Quantity) + Sum(IsNull(TGAD.WeightGainLossGSL, 0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_GSL MGSL on TGAD.Mst_GSL_ID = MGSL.Mst_GSL_ID And TAD.SM_Prefix = MGSL.SM_Prefix"
    tmp = tmp & " inner Join Mst_Godown MGod on MGSL.GodownID = MGod.GodownID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType <> 'AR'"
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )"
    tmp = tmp & " Union All ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',MGod.GodownNo,sum(TC.SpillageNoOfBags) 'NoOfBags',"
    tmp = tmp & " Sum (TC.SpillageQuantity) 'GrossWt',0"
    tmp = tmp & " From Txn_Spillage TC"
    tmp = tmp & " inner Join Mst_GSL MG On TC.Mst_GSL_ID = MG.Mst_GSL_ID And TC.SM_Prefix = MG.SM_Prefix  And TC.Flag = 'S'"
    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_Godown MGod on MG.GodownID = MGod.GodownID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'"
    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )"
    tmp = tmp & " Union All ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',MGod.GodownNo,sum(TC.SpillageNoOfBags) 'NoOfBags',"
    tmp = tmp & " Sum (TC.SpillageQuantity) 'GrossWt',0"
    tmp = tmp & " From Txn_Spillage TC"
    tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.Mst_GSL_ID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_Godown MGod on MG.GodownID = MGod.GodownID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'"
    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone,MGod.GodownNo )) TempTable"
    tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity,TxnType,GodownNo"
    
    Call Tmp2Table
    
    If TmpRs2.RecordCount > 0 Then
       For r = 1 To TmpRs2.RecordCount
           Call IndataGrid1
           Call IndataGrid2
           TmpRs2.MoveNext
       Next
    End If
    Call CopyGrid3ToGrid1
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max


tmp = "select MS.StateName,MC.Commodity ,MCP.* from Mst_CommodityPrice MCP"
tmp = tmp & " inner join Mst_State MS on MCP.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MC on MCP.CommodityID=MC.CommodityID"
'tmp = tmp & " Where MCP.MonthYear='" & Left(CmbMonth.Text, 3) & "-" & mYear & "'"
tmp = tmp & " Where MCP.YearMonth='" & Format(CDate(mDate), "yyyy-mm") & "'"

Call Tmp3Table
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For I = 1 To MSHFlexGrid1.Rows - 2
  
    TmpRs3.Filter = "StateName='" & MSHFlexGrid1.TextMatrix(I, 0) & "' And Commodity='" & MSHFlexGrid1.TextMatrix(I, 3) & "' "
    mDayNo = Day(MSHFlexGrid1.TextMatrix(I, 6))
    If TmpRs3.RecordCount > 0 Then
       'MSHFlexGrid1.TextMatrix(I, 7) = TmpRs3.Fields(mDayNo + 4)
       
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs3.Fields(mDayNo + 4)), 0, TmpRs3.Fields(mDayNo + 4))
       
    Else
       MSHFlexGrid1.TextMatrix(I, 8) = 0
    End If
    MSHFlexGrid1.TextMatrix(I, 9) = Round(Val(MSHFlexGrid1.TextMatrix(I, 8)) * Val(MSHFlexGrid1.TextMatrix(I, 6)), 0)
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Value + 1
Call CalAvgPrice
Label2.Caption = MSHFlexGrid1.Rows - 2
CmdExcel.Enabled = True
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!"

End Sub
Private Sub CalAvgPrice()
Dim mTotAmt As Double
For I = 1 To MSHFlexGrid2.Rows - 2
    mTotAmt = 0
    For j = 0 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(j, 1) = MSHFlexGrid2.TextMatrix(I, 1) And MSHFlexGrid1.TextMatrix(j, 3) = MSHFlexGrid2.TextMatrix(I, 3) And MSHFlexGrid1.TextMatrix(j, 4) = MSHFlexGrid2.TextMatrix(I, 4) Then
           mTotAmt = mTotAmt + Val(MSHFlexGrid1.TextMatrix(j, 9))
        End If
    Next
    MSHFlexGrid2.TextMatrix(I, 9) = Round(mTotAmt)
    MSHFlexGrid2.TextMatrix(I, 10) = Round(Val(MSHFlexGrid2.TextMatrix(I, 9)) / mTotDays, 0)
Next

End Sub

Private Sub CopyGrid3ToGrid1()
For I = 1 To MSHFlexGrid3.Rows - 2
    mR = MSHFlexGrid1.Rows - 1
    For C = 0 To 7
        MSHFlexGrid1.TextMatrix(mR, C) = MSHFlexGrid3.TextMatrix(I, C)
    Next
    MSHFlexGrid1.TextMatrix(mR, 10) = MSHFlexGrid3.TextMatrix(I, 12)
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
End Sub
Private Sub IndataGrid2()
For m = 1 To MSHFlexGrid2.Rows - 1
    If TmpRs2!Locationname = MSHFlexGrid2.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid2.TextMatrix(m, 3) And TmpRs2!GodownNo = MSHFlexGrid2.TextMatrix(m, 4) Then
       If TmpRs2!TxnType = "D" Then
          MSHFlexGrid2.TextMatrix(m, 5) = Val(MSHFlexGrid2.TextMatrix(m, 5)) + TmpRs2!NoofBags
          MSHFlexGrid2.TextMatrix(m, 6) = Val(MSHFlexGrid2.TextMatrix(m, 6)) + TmpRs2!GrossWt
       Else
          MSHFlexGrid2.TextMatrix(m, 5) = Val(MSHFlexGrid2.TextMatrix(m, 5)) - TmpRs2!NoofBags
          MSHFlexGrid2.TextMatrix(m, 6) = Val(MSHFlexGrid2.TextMatrix(m, 6)) - TmpRs2!GrossWt
       End If
       If Val(MSHFlexGrid2.TextMatrix(m, 6)) < 0 Then
          MSHFlexGrid2.TextMatrix(m, 6) = 0
       End If
       MSHFlexGrid2.TextMatrix(m, 7) = Round(Val(MSHFlexGrid2.TextMatrix(m, 5)) / mTotDays, 0)
       MSHFlexGrid2.TextMatrix(m, 8) = Round(Val(MSHFlexGrid2.TextMatrix(m, 6)) / mTotDays, 3)
   End If
Next
End Sub
Private Sub IndataGrid1()
Dim flg, mRowNo As Variant
flg = "N"
For m = MSHFlexGrid3.Rows - 1 To 1 Step -1
    If TmpRs2!Locationname = MSHFlexGrid3.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid3.TextMatrix(m, 3) And TmpRs2!GodownNo = MSHFlexGrid3.TextMatrix(m, 4) And mDate = MSHFlexGrid3.TextMatrix(m, 7) Then
       flg = "Y"
       mRowNo = m
       Exit For
    End If
Next

If flg = "N" Then
   mRowNo = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.TextMatrix(mRowNo, 0) = TmpRs2!StateName
   MSHFlexGrid3.TextMatrix(mRowNo, 1) = TmpRs2!Locationname
   MSHFlexGrid3.TextMatrix(mRowNo, 2) = IIf(IsNull(TmpRs2!EQZone), "", TmpRs2!EQZone)
   MSHFlexGrid3.TextMatrix(mRowNo, 3) = TmpRs2!Commodity
   MSHFlexGrid3.TextMatrix(mRowNo, 4) = TmpRs2!GodownNo
   MSHFlexGrid3.TextMatrix(mRowNo, 7) = mDate
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If
If TmpRs2!TxnType = "D" Then
   MSHFlexGrid3.TextMatrix(mRowNo, 8) = TmpRs2!NoofBags
   MSHFlexGrid3.TextMatrix(mRowNo, 9) = TmpRs2!GrossWt
Else
   MSHFlexGrid3.TextMatrix(mRowNo, 10) = TmpRs2!NoofBags
   MSHFlexGrid3.TextMatrix(mRowNo, 11) = TmpRs2!GrossWt
End If
MSHFlexGrid3.TextMatrix(mRowNo, 12) = TmpRs2!GainLoss
MSHFlexGrid3.TextMatrix(mRowNo, 5) = Round(Val(MSHFlexGrid3.TextMatrix(mRowNo, 8)) - Val(MSHFlexGrid3.TextMatrix(mRowNo, 10)), 0)
MSHFlexGrid3.TextMatrix(mRowNo, 6) = Round(Val(MSHFlexGrid3.TextMatrix(mRowNo, 9)) - Val(MSHFlexGrid3.TextMatrix(mRowNo, 11)), 3)


'Commented on 12 Sept 2013 while testing of data mismatch for transaction and GSL wise
If Val(MSHFlexGrid3.TextMatrix(mRowNo, 5)) <= 0 Then
   MSHFlexGrid3.TextMatrix(mRowNo, 5) = 0
   MSHFlexGrid3.TextMatrix(mRowNo, 6) = 0
ElseIf Val(MSHFlexGrid3.TextMatrix(mRowNo, 5)) > 0 Then
   If Val(MSHFlexGrid3.TextMatrix(mRowNo, 6)) <= 0 Then
      MSHFlexGrid3.TextMatrix(mRowNo, 6) = 0
   End If
End If

If Val(MSHFlexGrid3.TextMatrix(mRowNo, 6)) < 0 Then
   If Val(MSHFlexGrid3.TextMatrix(mRowNo, 5)) <= 0 Then
      MSHFlexGrid3.TextMatrix(mRowNo, 5) = 0
   End If
   MSHFlexGrid3.TextMatrix(mRowNo, 6) = 0
End If

End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Call Grid_Head
Call Grid_Head1
CmdExcel.Enabled = False
TxtFromDate.Value = Date
End Sub

