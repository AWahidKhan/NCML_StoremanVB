VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_GodownUtilizationNew 
   Caption         =   "Godown Utilization Report"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   7290
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   15000
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
         Left            =   13560
         TabIndex        =   8
         Top             =   930
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
         Left            =   13560
         TabIndex        =   6
         Top             =   180
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
         Left            =   13560
         TabIndex        =   7
         Top             =   550
         Width           =   1215
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   380
         Width           =   3900
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
         Left            =   5640
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   380
         Width           =   3600
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
         Left            =   9300
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   380
         Width           =   4125
      End
      Begin VB.ComboBox CmbGodownID 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1000
         Width           =   5460
      End
      Begin VB.ComboBox CmbUnitTypeID 
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
         Left            =   5640
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1000
         Width           =   3600
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   380
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   111935489
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MonthYear"
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
         Left            =   360
         TabIndex        =   20
         Top             =   120
         Width           =   975
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
         TabIndex        =   19
         Top             =   1050
         Width           =   90
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
         TabIndex        =   18
         Top             =   1050
         Width           =   2325
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3150
         TabIndex        =   17
         Top             =   120
         Width           =   1065
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
         Left            =   7125
         TabIndex        =   16
         Top             =   120
         Width           =   765
      End
      Begin VB.Label LblCMPName 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   11115
         TabIndex        =   15
         Top             =   120
         Width           =   435
      End
      Begin VB.Label LblGodownID 
         AutoSize        =   -1  'True
         Caption         =   "Godown"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   750
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Godown Type"
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
         Left            =   6945
         TabIndex        =   13
         Top             =   750
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8295
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   150
         Left            =   120
         TabIndex        =   11
         Top             =   8100
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7815
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13785
         _Version        =   393216
         AllowUserResizing=   3
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownUtilizationNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFrmDate, mStateID, mLocationID, mCMPID, mGodownID, mUnitTypeID As Variant
Dim mR, mC As Double

''-- State Name & ID
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

''-- Location Name & ID
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
''-- CMP Name & ID
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
   tmp1 = " And LocationID = " & mLocationID
End If
Call TableMst_CMP(" where CloseDate is null And LocationID =" & mLocationID)
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

''-- Godown Name & ID
Private Sub CmbGodownID_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownID.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownID.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub
Private Sub CmbGodownID_LostFocus()
If CmbGodownID.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodownID.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

''-- Godown Type & ID
Private Sub CmbUnitTypeID_GotFocus()
tmp = CmbUnitTypeID.Text
CmbUnitTypeID.Clear
Call TableMst_UnitType
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record Found !!"
   CmbUnitTypeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_UnitType.RecordCount
    CmbUnitTypeID.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbUnitTypeID.Text = tmp
End Sub
Private Sub CmbUnitTypeID_LostFocus()
If CmbUnitTypeID.Text = "" Then Exit Sub
RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType='" & CmbUnitTypeID.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid Selection !!"
   CmbUnitTypeID.SetFocus
   Exit Sub
End If
mUnitTypeID = RsMst_UnitType!UnitTypeID
End Sub



Private Sub CmdClear_Click()
Call Grid_Head
lblRecordCount.Caption = ""
CmdExcel.Enabled = False
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbGodownID.Text = ""
CmbUnitTypeID.Text = ""
TxtFromDate.Value = CDate(Date)
ProgressBar1.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownUtilizationNew.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownUtilizationNew.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
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
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim mVal As Variant
Dim wc As String
Dim tmp1 As String
wc = ""
tmp1 = ""
Me.MousePointer = vbHourglass
ProgressBar1.Value = 0
lblRecordCount.Caption = "-"
CmdExcel.Enabled = True
Call Grid_Head

mFrmDate = Format(TxtFromDate, "YYYY-MM-01")

If CmbStateID.Text <> "" Then
   wc = " AND S.StateID = " & mStateID
End If

If CmbLocationID.Text <> "" Then
    wc = wc & " AND L.LocationID = " & mLocationID
End If

If CmbCMPName.Text <> "" Then
    wc = wc & " AND C.CMPID = " & mCMPID
End If

If CmbGodownID.Text <> "" Then
    wc = wc & " AND G.GodownID = " & mGodownID
End If

If CmbUnitTypeID.Text <> "" Then
    wc = wc & " AND U.UnitTypeID = " & mUnitTypeID
End If

tmp = " SELECT a.YearMonth,s.statename,l.locationname,c.CMPName,g.godownno,g.StartDate,g.CloseDate,u.UnitType,E.ExchangeType,g.GodownID, "
tmp = tmp & " g.GodownCapacity * a.CapacityNoOfDay 'Capacity' ,sum(a.WtQty) 'Qty', 0 'Utilization' "
tmp = tmp & " from Txn_InvMonGSLDetail a "
tmp = tmp & " inner join Mst_Godown g on a.godownid = g.godownid "
tmp = tmp & " inner join Mst_CMP c on g.cmpid = c.cmpid "
tmp = tmp & " inner join Mst_location l on c.locationid = l.locationid "
tmp = tmp & " inner join Mst_State s on s.stateid = l.StateID "
tmp = tmp & " inner join Mst_UnitType u on g.UnitTypeID = u.UnitTypeID "
tmp = tmp & " inner join Mst_ExchangeType E on a.ExchangeTypeID = E.ExchangeTypeID "
'tmp = tmp & " Where a.YearMonth = '2015-10-01' And (g.CloseDate >= '2015-10-01' Or  g.CloseDate is null) "
tmp = tmp & " Where a.YearMonth = '" & mFrmDate & "' And (g.CloseDate >= '" & mFrmDate & "' Or  g.CloseDate is null) "
tmp = tmp & wc
tmp = tmp & " group by a.YearMonth,s.statename,l.locationname,c.CMPName,g.godownno,g.StartDate,g.CloseDate,u.UnitType,E.ExchangeType, "
tmp = tmp & " g.GodownCapacity , a.CapacityNoOfDay, g.GodownID "
'tmp = tmp & " order by a.YearMonth,s.statename,l.locationname,c.CMPName,g.godownno"
tmp = tmp & "Union All "
tmp = tmp & " SELECT a.YearMonth,s.statename,l.locationname,c.CMPName,g.godownno,g.StartDate,g.CloseDate,u.UnitType,E.ExchangeType,g.GodownID, "
tmp = tmp & " g.GodownCapacity * a.CapacityNoOfDay 'Capacity' ,sum(a.WtQty) 'Qty', 0 'Utilization' "
tmp = tmp & " from Txn_InvMonGSLDetailUpto2013 a "
tmp = tmp & " inner join Mst_Godown g on a.godownid = g.godownid "
tmp = tmp & " inner join Mst_CMP c on g.cmpid = c.cmpid "
tmp = tmp & " inner join Mst_location l on c.locationid = l.locationid "
tmp = tmp & " inner join Mst_State s on s.stateid = l.StateID "
tmp = tmp & " inner join Mst_UnitType u on g.UnitTypeID = u.UnitTypeID "
tmp = tmp & " inner join Mst_ExchangeType E on a.ExchangeTypeID = E.ExchangeTypeID "
'tmp = tmp & " Where a.YearMonth = '2015-10-01' And (g.CloseDate >= '2015-10-01' Or  g.CloseDate is null) "
tmp = tmp & " Where a.YearMonth = '" & mFrmDate & "' And (g.CloseDate >= '" & mFrmDate & "' Or  g.CloseDate is null) "
tmp = tmp & wc
tmp = tmp & " group by a.YearMonth,s.statename,l.locationname,c.CMPName,g.godownno,g.StartDate,g.CloseDate,u.UnitType,E.ExchangeType, "
tmp = tmp & " g.GodownCapacity , a.CapacityNoOfDay, g.GodownID "
tmp = tmp & " Union All "
tmp = tmp & " SELECT a.YearMonth,s.statename,l.locationname,c.CMPName,g.godownno,g.StartDate,g.CloseDate,u.UnitType,E.ExchangeType,g.GodownID, "
tmp = tmp & " g.GodownCapacity * a.CapacityNoOfDay 'Capacity' ,sum(a.WtQty) 'Qty', 0 'Utilization' "
tmp = tmp & " from Txn_InvMonGSLDetail2014to2015 a "
tmp = tmp & " inner join Mst_Godown g on a.godownid = g.godownid "
tmp = tmp & " inner join Mst_CMP c on g.cmpid = c.cmpid "
tmp = tmp & " inner join Mst_location l on c.locationid = l.locationid "
tmp = tmp & " inner join Mst_State s on s.stateid = l.StateID "
tmp = tmp & " inner join Mst_UnitType u on g.UnitTypeID = u.UnitTypeID "
tmp = tmp & " inner join Mst_ExchangeType E on a.ExchangeTypeID = E.ExchangeTypeID "
'tmp = tmp & " Where a.YearMonth = '2015-10-01' And (g.CloseDate >= '2015-10-01' Or  g.CloseDate is null) "
tmp = tmp & " Where a.YearMonth = '" & mFrmDate & "' And (g.CloseDate >= '" & mFrmDate & "' Or  g.CloseDate is null) "
tmp = tmp & wc
tmp = tmp & " group by a.YearMonth,s.statename,l.locationname,c.CMPName,g.godownno,g.StartDate,g.CloseDate,u.UnitType,E.ExchangeType, "
tmp = tmp & " g.GodownCapacity , a.CapacityNoOfDay, g.GodownID "

Call TmpTable

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2

With MSHFlexGrid1
    For I = 1 To TmpRs.RecordCount
        .TextMatrix(I, 0) = IIf(IsNull(TmpRs!YearMonth), "", Format(TmpRs!YearMonth, "mmm-YYYY"))
        .TextMatrix(I, 1) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        .TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
        .TextMatrix(I, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
        .TextMatrix(I, 5) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "DD-mmm-YYYY"))
        .TextMatrix(I, 6) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "DD-mmm-YYYY"))
        .TextMatrix(I, 7) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
        .TextMatrix(I, 8) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
        .TextMatrix(I, 9) = IIf(IsNull(TmpRs!Capacity), 0, TmpRs!Capacity)
        .TextMatrix(I, 10) = IIf(IsNull(TmpRs!Qty), 0, TmpRs!Qty)
        If Val(.TextMatrix(I, 9)) <> 0 Then
           .TextMatrix(I, 11) = Format(Val(.TextMatrix(I, 10)) * 100 / Val(.TextMatrix(I, 9)), "#######0.00") 'IIf(IsNull(TmpRs!Utilization), "", TmpRs!Utilization)
        Else
           .TextMatrix(I, 11) = "0.00" 'IIf(IsNull(TmpRs!Utilization), "", TmpRs!Utilization)
        End If
        ProgressBar1.Value = ProgressBar1.Value + 1
        .Rows = .Rows + 1
        TmpRs.MoveNext
    Next
End With
Me.MousePointer = vbDefault
ProgressBar1.Value = ProgressBar1.Max
lblRecordCount.Caption = TmpRs.RecordCount
MsgBox "Done.", vbOKOnly + vbInformation, App.Title
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
Call CmdClear_Click
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 12
    .FixedRows = 1
    .WordWrap = True
    .RowHeight(0) = 500
    .Font.Size = 10
    .TextMatrix(0, 0) = "Month-Year":                           .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "State":                                .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "Location":                             .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "CMP":                                  .ColWidth(3) = 2000
    .TextMatrix(0, 4) = "Godown No":                            .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "Godown Start Date":                    .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "Godown Close Date":                    .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "Godown Type":                          .ColWidth(7) = 2000
    .TextMatrix(0, 8) = "Exchange Type":                        .ColWidth(8) = 2000
    .TextMatrix(0, 9) = "Capacity":                             .ColWidth(9) = 2000
    .TextMatrix(0, 10) = "Qty in MT":                           .ColWidth(10) = 2000
    .TextMatrix(0, 11) = "Utilization in %":                    .ColWidth(11) = 2000
End With
End Sub
