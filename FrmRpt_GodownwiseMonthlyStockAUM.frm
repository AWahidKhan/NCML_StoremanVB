VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_GodownwiseMonthlyStockAUM 
   Caption         =   "Godown wise monthly Stock AUM"
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
      Height          =   1425
      Left            =   45
      TabIndex        =   13
      Top             =   60
      Width           =   15000
      Begin VB.ComboBox CmbExchangeType 
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
         Left            =   5865
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   2460
      End
      Begin VB.ComboBox CmbLocationName 
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
         Left            =   4305
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   510
         Width           =   2460
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
         Height          =   350
         Left            =   13125
         TabIndex        =   9
         Top             =   965
         Width           =   885
      End
      Begin VB.CommandButton CmdGo 
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
         Height          =   350
         Left            =   12240
         TabIndex        =   8
         Top             =   965
         Width           =   885
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
         Height          =   350
         Left            =   14010
         TabIndex        =   10
         Top             =   965
         Width           =   885
      End
      Begin VB.ComboBox CmbStateName 
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
         Left            =   1425
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   510
         Width           =   2790
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
         Left            =   6840
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   510
         Width           =   4665
      End
      Begin VB.ComboBox CmbGodown 
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
         Left            =   1425
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   2790
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   57671683
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1395
         TabIndex        =   2
         Top             =   525
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   57671683
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Exchange Type"
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
         Left            =   4320
         TabIndex        =   22
         Top             =   1020
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location Name"
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
         Left            =   4860
         TabIndex        =   21
         Top             =   195
         Width           =   1365
      End
      Begin VB.Label Label4 
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
         Left            =   2295
         TabIndex        =   20
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To "
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
         Left            =   1845
         TabIndex        =   19
         Top             =   195
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   495
         TabIndex        =   18
         Top             =   195
         Width           =   465
      End
      Begin VB.Label LblGodown 
         AutoSize        =   -1  'True
         Caption         =   "Godown No"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   8655
         TabIndex        =   16
         Top             =   195
         Width           =   1035
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Total Records Found :"
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
         Left            =   11520
         TabIndex        =   15
         Top             =   195
         Width           =   2325
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
         Left            =   13920
         TabIndex        =   14
         Top             =   195
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   60
      TabIndex        =   0
      Top             =   1485
      Width           =   14970
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   7545
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   255
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   12726
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
End
Attribute VB_Name = "FrmRpt_GodownwiseMonthlyStockAUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mGodownID, mStartDate, mEndDate, wc, mExchangeTypeID As Variant

Private Sub CmbLocationName_GotFocus()
If CmbStateName.Text = "" Then
   MsgBox "Plase Select State !!!"
   CmbStateName.SetFocus
   Exit Sub
End If

tmp = CmbLocationName.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location(" where StateID=" & mStateID)
Else
   Call TableMst_Location("where StateID=" & mStateID & " and " & Gen_WcSMLocation)
End If

CmbLocationName.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocationName.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationName.Text = tmp

End Sub

Private Sub CmbLocationName_LostFocus()
If CmbLocationName.Text = "" Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If


RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationName.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid PreFix Selection !!!!"
   CmbLocationName.SetFocus
   Exit Sub
End If
If tmp <> CmbLocationName.Text Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocationName.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbLocationName.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" Where LocationID = " & mLocationID)
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
   CmbGodown.Text = ""
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If tmp <> CmbCMPName.Text Then
   CmbGodown.Text = ""
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown("Where UnitTypeID in (1,6) And CMPID = " & mCMPID)
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then Exit Sub

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!godownid
End Sub

Private Sub CmbStateName_GotFocus()
tmp = CmbStateName.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateName.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateName.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateName.Text = tmp
End Sub

Private Sub CmbStateName_LostFocus()
If CmbStateName.Text = "" Then
   CmbLocationName.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateName.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateName.SetFocus
   Exit Sub
End If
If tmp <> CmbStateName.Text Then
   CmbLocationName.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
'TxtToDate.Value = Date
TxtFromDate.Value = Date
CmbStateName.Text = ""
CmbLocationName.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
CmbExchangeType.Text = ""
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownwiseMonthlyStockAUM.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownwiseMonthlyStockAUM.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1 ' + 6
   For c = 0 To MSHFlexGrid1.Cols - 3
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

Call Grid_Head
MsgBox "Wait till next message!!!"


If Format(TxtFromDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future Date not allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
'If Format(TxtToDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
'   MsgBox "Future Date not allowed!!!"
'   TxtToDate.SetFocus
'   Exit Sub
'End If
'If Format(TxtFromDate.Value, "yyyy-mm") > Format(TxtToDate.Value, "yyyy-mm") Then
'   MsgBox "From Date should not greater than To Date!!!"
'   TxtToDate.SetFocus
'   Exit Sub
'End If

wc = ""
If CmbGodown.Text <> "" Then
   wc = " and MG.GodownID = " & mGodownID
ElseIf CmbCMPName.Text <> "" Then
   wc = " and MCMP.CMPID = " & mCMPID
Else
   If CmbLocationName.Text <> "" Then
      wc = " and ML.LocationID = " & mLocationID
   ElseIf Gen_WcLocation <> "" Then
      wc = " and ML." & Gen_WcLocation
   ElseIf CmbStateName.Text <> "" Then
      wc = " and MS.StateID = " & mStateID
   ElseIf Gen_WcState <> "" Then
      wc = " and MS." & Gen_WcState
   End If
End If


mStartDate = Format(CDate(Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01"), "yyyy-mm-dd")
'mEndDate = Format(DateSerial(Year(TxtToDate.Value), Month(TxtToDate.Value) + 1, 0), "yyyy-mm-dd")
mEndDate = Format(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1, 0), "yyyy-mm-dd")

tmp = " select distinct timg.yearmonth, ms.statename, ml.locationname, mcmp.cmpname, mg.godownno,"
tmp = tmp & " mut.unittype, me.exchangetype,"
tmp = tmp & " convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) 'key', ml.locationid"
tmp = tmp & " from txn_invmongsldetail timg"
tmp = tmp & " inner join mst_godown mg on mg.godownid = timg.godownid"
tmp = tmp & " inner join mst_cmp mcmp on mcmp.cmpid = mg.cmpid"
tmp = tmp & " inner join mst_location ml on ml.locationid = mcmp.locationid"
tmp = tmp & " inner join mst_state ms on ms.stateid = ml.stateid"
tmp = tmp & " inner join mst_exchangetype me on me.exchangetypeid = timg.exchangetypeid"
tmp = tmp & " left join mst_unittype mut on mut.unittypeid = mg.unittypeid"
tmp = tmp & " where timg.yearmonth between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & wc
tmp = tmp & " Order by timg.yearmonth, ms.statename, ml.locationname, mcmp.cmpname, mg.godownno"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   '---pledge stock
'   tmp = " select sum(mgsl.pledgeweight) stock, convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) 'key'"
'   tmp = tmp & " from txn_invmongsldetail timg"
'   tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
'   tmp = tmp & " where mgsl.pledgeweight > 0 and timg.yearmonth between '" & mStartDate & "' and '" & mEndDate & "'"
'   tmp = tmp & " group by timg.yearmonth, timg.godownid, timg.exchangetypeid"
'   Call Tmp2Table
   
   '---scm stock
   tmp = " select sum(timg.balanceweight) stock, convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) 'key'"
   tmp = tmp & " from txn_invmongsldetail timg"
   tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
   tmp = tmp & " where timg.balanceweight > 0 and timg.yearmonth between '" & mStartDate & "' and '" & mEndDate & "' and mgsl.scmflag = 1"
   tmp = tmp & " group by timg.yearmonth, timg.godownid, timg.exchangetypeid"
   'tmp = tmp & " and mgsl.pledgeweight = 0
   Call Tmp3Table
   
   '---for other than scm
   tmp = " select sum(timg.balanceweight) stock, convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) 'key'"
   tmp = tmp & " from txn_invmongsldetail timg"
   tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
   tmp = tmp & " where timg.balanceweight > 0 and timg.yearmonth between '" & mStartDate & "' and '" & mEndDate & "' and mgsl.scmflag <> 1"
   tmp = tmp & " group by timg.yearmonth, timg.godownid, timg.exchangetypeid"
   'tmp = tmp & " and mgsl.pledgeweight = 0
   Call Tmp4Table
   
   '---for scm AUM
   tmp = " select key1 'key', commodityid, sum(balanceqty*price) aum, avg(price) price from"
   tmp = tmp & " (select convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) 'key1',timg.commodityid, sum(timg.balanceweight) BalanceQty,"
   tmp = tmp & " (Select top 1 isnull(MCP.Day31,isnull(MCP.Day30,isnull(MCP.Day29,MCP.Day28))) from Mst_CommodityPrice MCP"
   tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MCP.StateID"
   tmp = tmp & " Where yearmonth = Left(Convert(VarChar, timg.yearmonth, 120), 7)"
   tmp = tmp & " AND ML.SM_Prefix = TIMG.SM_Prefix And MCP.COmmodityID = TIMG.CommodityID) price"
   tmp = tmp & " from txn_invmongsldetail timg"
   tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
   tmp = tmp & " Where timg.balanceweight > 0 And mgsl.scmflag = 1"
   tmp = tmp & " and timg.yearmonth between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " group by timg.godownid, timg.exchangetypeid, timg.commodityid,  timg.yearmonth, timg.sm_prefix"
   tmp = tmp & " ) a group by key1, commodityid"
   Call Tmp5Table
   
   '---for scm AUM
   tmp = " select key1 'key', commodityid, sum(balanceqty*price) aum, avg(price) price from"
   tmp = tmp & " (select convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) 'key1',timg.commodityid, sum(timg.balanceweight) BalanceQty,"
   tmp = tmp & " (Select top 1 isnull(MCP.Day31,isnull(MCP.Day30,isnull(MCP.Day29,MCP.Day28))) from Mst_CommodityPrice MCP"
   tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MCP.StateID"
   tmp = tmp & " Where yearmonth = Left(Convert(VarChar, timg.yearmonth, 120), 7)"
   tmp = tmp & " AND ML.SM_Prefix = TIMG.SM_Prefix And MCP.COmmodityID = TIMG.CommodityID) price"
   tmp = tmp & " from txn_invmongsldetail timg"
   tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
   tmp = tmp & " Where timg.balanceweight > 0 And mgsl.scmflag <> 1"
   tmp = tmp & " and timg.yearmonth between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " group by timg.godownid, timg.exchangetypeid, timg.commodityid,  timg.yearmonth, timg.sm_prefix"
   tmp = tmp & " ) a group by key1, commodityid"
   Call Tmp6Table
   
   CmdExcel.Enabled = True
   LblTotalRecords.Caption = TmpRs.RecordCount
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      ProgressBar1.Max = .Rows
      For I = 2 To TmpRs.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!yearmonth), "", Format(TmpRs!yearmonth, "MMM-yyyy"))
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
         .TextMatrix(I, 13) = IIf(IsNull(TmpRs!Key), "", TmpRs!Key)
         .TextMatrix(I, 14) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
'         If TmpRs2.RecordCount > 0 Then
'            TmpRs2.MoveFirst
'            TmpRs2.Find "Key = '" & .TextMatrix(I, 13) & "'"
'            If Not TmpRs2.EOF Then
'               .TextMatrix(I, 7) = IIf(IsNull(TmpRs2!stock), "", TmpRs2!stock)
'               .TextMatrix(0, 7) = Val(.TextMatrix(I, 7)) + Val(.TextMatrix(0, 7))
'            End If
'         End If
         If TmpRs3.RecordCount > 0 Then
            TmpRs3.MoveFirst
            TmpRs3.Find "Key = '" & .TextMatrix(I, 13) & "'"
            If Not TmpRs3.EOF Then
               .TextMatrix(I, 7) = IIf(IsNull(TmpRs3!stock), "", TmpRs3!stock)
               .TextMatrix(0, 7) = Val(.TextMatrix(I, 7)) + Val(.TextMatrix(0, 7))
            End If
         End If
         If TmpRs5.RecordCount > 0 Then
            TmpRs5.MoveFirst
            TmpRs5.Find "Key = '" & .TextMatrix(I, 13) & "'"
            If Not TmpRs5.EOF Then
               .TextMatrix(I, 8) = IIf(IsNull(TmpRs5!Price), "", TmpRs5!Price)
               .TextMatrix(I, 9) = IIf(IsNull(TmpRs5!AUM), "", TmpRs5!AUM)
               .TextMatrix(0, 9) = Val(.TextMatrix(I, 9)) + Val(.TextMatrix(0, 9))
            End If
         End If
         If TmpRs4.RecordCount > 0 Then
            TmpRs4.MoveFirst
            TmpRs4.Find "Key = '" & .TextMatrix(I, 13) & "'"
            If Not TmpRs4.EOF Then
               .TextMatrix(I, 10) = IIf(IsNull(TmpRs4!stock), "", TmpRs4!stock)
               .TextMatrix(0, 10) = Val(.TextMatrix(I, 10)) + Val(.TextMatrix(0, 10))
            End If
         End If
         If TmpRs6.RecordCount > 0 Then
            TmpRs6.MoveFirst
            TmpRs6.Find "Key = '" & .TextMatrix(I, 13) & "'"
            If Not TmpRs6.EOF Then
               .TextMatrix(I, 11) = IIf(IsNull(TmpRs6!Price), "", TmpRs6!Price)
               .TextMatrix(I, 12) = IIf(IsNull(TmpRs6!AUM), "", TmpRs6!AUM)
               .TextMatrix(0, 12) = Val(.TextMatrix(I, 12)) + Val(.TextMatrix(0, 12))
            End If
         End If
         TmpRs.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
   MsgBox "Done!!!"
   ProgressBar1.Value = ProgressBar1.Max
Else
   MsgBox "No record found!!!"
End If
End Sub

Private Sub CmbExchangeType_GotFocus()
tmp = CmbExchangeType.Text
Call TableMst_ExchangeType
CmbExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeType.Text = tmp
End Sub

Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click
End Sub

Private Sub Grid_Head()
ProgressBar1.Value = 0
LblTotalRecords.Caption = ""
CmdExcel.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 3
   .FixedRows = 2
   .WordWrap = True
   .Cols = 15
      
   .TextMatrix(1, 0) = "Year Month"
   .TextMatrix(1, 1) = "State Name"
   .TextMatrix(1, 2) = "Location Name"
   .TextMatrix(1, 3) = "CMP Name"
   .TextMatrix(1, 4) = "Godown No"
   .TextMatrix(1, 5) = "Godown Type"
   .TextMatrix(1, 6) = "Exchange Type"
'   .TextMatrix(1, 7) = "Pledge Stock"
'   .TextMatrix(1, 8) = "Price"
'   .TextMatrix(1, 9) = "AUM"
   .TextMatrix(1, 7) = "SCM Stock"
   .TextMatrix(1, 8) = "Price"
   .TextMatrix(1, 9) = "AUM"
   .TextMatrix(1, 10) = "Other Stock"
   .TextMatrix(1, 11) = "Price"
   .TextMatrix(1, 12) = "AUM"
   .TextMatrix(1, 13) = "Key"
   .TextMatrix(1, 14) = "LocationID"
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 2200
   .ColWidth(4) = 1500
   .ColWidth(5) = 1500
   .ColWidth(6) = 1500
   .ColWidth(13) = 0
   .ColWidth(14) = 0
   
   .RowHeight(1) = 600
End With
End Sub

'Private Sub GetAUMDetail()
'Dim mLocationID_ As Double
'Dim mKey_ As String, mKeyPrice As String
'
'ProgressBar1.Value = 0
'If TmpRs1.RecordCount > 0 Then
'   For I = 2 To MSHFlexGrid1.Rows - 2
'      mKey_ = MSHFlexGrid1.TextMatrix(I, 13)
'      mLocationID_ = Val(MSHFlexGrid1.TextMatrix(I, 14))
'   '========for pledge BalanceQty
'   '   tmp = " select mgsl.commodityid, sum(mgsl.pledgeweight) BalanceQty"
'   '   tmp = tmp & " from txn_invmongsldetail timg"
'   '   tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
'   '   tmp = tmp & " where convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) = '" & mKey_ & "'"
'   '   tmp = tmp & " and mgsl.pledgeweight > 0 "
'   '   tmp = tmp & " group by mgsl.commodityid"
'   '   Call TmpTable
'   '
'   '   mPrice = 0
'   '   mTotalAUM = 0
'   '
'   '   If TmpRs.RecordCount > 0 Then
'   '       For j = 1 To TmpRs.RecordCount
'   '           tmp = " Select MCP.* From Mst_CommodityPrice MCP"
'   '           tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
'   '           tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
'   '           tmp = tmp & " Where ML.LocationID = " & mLocationID_ & " And MCP.CommodityID = " & TmpRs!CommodityID
'   '           tmp = tmp & " And YearMonth = ("
'   '           tmp = tmp & " Select Max(YearMonth) From Mst_CommodityPrice MCP"
'   '           tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
'   '           tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
'   '           tmp = tmp & " Where ML.LocationID = " & mLocationID_ & " And MCP.CommodityID = " & TmpRs!CommodityID
'   '           tmp = tmp & " )"
'   '           Call Tmp1Table
'   '
'   '           If TmpRs1.RecordCount > 0 Then
'   '               If IsNull(TmpRs1!Day31) = False Then
'   '                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day31)
'   '                   mPrice = Val(mPrice) + TmpRs1!Day31
'   '               ElseIf IsNull(TmpRs1!Day30) = False Then
'   '                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day30)
'   '                   mPrice = Val(mPrice) + TmpRs1!Day30
'   '               ElseIf IsNull(TmpRs1!Day29) = False Then
'   '                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day29)
'   '                   mPrice = Val(mPrice) + TmpRs1!Day29
'   '               ElseIf IsNull(TmpRs1!Day28) = False Then
'   '                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day28)
'   '                   mPrice = Val(mPrice) + TmpRs1!Day28
'   '               End If
'   '           End If
'   '           TmpRs.MoveNext
'   '       Next
'   '   MSHFlexGrid1.TextMatrix(I, 8) = Format(Val(mPrice / TmpRs.RecordCount), "#0")
'   '   End If
'   '   MSHFlexGrid1.TextMatrix(I, 9) = Format(Val(mTotalAUM), "#0")
'   '   MSHFlexGrid1.TextMatrix(0, 9) = Format(Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(0, 9)), "#0")
'   '
'   ''===========for scm BalanceQty
'SCMStock:
'If Val(MSHFlexGrid1.TextMatrix(I, 7)) = 0 Then GoTo OtherStock
'      tmp = " select timg.commodityid, sum(timg.balanceweight) BalanceQty"
'      tmp = tmp & " from txn_invmongsldetail timg"
'      tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
'      tmp = tmp & " where timg.balanceweight > 0 and mgsl.scmflag = 1"
'      tmp = tmp & " and convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) = '" & mKey_ & "'"
'      tmp = tmp & " and mgsl.pledgeweight = 0 "
'      tmp = tmp & " group by timg.commodityid"
'      Call TmpTable
'
'      mPrice = 0
'      mTotalAUM = 0
'
'      If TmpRs.RecordCount > 0 Then
'         For j = 1 To TmpRs.RecordCount
'            mKeyPrice = mLocationID_ & "_" & TmpRs!CommodityID
'            TmpRs1.MoveFirst
'            TmpRs1.Find "Key = " & mKeyPrice
'            If Not TmpRs1.EOF Then
'               If IsNull(TmpRs1!Day31) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day31)
'                   mPrice = Val(mPrice) + TmpRs1!Day31
'               ElseIf IsNull(TmpRs1!Day30) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day30)
'                   mPrice = Val(mPrice) + TmpRs1!Day30
'               ElseIf IsNull(TmpRs1!Day29) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day29)
'                   mPrice = Val(mPrice) + TmpRs1!Day29
'               ElseIf IsNull(TmpRs1!Day28) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day28)
'                   mPrice = Val(mPrice) + TmpRs1!Day28
'               End If
'            End If
'            TmpRs.MoveNext
'         Next
'         MSHFlexGrid1.TextMatrix(I, 8) = Format(Val(mPrice / TmpRs.RecordCount), "#0")
'      End If
'      MSHFlexGrid1.TextMatrix(I, 9) = Format(Val(mTotalAUM), "#0")
'      MSHFlexGrid1.TextMatrix(0, 9) = Format(Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(0, 9)), "#0")
'
'   ''===========for other than scm BalanceQty
'OtherStock:
'If Val(MSHFlexGrid1.TextMatrix(I, 10)) = 0 Then GoTo ExitRow
'      tmp = " select timg.commodityid, sum(timg.balanceweight) BalanceQty"
'      tmp = tmp & " from txn_invmongsldetail timg"
'      tmp = tmp & " inner join mst_gsl mgsl on mgsl.gslid = timg.gslid and mgsl.sm_prefix = timg.sm_prefix"
'      tmp = tmp & " where timg.balanceweight > 0 and mgsl.scmflag <> 1"
'      tmp = tmp & " and convert(varchar ,timg.yearmonth, 111)+'_'+convert(varchar ,timg.godownid)+'_'+convert(varchar ,timg.exchangetypeid) = '" & mKey_ & "'"
'      tmp = tmp & " and mgsl.pledgeweight = 0 "
'      tmp = tmp & " group by timg.commodityid"
'      Call TmpTable
'
'      mPrice = 0
'      mTotalAUM = 0
'
'      If TmpRs.RecordCount > 0 Then
'         For j = 1 To TmpRs.RecordCount
'            mKeyPrice = mLocationID_ & "_" & TmpRs!CommodityID
'            TmpRs1.MoveFirst
'            TmpRs1.Find "Key = " & mKeyPrice
'            If Not TmpRs1.EOF Then
'               If IsNull(TmpRs1!Day31) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day31)
'                   mPrice = Val(mPrice) + TmpRs1!Day31
'               ElseIf IsNull(TmpRs1!Day30) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day30)
'                   mPrice = Val(mPrice) + TmpRs1!Day30
'               ElseIf IsNull(TmpRs1!Day29) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day29)
'                   mPrice = Val(mPrice) + TmpRs1!Day29
'               ElseIf IsNull(TmpRs1!Day28) = False Then
'                   mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day28)
'                   mPrice = Val(mPrice) + TmpRs1!Day28
'               End If
'            End If
'            TmpRs.MoveNext
'         Next
'         MSHFlexGrid1.TextMatrix(I, 11) = Format(Val(mPrice / TmpRs.RecordCount), "#0")
'      End If
'      MSHFlexGrid1.TextMatrix(I, 12) = Format(Val(mTotalAUM), "#0")
'      MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(MSHFlexGrid1.TextMatrix(0, 12)), "#0")
'ExitRow:
'      ProgressBar1.Value = ProgressBar1.Value + 1
'   Next
'End If
'End Sub

