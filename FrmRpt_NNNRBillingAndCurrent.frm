VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_NNNRBillingAndCurrent 
   Caption         =   "NNNR Billing vs Current Rate Card Details "
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
      Height          =   1590
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbCommodityGroupID 
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
         Left            =   10455
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   375
         Width           =   4515
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
         Left            =   6315
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   375
         Width           =   4080
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
         Left            =   1875
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   375
         Width           =   4395
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
         Left            =   13710
         TabIndex        =   6
         Top             =   1110
         Width           =   1095
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
         Left            =   11520
         TabIndex        =   5
         Top             =   1110
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
         Left            =   12615
         TabIndex        =   4
         Top             =   1110
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   75
         TabIndex        =   15
         Top             =   375
         Width           =   1755
         _ExtentX        =   3096
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
         Format          =   115277825
         CurrentDate     =   39884
      End
      Begin VB.Label Label23 
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   16
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Commodity Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11857
         TabIndex        =   14
         Top             =   120
         Width           =   1710
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
         Left            =   7973
         TabIndex        =   12
         Top             =   120
         Width           =   765
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
         Left            =   3540
         TabIndex        =   11
         Top             =   120
         Width           =   1065
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
         Left            =   10080
         TabIndex        =   10
         Top             =   1170
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
         Left            =   7680
         TabIndex        =   9
         Top             =   1170
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8220
      Left            =   120
      TabIndex        =   0
      Top             =   1605
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   7755
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
         _Version        =   393216
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
Attribute VB_Name = "FrmRpt_NNNRBillingAndCurrent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mStateID, mLOcationID, mSM_Prefix, mCommodityId, mCommodityGroupId As Variant

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocationID.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
CmbLocationID.Clear

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLOcationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLOcationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub


Private Sub CmbCommodityGroupID_GotFocus()

tmp = CmbCommodityGroupID.Text
Call TableMst_CommodityGroup
CmbCommodityGroupID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_CommodityGroup.RecordCount
   CmbCommodityGroupID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbCommodityGroupID.Text = tmp
    
End Sub

Private Sub CmbCommodityGroupID_LostFocus()
If CmbCommodityGroupID.Text = "" Then
   mCommodityGroupId = Null
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbCommodityGroupID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityGroupID.SetFocus
   Exit Sub
End If
mCommodityGroupId = RsMst_CommodityGroup!CommodityGroupID

End Sub

Private Sub CmdClear_Click()
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCommodityGroupID.Text = ""
TxtFromDate = Date
TxtFromDate.Value = CDate(Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01") 'GetFirstDateOfMonth(TxtFromDate.Value)
lblRecordCount.Caption = "-"
ProgressBar1.Value = 0
Call Grid_Head
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NNNRBillingAndCurrentRate.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NNNRBillingAndCurrentRate.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Max = MSHFlexGrid1.Rows + 5
ProgressBar1.Value = 0
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "mm/dd/yyyy")
           Else
              If C = 7 Then
                 oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(i, C)
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
              End If
           End If
        Else
'           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           If C = 7 Then
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(i, C)
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
Set oWB = Nothing
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

Dim wc As Variant

wc = ""
Gen_mTimeStamp = GetTimeStamp
Call Grid_Head
lblRecordCount.Caption = "-"

MsgBox "Wait till next message!!!"

ProgressBar1.Value = 0

If CmbStateID.Text <> "" Then
   wc = wc & " And  ms.STateName = '" & CmbStateID.Text & "'"
End If

If CmbLocationID.Text <> "" Then
   wc = wc & " ANd ml.LocationName = '" & CmbLocationID.Text & "'"
End If

If CmbCommodityGroupID.Text <> "" Then
   wc = wc & " ANd mc.CommodityGroup = '" & CmbCommodityGroupID.Text & "'"
End If

tmp = "Select Distinct YearMOnth,StateName,LocationName,LocationID,CommodityGroup,CommodityGroupID,BillingCycle,BillingUnit,Rate,[Key]"
tmp = tmp & " From (Select Distinct a.YearMOnth,MS.StateName,ML.LocationName,ML.LocationID,MC.CommodityGroup,MC.CommodityGroupID,MB.BillingCycle,a.BillingUnit,a.Rate " ' --,a.BagSize"
tmp = tmp & " , Convert(varchar,ml.LocationID)+'-'+Convert(Varchar,mc.CommodityGroupID) 'Key'"
tmp = tmp & " from Txn_InvMOnGSLDEtail a"
tmp = tmp & " inner join mst_location ml on a.SM_Prefix = ml.SM_Prefix"
tmp = tmp & " inner join mst_commodity mcom On a.CommodityID = mcom.commodityID"
tmp = tmp & " inner join Mst_CommodityGRoup mc on mcom.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_billingcycle mb on a.BillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " Where a.YearMonth = '" & Format(TxtFromDate, "yyyy-mm-dd") & "' And a.ExchangeTypeID = 1 And a.Flag1 = 'N' and a.Rate is not null and a.TFlag in ('D') ANd a.BillingUnit = 'B'"
tmp = tmp & wc
tmp = tmp & " Union All"
tmp = tmp & " Select Distinct a.YearMOnth,MS.StateName,ML.LocationName,ML.LocationID,MC.CommodityGroup,MC.CommodityGroupID,MB.BillingCycle,a.BillingUnit,a.Rate" '--,a.BagSize"
tmp = tmp & " , Convert(varchar,ml.LocationID)+'-'+Convert(Varchar,mc.CommodityGroupID) 'Key'"
tmp = tmp & " from Txn_InvMOnGSLDEtail a"
tmp = tmp & " inner join mst_location ml on a.SM_Prefix = ml.SM_Prefix"
tmp = tmp & " inner join mst_commodity mcom On a.CommodityID = mcom.commodityID"
tmp = tmp & " inner join Mst_CommodityGRoup mc on mcom.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_billingcycle mb on a.BillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " Where a.YearMonth = '" & Format(TxtFromDate, "yyyy-mm-dd") & "' And a.ExchangeTypeID = 1 And a.Flag1 = 'N' and a.Rate is not null and a.TFlag in ('O','W') ANd a.BillingUnit = 'B'"
tmp = tmp & wc
tmp = tmp & " Union All"
tmp = tmp & " Select Distinct a.YearMOnth,MS.StateName,ML.LocationName,ML.LocationID,MC.CommodityGroup,MC.CommodityGroupID,MB.BillingCycle,a.BillingUnit,a.Rate"
tmp = tmp & " , Convert(varchar,ml.LocationID)+'-'+Convert(Varchar,mc.CommodityGroupID) 'Key'"
tmp = tmp & " from Txn_InvMOnGSLDEtail a"
tmp = tmp & " inner join mst_location ml on a.SM_Prefix = ml.SM_Prefix"
tmp = tmp & " inner join mst_commodity mcom On a.CommodityID = mcom.commodityID"
tmp = tmp & " inner join Mst_CommodityGRoup mc on mcom.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_billingcycle mb on a.BillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " Where a.YearMonth = '" & Format(TxtFromDate, "yyyy-mm-dd") & "' And a.ExchangeTypeID = 1 And a.Flag1 = 'N' and a.Rate is not null and a.TFlag in ('D') ANd a.BillingUnit = 'Q'"
tmp = tmp & wc
tmp = tmp & " Union All"
tmp = tmp & " Select Distinct a.YearMOnth,MS.StateName,ML.LocationName,ML.LocationID,MC.CommodityGroup,MC.CommodityGroupID,MB.BillingCycle,a.BillingUnit,a.Rate"
tmp = tmp & " ,Convert(varchar,ml.LocationID)+'-'+Convert(Varchar,mc.CommodityGroupID) 'Key'"
tmp = tmp & " from Txn_InvMOnGSLDEtail a"
tmp = tmp & " inner join mst_location ml on a.SM_Prefix = ml.SM_Prefix"
tmp = tmp & " inner join mst_commodity mcom On a.CommodityID = mcom.commodityID"
tmp = tmp & " inner join Mst_CommodityGRoup mc on mcom.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_billingcycle mb on a.BillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " Where a.YearMonth = '" & Format(TxtFromDate, "yyyy-mm-dd") & "' And a.ExchangeTypeID = 1 And a.Flag1 = 'N' and a.Rate is not null and a.TFlag in ('O','W') ANd a.BillingUnit = 'Q'"
tmp = tmp & wc
tmp = tmp & " )  A"
tmp = tmp & " Order By StateName,LocationName"

Call TmpTable


tmp = " Select b.LocationID,b.CommodityGroupID,mb.BillingCycle,b.FLag,a.RatePErBag, Convert(varchar,b.LocationID)+'-'+Convert(Varchar,b.CommodityGroupID) 'Key'"
tmp = tmp & " from Mst_StorageRateCardDepoDetail a"
tmp = tmp & " inner join Mst_StorageRateCard b on a.SRCID = b.SRCID"
tmp = tmp & " inner join mst_billingcycle mb on b.DepoBillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_CommodityGRoup mc on b.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_location ml on b.LocationID = ml.LocationID"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " WHere b.Flag = 'B' And b.DefaultRate = '1'"
tmp = tmp & wc
tmp = tmp & " Union All"
tmp = tmp & " Select b.LocationID,b.CommodityGroupID,mb.BillingCycle,b.FLag,a.RatePErBag , Convert(varchar,b.LocationID)+'-'+Convert(Varchar,b.CommodityGroupID) 'Key'"
tmp = tmp & " from Mst_StorageRateCardDetail a"
tmp = tmp & " inner join Mst_StorageRateCard b on a.SRCID = b.SRCID"
tmp = tmp & " inner join mst_billingcycle mb on b.BillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_CommodityGRoup mc on b.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_location ml on b.LocationID = ml.LocationID"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " WHere b.Flag = 'B' And b.DefaultRate = '1'"
tmp = tmp & wc

Call Tmp1Table

tmp = " Select a.LocationID,a.CommodityGroupID,mb.BillingCycle,a.FLag,a.RateOnQty, Convert(varchar,a.LocationID)+'-'+Convert(Varchar,a.CommodityGroupID) 'Key'"
tmp = tmp & " from Mst_StorageRateCard a"
tmp = tmp & " inner join mst_billingcycle mb on a.DepoBillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_CommodityGRoup mc on a.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_location ml on a.LocationID = ml.LocationID"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " WHere a.Flag = 'Q' And a.DefaultRate = '1'"
tmp = tmp & wc
tmp = tmp & " Union All"
tmp = tmp & " Select a.LocationID,a.CommodityGroupID,mb.BillingCycle,a.FLag,a.RateOnQty , Convert(varchar,a.LocationID)+'-'+Convert(Varchar,a.CommodityGroupID) 'Key'"
tmp = tmp & " from Mst_StorageRateCard a"
tmp = tmp & " inner join mst_billingcycle mb on a.BillingCycleID = mb.billingcycleid"
tmp = tmp & " inner join Mst_CommodityGRoup mc on a.CommodityGroupID = mc.CommodityGroupID"
tmp = tmp & " inner join mst_location ml on a.LocationID = ml.LocationID"
tmp = tmp & " inner join Mst_State ms on ml.StateID = ms.stateID"
tmp = tmp & " WHere a.Flag = 'Q' And a.DefaultRate = '1'"
tmp = tmp & wc

Call Tmp2Table

ProgressBar1.Max = TmpRs.RecordCount + 3

For i = 1 To TmpRs.RecordCount
    With MSHFlexGrid1
        .TextMatrix(i, 0) = IIf(IsNull(TmpRs!YearMonth), "", Format(TmpRs!YearMonth, "MMM-YYYY"))
        .TextMatrix(i, 1) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName) '"Location"
        .TextMatrix(i, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location"
        .TextMatrix(i, 3) = IIf(IsNull(TmpRs!COmmodityGroup), "", TmpRs!COmmodityGroup) '"Location"
        .TextMatrix(i, 4) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle) '"Location"
        .TextMatrix(i, 5) = IIf(IsNull(TmpRs!BillingUnit), "", TmpRs!BillingUnit) '"Location"
        .TextMatrix(i, 6) = IIf(IsNull(TmpRs!Rate), "", TmpRs!Rate) '"Location"
        .TextMatrix(i, 7) = IIf(IsNull(TmpRs!Key), "", TmpRs!Key) '"Location"
                        
        If LCase(.TextMatrix(i, 5)) = "b" Then
           .TextMatrix(i, 5) = "Bag"
        ElseIf LCase(.TextMatrix(i, 5)) = "q" Then
           .TextMatrix(i, 5) = "Quantity"
        End If
           
        If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
        TmpRs1.Find "Key = '" & TmpRs!Key & "'"
        If TmpRs1.EOF = False Then
           .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle) '"Location"
           .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag) '"Location"
           If LCase(.TextMatrix(i, 9)) = "b" Then
              .TextMatrix(i, 9) = "Bag"
           ElseIf LCase(.TextMatrix(i, 9)) = "q" Then
              .TextMatrix(i, 9) = "Quantity"
           End If
           .TextMatrix(i, 10) = IIf(IsNull(TmpRs1!RatePErBag), "", TmpRs1!RatePErBag) '"Location"
        Else
           If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
           TmpRs2.Find "Key = '" & TmpRs!Key & "'"
           If TmpRs2.EOF = False Then
              .TextMatrix(i, 8) = IIf(IsNull(TmpRs2!BillingCycle), "", TmpRs2!BillingCycle) '"Location"
              .TextMatrix(i, 9) = IIf(IsNull(TmpRs2!Flag), "", TmpRs2!Flag) '"Location"
              If LCase(.TextMatrix(i, 9)) = "b" Then
                 .TextMatrix(i, 9) = "Bag"
              ElseIf LCase(.TextMatrix(i, 9)) = "q" Then
                 .TextMatrix(i, 9) = "Quantity"
              End If
              .TextMatrix(i, 10) = IIf(IsNull(TmpRs2!RateOnQty), "", TmpRs2!RateOnQty) '"Location"
           End If
        End If
        .Rows = .Rows + 1
    End With
    ProgressBar1.Value = ProgressBar1.Value + 1
    TmpRs.MoveNext
Next

ProgressBar1.Value = ProgressBar1.Max
lblRecordCount.Caption = TmpRs.RecordCount
MsgBox "Done!!!"


End Sub



Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = CDate(Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01") 'GetFirstDateOfMonth(TxtFromDate.Value)
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
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11 '17
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.WordWrap = True

'YearMonth,LocationName,CommodityGroup,BillingCycle,BillingUnit,Rate,[Key]

MSHFlexGrid1.TextMatrix(0, 0) = "Year Month"
MSHFlexGrid1.TextMatrix(0, 1) = "NCML Area"
MSHFlexGrid1.TextMatrix(0, 2) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity Group"
MSHFlexGrid1.TextMatrix(0, 4) = "Billing Cycle"
MSHFlexGrid1.TextMatrix(0, 5) = "Billing Unit"
MSHFlexGrid1.TextMatrix(0, 6) = "Rate"
MSHFlexGrid1.TextMatrix(0, 7) = "Key Value"
MSHFlexGrid1.TextMatrix(0, 8) = "Current Billing Cycle"
MSHFlexGrid1.TextMatrix(0, 9) = "Current Billing Unit"
MSHFlexGrid1.TextMatrix(0, 10) = "Current Billing Rate"
     


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1700
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400

MSHFlexGrid1.RowHeight(0) = 800

End Sub
