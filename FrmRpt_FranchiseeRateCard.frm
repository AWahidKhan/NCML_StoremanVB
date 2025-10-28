VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_FranchiseeRateCard 
   Caption         =   "Report for Franchisee Rate Card"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   15000
      Begin VB.OptionButton OptInRateCard 
         Caption         =   "In Rate Card"
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
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton OptNotInRateCard 
         Caption         =   "Not in Rate Card"
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
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   1890
      End
      Begin VB.ComboBox CmbLessorName 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   795
         Width           =   6255
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmRpt_FranchiseeRateCard.frx":0000
         Left            =   9480
         List            =   "FrmRpt_FranchiseeRateCard.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
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
         Left            =   12840
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2055
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
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3015
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
         Left            =   12885
         TabIndex        =   6
         Top             =   795
         Width           =   1000
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
         Left            =   11880
         TabIndex        =   5
         Top             =   795
         Width           =   1000
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
         Left            =   13890
         TabIndex        =   7
         Top             =   795
         Width           =   1000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Godowns"
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
         Left            =   150
         TabIndex        =   19
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Left            =   8280
         TabIndex        =   16
         Top             =   855
         Width           =   2415
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
         Left            =   10680
         TabIndex        =   15
         Top             =   855
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Commodity"
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
         Left            =   11640
         TabIndex        =   13
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label3 
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
         Left            =   3960
         TabIndex        =   12
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label1 
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
         Left            =   7920
         TabIndex        =   11
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Franchisee Name "
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
         Left            =   150
         TabIndex        =   10
         Top             =   855
         Width           =   1650
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
      Height          =   8340
      Left            =   120
      TabIndex        =   0
      Top             =   1215
      Width           =   15000
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   8
         Top             =   255
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   13361
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   7920
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_FranchiseeRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityGroupId, mLocationID, mVendorID, mExchangeTypeID As Variant

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

'----
Private Sub CmbCommodityGroupID_LostFocus()
If CmbCommodityGroupID.Text = "" Then
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

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then Exit Sub

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
    MsgBox "Invalid selection "
    CmbExchangeTypeID.SetFocus
    Exit Sub
End If

mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmbLessorName_GotFocus()
tmp = CmbLessorName.Text
mLessorID = ""
Call TableMst_Vendor
CmbLessorName.Clear

If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Vendor.RecordCount
    CmbLessorName.AddItem RsMst_Vendor!VendorName
    RsMst_Vendor.MoveNext
Next
CmbLessorName.Text = tmp
End Sub

Private Sub CmbLessorName_LostFocus()
If CmbLessorName.Text = "" Then
   Exit Sub
End If
RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbLessorName.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid selection "
   CmbLessorName.SetFocus
   Exit Sub
End If
mVendorID = RsMst_Vendor!VendorID
End Sub

Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation & "")
End If

CmbLocationID.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmdClear_Click()
OptInRateCard.Value = True
CmbCommodityGroupID.Text = ""
CmbExchangeTypeID.Text = ""
CmbLessorName.Text = ""
CmbLocationID.Text = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
Call Grid_Head
LblTotalRecords.Caption = ""

End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 16
       
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "Franchisee Name"
    .TextMatrix(0, 2) = "Commodity"
    .TextMatrix(0, 3) = "Exchange Type"
    .TextMatrix(0, 4) = "CMP Name"
    .TextMatrix(0, 5) = "Godown No"
    .TextMatrix(0, 6) = "Godown Address"
    .TextMatrix(0, 7) = "Billing Cycle"
    .TextMatrix(0, 8) = "Total Days"
    .TextMatrix(0, 9) = "Billing Unit"
    .TextMatrix(0, 10) = "Rate on Units"
    .TextMatrix(0, 11) = "Bag Size From"
    .TextMatrix(0, 12) = "Bag Size To"
    .TextMatrix(0, 13) = "Rate Per Bag per Cycle"
    .TextMatrix(0, 14) = "Godown Start Date"
    .TextMatrix(0, 15) = "Godown Close Date"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 2400
    .ColWidth(2) = 1700
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 900
    .ColWidth(6) = 2800
    .ColWidth(7) = 1000
    .ColWidth(8) = 700
    .ColWidth(9) = 1200
    .ColWidth(10) = 900
    .ColWidth(11) = 900
    .ColWidth(12) = 900
    .ColWidth(13) = 1400
    .ColWidth(14) = 1500
    .ColWidth(15) = 1500
    
    .RowHeight(0) = 700

End With
End Sub

Private Sub CmdExcel_Click()

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_FranchiseeRateCard.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_FranchiseeRateCard.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Storage Rate Card"
oWS.ClearArrows
mRow = 1

tmp = "a" & mRow & ":" & "p" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, C + 1) = "Franchisee Rate Card"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, C + 1).WrapText = True

For i = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        If i = 0 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, C + 1).Font.Bold = True
           oWS.Cells(mRow, C + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, C + 1).WrapText = True
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGo_Click()
Dim wc As String
wc = ""

ProgressBar1.Value = 0
Call Grid_Head
MsgBox "Wait till next message!!!"

If CmbLocationID.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And ML.LocationID = " & mLocationID
   Else
      wc = " Where ML.LocationID = " & mLocationID
   End If
Else
   If Gen_WcLocation <> "" Then
      If wc <> "" Then
         wc = wc & " And ML." & Gen_WcLocation
      Else
         wc = " Where ML." & Gen_WcLocation
      End If
   End If
End If

If CmbCommodityGroupID.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And MCG.CommodityGroupID = " & mCommodityGroupId
   Else
      wc = " Where MCG.CommodityGroupID = " & mCommodityGroupId
   End If
   'wc = GenWc(wc, mCommodityGroupId.Text, "MCG.CommodityGroupID", "N", "I")
End If

If CmbExchangeTypeID.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And MET.ExchangeTypeID = " & mExchangeTypeID
   Else
      wc = " Where MET.ExchangeTypeID = " & mExchangeTypeID
   End If
   'wc = GenWc(wc, mExchangeTypeID, "MET.ExchangeTypeID", "N", "I")
End If

If OptInRateCard.Value = True Then
    If CmbLessorName.Text <> "" Then
       If wc <> "" Then
          wc = wc & " And MV.VendorID = " & mVendorID
       Else
          wc = " Where MV.VendorID = " & mVendorID
       End If
       'wc = GenWc(wc, mVendorID, "MV.VendorID", "N", "I")
    End If
    
    tmp = " select ML.LocationName, MV.VendorName, MCG.CommodityGroup, MET.ExchangeType,"
    tmp = tmp & " MCMP.CMPName, MG.GodownNo, MG.Address, MBC.BillingCycle, MBC.NoofDays, MFRG.BillingUnit,"
    tmp = tmp & " MFRG.RateOnQty , MFRGB.BagSizeFrom, MFRGB.BagSizeTo, MFRGB.RatePerBag, MG.StartDate, MG.CloseDate"
    tmp = tmp & " from Mst_FF_RateCard MFR"
    tmp = tmp & " Inner Join Mst_FF_RateCard_Godown MFRG on MFRG.FF_RateCardID = MFR.FF_RateCardID"
    tmp = tmp & " Left Join Mst_FF_RateCard_Godown_Bag MFRGB on  MFRGB.FF_RateCard_GodownID = MFRG.FF_RateCard_GodownID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MFR.LocationID"
    tmp = tmp & " Inner Join Mst_CommodityGroup MCG on MCG.CommodityGroupID = MFR.CommodityGroupID"
    tmp = tmp & " Inner Join Mst_Vendor MV on MV.VendorID = MFR.VendorID"
    tmp = tmp & " Inner Join Mst_ExchangeType MET on MET.ExchangeTypeID = MFR.ExchangeTypeID"
    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MFRG.CMPID"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MFRG.GodownID"
    tmp = tmp & " Inner Join Mst_BillingCycle MBC on MBC.BillingCycleID = MFRG.BillingCycleID"
    tmp = tmp & " " & wc & " Order by ML.LocationName, MV.VendorName, MCG.CommodityGroup "
ElseIf OptNotInRateCard.Value = True Then
'Commented on 20 March 2018 by parag for rectification of issue mail from Som Dat (IVG)

'    tmp = " select distinct ml.locationname, mcmp.cmpname, mg.godownno, mg.address, mcg.commoditygroup, met.exchangetype, MG.StartDate, MG.CloseDate from mst_gsl gsl"
'    tmp = tmp & " inner join mst_godown mg on mg.godownid = gsl.godownid"
'    tmp = tmp & " inner join mst_cmp mcmp on mcmp.cmpid = gsl.cmpid"
'    tmp = tmp & " inner join mst_location ml on ml.sm_prefix = gsl.sm_prefix"
'    tmp = tmp & " inner join mst_commodity mc on mc.commodityid = gsl.commodityid"
'    tmp = tmp & " inner join mst_commoditygroup mcg on mcg.commoditygroupid = mc.commoditygroupid"
'    tmp = tmp & " inner join mst_exchangetype met on met.exchangetypeid = gsl.exchangetypeid"
'    If wc <> "" Then
'        tmp = tmp & " " & wc & " and "
'    Else
'        tmp = tmp & " where "
'    End If
'    tmp = tmp & " convert(varchar, mg.godownid)+'_'+convert(varchar, mc.commoditygroupid)"
'    tmp = tmp & " not in (select convert(varchar, mfg.godownid)+'_'+convert(varchar, mf.commoditygroupid) from mst_ff_ratecard_godown mfg"
'    tmp = tmp & " inner join mst_ff_ratecard mf on mf.ff_ratecardid = mfg.ff_ratecardid)"
'    tmp = tmp & " and mg.unittypeid in (2,6) "
'    tmp = tmp & " order by ml.locationname, mcmp.cmpname, mg.godownno, mcg.commoditygroup"



    tmp = " select distinct ml.locationname, mcmp.cmpname, mg.godownno, mg.address, mcg.commoditygroup, met.exchangetype, MG.StartDate, MG.CloseDate from mst_gsl gsl"
    tmp = tmp & " inner join mst_godown mg on mg.godownid = gsl.godownid"
    tmp = tmp & " inner join mst_cmp mcmp on mcmp.cmpid = gsl.cmpid"
    tmp = tmp & " inner join mst_location ml on ml.sm_prefix = gsl.sm_prefix"
    tmp = tmp & " inner join mst_commodity mc on mc.commodityid = gsl.commodityid"
    tmp = tmp & " inner join mst_commoditygroup mcg on mcg.commoditygroupid = mc.commoditygroupid"
    tmp = tmp & " inner join mst_exchangetype met on met.exchangetypeid = gsl.exchangetypeid"
    If wc <> "" Then
        tmp = tmp & " " & wc & " and "
    Else
        tmp = tmp & " where "
    End If
    tmp = tmp & " convert(varchar, mg.godownid)+'_'+convert(varchar, mc.commoditygroupid)"
    tmp = tmp & " not in (select convert(varchar, mfg.godownid)+'_'+convert(varchar, mf.commoditygroupid) from mst_ff_ratecard_godown mfg"
    tmp = tmp & " inner join mst_ff_ratecard mf on mf.ff_ratecardid = mfg.ff_ratecardid and mfg.BillingUnit <> 'F')"
    tmp = tmp & " and mg.unittypeid in (2,6) And MG.GodownID not in ("
    tmp = tmp & " Select distinct MG.GodownID"
    tmp = tmp & " from mst_gsl gsl"
    tmp = tmp & " inner join mst_godown mg on mg.godownid = gsl.godownid"
    tmp = tmp & " inner join mst_cmp mcmp on mcmp.cmpid = gsl.cmpid"
    tmp = tmp & " inner join mst_location ml on ml.sm_prefix = gsl.sm_prefix"
    tmp = tmp & " inner join mst_exchangetype met on met.exchangetypeid = gsl.exchangetypeid"
    If wc <> "" Then
        tmp = tmp & " " & wc & " and "
    Else
        tmp = tmp & " where "
    End If
    tmp = tmp & " mg.godownid in (select godownid"
    tmp = tmp & " from mst_ff_ratecard_godown Where BillingUnit = 'F')"
    tmp = tmp & " and mg.unittypeid in (2,6)) "
    tmp = tmp & " order by ml.locationname, mcmp.cmpname, mg.godownno, mcg.commoditygroup"




End If

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 2
LblTotalRecords.Caption = TmpRs.RecordCount

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   If OptInRateCard.Value = True Then
      For i = 1 To TmpRs.RecordCount
         MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(TmpRs!VendorName), "", TmpRs!VendorName)
         MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(TmpRs!COmmodityGroup), "", TmpRs!COmmodityGroup)
         MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
         MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
         MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)
         MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(TmpRs!NoOfDays), "", TmpRs!NoOfDays)
         MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(TmpRs!BillingUnit), "", TmpRs!BillingUnit)
         If LCase(MSHFlexGrid1.TextMatrix(i, 9)) = "b" Then
           MSHFlexGrid1.TextMatrix(i, 9) = "Bags"
         ElseIf LCase(MSHFlexGrid1.TextMatrix(i, 9)) = "q" Then
            MSHFlexGrid1.TextMatrix(i, 9) = "Quantity"
         ElseIf LCase(MSHFlexGrid1.TextMatrix(i, 9)) = "f" Then
            MSHFlexGrid1.TextMatrix(i, 9) = "Fixed"
         ElseIf LCase(MSHFlexGrid1.TextMatrix(i, 9)) = "p" Then
            MSHFlexGrid1.TextMatrix(i, 9) = "Percentage"
         End If
         MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(TmpRs!RateOnQty), "", TmpRs!RateOnQty)
         MSHFlexGrid1.TextMatrix(i, 11) = IIf(IsNull(TmpRs!BagSizeFrom), "", TmpRs!BagSizeFrom)
         MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs!BagSizeTo), "", TmpRs!BagSizeTo)
         MSHFlexGrid1.TextMatrix(i, 13) = IIf(IsNull(TmpRs!RatePerBag), "", TmpRs!RatePerBag)
         MSHFlexGrid1.TextMatrix(i, 14) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
         MSHFlexGrid1.TextMatrix(i, 15) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   ElseIf OptNotInRateCard.Value = True Then
      For i = 1 To TmpRs.RecordCount
         MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(TmpRs!COmmodityGroup), "", TmpRs!COmmodityGroup)
         MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
         MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
         MSHFlexGrid1.TextMatrix(i, 14) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
         MSHFlexGrid1.TextMatrix(i, 15) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End If
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
   CmdExcel.Enabled = True
Else
   MsgBox "No record found!!!"
   CmdExcel.Enabled = False
End If

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

Private Sub OptInRateCard_Click()
CmbLessorName.Enabled = True
End Sub

Private Sub OptNotInRateCard_Click()
CmbLessorName.Enabled = False
CmbLessorName.Text = ""
End Sub
