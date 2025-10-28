VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_StockSummaryReport 
   Caption         =   "Stock Summary Report"
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
      Height          =   990
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   15015
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
         Left            =   2940
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   2700
      End
      Begin VB.ComboBox CmbCommodity 
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
         Left            =   8475
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   2760
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
         Left            =   5685
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2760
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
         Left            =   12510
         TabIndex        =   7
         Top             =   465
         Width           =   1215
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
         Left            =   11295
         TabIndex        =   6
         Top             =   465
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   90
         TabIndex        =   1
         Top             =   480
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   113770497
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1515
         TabIndex        =   2
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   113770497
         CurrentDate     =   39850
         MinDate         =   38353
      End
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
         Left            =   14190
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   720
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
         Left            =   13725
         TabIndex        =   8
         Top             =   465
         Width           =   1215
      End
      Begin VB.Label LblStateID 
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
         Height          =   300
         Left            =   3660
         TabIndex        =   18
         Top             =   165
         Width           =   1200
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   9315
         TabIndex        =   17
         Top             =   165
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "To  Date"
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
         Left            =   1770
         TabIndex        =   15
         Top             =   165
         Width           =   870
      End
      Begin VB.Label Label6 
         Caption         =   "From  Date"
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
         Left            =   248
         TabIndex        =   14
         Top             =   165
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
         Left            =   6683
         TabIndex        =   12
         Top             =   165
         Width           =   765
      End
      Begin VB.Label LblTotalRecord 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   13440
         TabIndex        =   11
         Top             =   165
         Width           =   60
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record Found :-"
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
         Left            =   11280
         TabIndex        =   10
         Top             =   165
         Width           =   2175
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   14205
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15266
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_StockSummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCommodityID, mExchangeTypeID, mRow As Variant

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmbLocationID.Text = ""
CmbCommodity.Text = ""
CmbStateID.Text = ""
CmbExchangeType.Text = ""
LblTotalRecord.Caption = ""
End Sub

Private Sub CmdGetReport_Click()
Dim WcMS, WcML, WcMC, WcME As Variant
WcMS = ""
WcML = ""
WcMC = ""
WcME = ""

Call Grid_Head
MsgBox "Wait till next message!!!"

If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox "To Date should be greater than From Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If CmbStateID.Text <> "" Then
   WcMS = " And MS.StateID=" & mStateID & ""
End If

If CmbLocationID.Text <> "" Then
   WcML = " And ML.LocationID=" & mLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      WcML = " And ML." & Gen_WcLocation
   End If
End If

If CmbCommodity.Text <> "" Then
   WcMC = " And MC.CommodityID=" & mCommodityID & ""
End If


tmp = " Select TempTable.StateName,TempTable.LocationName,TempTable.SM_Prefix,TempTable.CommodityID,TempTable.Commodity,Sum(TempTable.NoOfBags) NoOfBags,Sum(TempTable.Qty) Qty,TempTable.SM_Prefix,TempTable.CommodityID From"
tmp = tmp & " ("
tmp = tmp & "    ("
tmp = tmp & "    Select  MS.StateName,ML.LocationName,MC.Commodity,Sum(TCG.NoofBags)NoOfBags,Sum(TCG.GrossWeight)Qty ,TCD.SM_Prefix,TCD.CommodityID"
tmp = tmp & "    From Txn_CxTF_GSL TCG"
tmp = tmp & "    Inner Join Txn_CxTF_Details TCD on TCD.SM_Prefix=TCG.SM_Prefix and TCD.CxTF_DetailsID=TCG.CxTF_DetailsID and TCD.TxnType=TCG.TxnType"
tmp = tmp & "    Inner Join Mst_Commodity MC On MC.CommodityID=TCD.CommodityID"
tmp = tmp & "    Inner Join Mst_Location ML On ML.SM_PreFix=TCD.SM_Prefix"
tmp = tmp & "    Inner Join Mst_State MS On MS.StateID = ML.StateID"
tmp = tmp & "    Where TCG.TxnType='D' and "
If Gen_DBType = "MDB" Then
   tmp = tmp & " TCD.CxTFDate < #" & TxtFromDate.Value & "#"
Else
   tmp = tmp & " TCD.CxTFDate < '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
End If
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & "    Group By MS.StateName,ML.LocationName,MC.Commodity,TCD.SM_Prefix,TCD.CommodityID"
tmp = tmp & "    )"
tmp = tmp & "    Union All"
tmp = tmp & "    ("
tmp = tmp & "    Select MS.StateName,ML.LocationName,MC.Commodity,Sum(TS.SpillageNoOfBags)SpillageNoOfBags,Sum(TS.SpillageQuantity)Qty,TS.SM_PreFix,TS.CommodityID"
tmp = tmp & "    From  Txn_Spillage TS"
tmp = tmp & "    Inner Join Mst_Location ML On ML.SM_PreFix=TS.SM_PreFix"
tmp = tmp & "    Inner Join Mst_State MS On MS.StateID=ML.StateID"
tmp = tmp & "    Inner Join Mst_Commodity MC On MC.CommodityID=TS.CommodityID"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TS.SpillageDate < #" & TxtFromDate.Value & "#"
Else
   tmp = tmp & " Where TS.SpillageDate < '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
End If
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & "    Group By MS.StateName,ML.LocationName,MC.Commodity,TS.SM_PreFix,TS.CommodityID"
tmp = tmp & "    )"
tmp = tmp & "    Union All"
tmp = tmp & "    ("
tmp = tmp & "    Select MS.StateName,ML.LocationName,MC.Commodity,Sum(NoOfBags)NoOfBags,Sum(TG.Quantity) Qty,TA.SM_PreFix,TA.CommodityID"
tmp = tmp & "    from Txn_AdjTransaction TA"
tmp = tmp & "    Inner Join  Txn_GSLAdjTransaction TG On TA.TxnID=TG.TxnID"
tmp = tmp & "    Inner Join Mst_Location ML On ML.SM_PreFix=TA.SM_PreFix"
tmp = tmp & "    Inner Join Mst_State MS On MS.StateID=ML.StateID"
tmp = tmp & "    Inner Join Mst_Commodity MC On MC.CommodityID=TA.CommodityID"
tmp = tmp & "    Where TA.TxnType='AR' "
If Gen_DBType = "MDB" Then
   tmp = tmp & " And TA.TxnDate < #" & TxtFromDate.Value & "#"
Else
   tmp = tmp & " And TA.TxnDate < '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
End If
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & "    Group By MS.StateName,ML.LocationName,MC.Commodity,TA.SM_PreFix ,TA.CommodityID"
tmp = tmp & "    )"

tmp = tmp & "    Union All"
tmp = tmp & "    ("
tmp = tmp & "       Select StateName,LocationName,Commodity,Sum(NoOfBags)*(-1) As 'NoOfBags',Sum(Qty)*(-1) As 'Qty',SM_PreFix,CommodityID From"
tmp = tmp & "       ("
tmp = tmp & "          ("
tmp = tmp & "          Select MS.StateName,ML.LocationName,MC.Commodity,Sum(TCG.NoofBags)NoOfBags,Sum(TCG.GrossWeight)Qty,TCD.SM_Prefix,TCD.CommodityID"
tmp = tmp & "          From Txn_CxTF_GSL TCG"
tmp = tmp & "          Inner Join Txn_CxTF_Details TCD on TCD.SM_Prefix=TCG.SM_Prefix and TCD.CxTF_DetailsID=TCG.CxTF_DetailsID and TCD.TxnType=TCG.TxnType"
tmp = tmp & "          Inner Join Mst_Commodity MC On MC.CommodityID=TCD.CommodityID"
tmp = tmp & "          Inner Join Mst_Location ML On ML.SM_PreFix=TCD.SM_Prefix"
tmp = tmp & "          Inner Join Mst_State MS On MS.StateID = ML.StateID"
tmp = tmp & "          Where TCG.TxnType='W' "
If Gen_DBType = "MDB" Then
   tmp = tmp & " And TCD.CxTFDate <#" & TxtFromDate.Value & "#"
Else
   tmp = tmp & " And TCD.CxTFDate <'" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
End If
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & "          Group By MS.StateName,ML.LocationName,MC.Commodity,TCD.SM_Prefix,TCD.CommodityID"
tmp = tmp & "          )"
tmp = tmp & "          Union All"
tmp = tmp & "          ("
tmp = tmp & "          Select MS.StateName,ML.LocationName,MC.Commodity,Sum(NoOfBags)NoOfBags,Sum(TG.Quantity) Qty,TA.SM_PreFix,TA.CommodityID"
tmp = tmp & "          From Txn_AdjTransaction TA"
tmp = tmp & "          Inner Join  Txn_GSLAdjTransaction TG On TA.TxnID=TG.TxnID"
tmp = tmp & "          Inner Join Mst_Commodity MC On MC.CommodityID=TA.CommodityID"
tmp = tmp & "          Inner Join Mst_Location ML On ML.SM_PreFix=TA.SM_PreFix"
tmp = tmp & "          Inner Join Mst_State MS On MS.StateID=ML.StateID"
tmp = tmp & "          Where TA.TxnType<>'AR' "
If Gen_DBType = "MDB" Then
   tmp = tmp & " And TA.TxnDate < #" & TxtFromDate.Value & "#"
Else
   tmp = tmp & " And TA.TxnDate < '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'"
End If
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & "          Group By MS.StateName,ML.LocationName,MC.Commodity,TA.SM_PreFix,TA.CommodityID"
tmp = tmp & "          )"
tmp = tmp & "       )TempTable1"
tmp = tmp & "       Group By StateName,LocationName,Commodity,SM_PreFix,CommodityID"
tmp = tmp & "    )"

tmp = tmp & " )TempTable"

tmp = tmp & " Group by TempTable.StateName,TempTable.LocationName,TempTable.Commodity,TempTable.SM_Prefix,TempTable.CommodityID"
tmp = tmp & " Order By TempTable.StateName,TempTable.LocationName,TempTable.Commodity,TempTable.SM_Prefix,TempTable.CommodityID"

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 3

For I = 2 To TmpRs.RecordCount + 1
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!StateName
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!Locationname
    MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!CommodityID
    MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!Commodity
    MSHFlexGrid1.TextMatrix(I, 14) = TmpRs!SM_Prefix
    MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!NoofBags
    
    MSHFlexGrid1.TextMatrix(I, 5) = 0 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "D", "B")
    MSHFlexGrid1.TextMatrix(I, 6) = 0 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "W", "B")
    MSHFlexGrid1.TextMatrix(I, 8) = TmpRs!Qty
    MSHFlexGrid1.TextMatrix(I, 9) = 0 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "D", "Q") '"Deposited Qty(in MT)"
    MSHFlexGrid1.TextMatrix(I, 10) = 0 'GetDeposited_WithdrawalDetails(TmpRs!SM_Prefix, TmpRs!CommodityID, "W", "Q") '"Withdrawn Qty(in MT)"

    TmpRs.MoveNext
Next


tmp = " Select MS.StateName,ML.LocationName,MC.Commodity,TCD.SM_Prefix,TCD.CommodityID,TCG.TxnType,Sum(TCG.NoofBags)NoOfBags,Sum(TCG.GrossWeight)Qty "
tmp = tmp & " From Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD on TCD.SM_Prefix=TCG.SM_Prefix and TCD.CxTF_DetailsID=TCG.CxTF_DetailsID and TCD.TxnType=TCG.TxnType"
tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCD.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCD.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCD.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TCD.TxnType = 'D' And TCD.CxTFDate >= #" & TxtFromDate.Value & "# And TCD.CxTFDate<= #" & TxtToDate.Value & "#"
Else
   tmp = tmp & " Where TCD.TxnType = 'D' And TCD.CxTFDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TCD.CxTFDate<= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
End If

tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & " Group By TCD.SM_Prefix,TCD.CommodityID,TCG.TxnType,MS.StateName,ML.LocationName,MC.Commodity"
tmp = tmp & " Order By TCD.SM_Prefix,TCD.CommodityID,TCG.TxnType"

Call TmpTable


For I = 2 To TmpRs.RecordCount + 1
    mSM_Prefix = TmpRs!SM_Prefix
    mCommodityID = TmpRs!CommodityID
    mTFlag = TmpRs!TxnType
    mBags = TmpRs!NoofBags
    mQty = TmpRs!Qty
    If GetDataInGrid(mSM_Prefix, mCommodityID, mTFlag, mBags, mQty) = False Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(mRow, 0) = TmpRs!StateName
       MSHFlexGrid1.TextMatrix(mRow, 1) = TmpRs!Locationname
       MSHFlexGrid1.TextMatrix(mRow, 2) = TmpRs!CommodityID
       MSHFlexGrid1.TextMatrix(mRow, 3) = TmpRs!Commodity
       MSHFlexGrid1.TextMatrix(mRow, 14) = TmpRs!SM_Prefix
       
       If mTFlag = "D" Then
          MSHFlexGrid1.TextMatrix(mRow, 4) = 0
          MSHFlexGrid1.TextMatrix(mRow, 8) = 0
          MSHFlexGrid1.TextMatrix(mRow, 5) = TmpRs!NoofBags
          MSHFlexGrid1.TextMatrix(mRow, 9) = TmpRs!Qty
       End If
    End If
    TmpRs.MoveNext
Next


tmp = " Select MS.StateName,ML.LocationName,MC.Commodity,TCD.SM_Prefix,TCD.CommodityID,TCG.TxnType,Sum(TCG.NoofBags)NoOfBags,Sum(TCG.GrossWeight)Qty "
tmp = tmp & " From Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD on TCD.SM_Prefix=TCG.SM_Prefix and TCD.CxTF_DetailsID=TCG.CxTF_DetailsID and TCD.TxnType=TCG.TxnType"
tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCD.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCD.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCD.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TCD.TxnType = 'W' And TCD.CxTFDate >= #" & TxtFromDate.Value & "# And TCD.CxTFDate<= #" & TxtToDate.Value & "#"
Else
   tmp = tmp & " Where TCD.TxnType = 'W' And TCD.CxTFDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TCD.CxTFDate<= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
End If

tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & " Group By TCD.SM_Prefix,TCD.CommodityID,TCG.TxnType,MS.StateName,ML.LocationName,MC.Commodity"
tmp = tmp & " Order By TCD.SM_Prefix,TCD.CommodityID,TCG.TxnType"

Call TmpTable


For I = 2 To TmpRs.RecordCount + 1
    mSM_Prefix = TmpRs!SM_Prefix
    mCommodityID = TmpRs!CommodityID
    mTFlag = TmpRs!TxnType
    mBags = TmpRs!NoofBags
    mQty = TmpRs!Qty
    If GetDataInGrid(mSM_Prefix, mCommodityID, mTFlag, mBags, mQty) = False Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(mRow, 6) = TmpRs!NoofBags
       MSHFlexGrid1.TextMatrix(mRow, 10) = TmpRs!Qty
    End If
    TmpRs.MoveNext
Next


''' Query For Spillage

tmp = "Select MS.StateName,ML.LocationName,MC.Commodity,TS.SM_PreFix,TS.CommodityID,'D' As TxnType,Sum(TS.SpillageNoOfBags)NoOfBags,Sum(TS.SpillageQuantity)Qty"
tmp = tmp & " From  Txn_Spillage TS"
tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TS.SM_PreFix"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID=ML.StateID"
tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TS.CommodityID"
tmp = tmp & " Where TS.SpillageDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TS.SpillageDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & " Group By MC.CommodityID,TS.SM_PreFix,TS.CommodityID,MS.StateName,ML.LocationName,MC.Commodity"

Call TmpTable

For I = 2 To TmpRs.RecordCount + 1
    mSM_Prefix = TmpRs!SM_Prefix
    mCommodityID = TmpRs!CommodityID
    mTFlag = TmpRs!TxnType
    mBags = TmpRs!NoofBags
    mQty = TmpRs!Qty
    If GetDataInGrid(mSM_Prefix, mCommodityID, mTFlag, mBags, mQty) = False Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(mRow, 0) = TmpRs!StateName
       MSHFlexGrid1.TextMatrix(mRow, 1) = TmpRs!Locationname
       MSHFlexGrid1.TextMatrix(mRow, 2) = TmpRs!CommodityID
       MSHFlexGrid1.TextMatrix(mRow, 3) = TmpRs!Commodity
       MSHFlexGrid1.TextMatrix(mRow, 14) = TmpRs!SM_Prefix
    
        'TmpRs!NoofBags
       
       If mTFlag = "D" Then
          MSHFlexGrid1.TextMatrix(mRow, 4) = 0
          MSHFlexGrid1.TextMatrix(mRow, 8) = 0
          MSHFlexGrid1.TextMatrix(mRow, 5) = TmpRs!NoofBags
          MSHFlexGrid1.TextMatrix(mRow, 9) = TmpRs!Qty
       Else
'          MSHFlexGrid1.TextMatrix(mRow, 6) = TmpRs!NoofBags
'          MSHFlexGrid1.TextMatrix(mRow, 10) = TmpRs!Qty
       End If
    End If
    TmpRs.MoveNext
Next



'''Query for Adjustment
'
tmp = " Select MS.StateName,ML.LocationName,MC.Commodity,TA.SM_PreFix,TA.CommodityID,Sum(NoOfBags)NoOfBags,Sum(TG.Quantity) Qty,'D' As TxnType"
tmp = tmp & " from Txn_AdjTransaction TA"
tmp = tmp & " Inner Join  Txn_GSLAdjTransaction TG On TA.TxnID=TG.TxnID"
tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TA.SM_PreFix"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID=ML.StateID"
tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TA.CommodityID"
tmp = tmp & " Where TA.TxnType='AR'  "
tmp = tmp & " And TA.TxnDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TA.TxnDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & " Group By TA.SM_PreFix ,TA.CommodityID,MS.StateName,ML.LocationName,MC.Commodity"

Call TmpTable

For I = 2 To TmpRs.RecordCount + 1
    mSM_Prefix = TmpRs!SM_Prefix
    mCommodityID = TmpRs!CommodityID
    mTFlag = TmpRs!TxnType
    mBags = TmpRs!NoofBags
    mQty = TmpRs!Qty
    If GetDataInGrid(mSM_Prefix, mCommodityID, mTFlag, mBags, mQty) = False Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(mRow, 0) = TmpRs!StateName
       MSHFlexGrid1.TextMatrix(mRow, 1) = TmpRs!Locationname
       MSHFlexGrid1.TextMatrix(mRow, 2) = TmpRs!CommodityID
       MSHFlexGrid1.TextMatrix(mRow, 3) = TmpRs!Commodity
       If mTFlag = "D" Then
          MSHFlexGrid1.TextMatrix(mRow, 4) = 0
          MSHFlexGrid1.TextMatrix(mRow, 8) = 0
          MSHFlexGrid1.TextMatrix(mRow, 5) = TmpRs!NoofBags
          MSHFlexGrid1.TextMatrix(mRow, 9) = TmpRs!Qty
       End If
    End If
    TmpRs.MoveNext
Next


''''Not Equal to Adjustment (Will minus further)
'
tmp = "Select TA.SM_PreFix,TA.CommodityID,Sum(NoOfBags)NoOfBags,Sum(TG.Quantity) Qty,'W' As TxnType"
tmp = tmp & " From Txn_AdjTransaction TA"
tmp = tmp & " Inner Join  Txn_GSLAdjTransaction TG On TA.TxnID=TG.TxnID"
tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TA.CommodityID"
tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TA.SM_PreFix"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID=ML.StateID"
tmp = tmp & " Where TA.TxnType<>'AR'  "
tmp = tmp & " And TA.TxnDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TA.TxnDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
tmp = tmp & WcMS & WcML & WcMC
tmp = tmp & " Group By TA.SM_PreFix,TA.CommodityID"

Call TmpTable

For I = 2 To TmpRs.RecordCount + 1
    mSM_Prefix = TmpRs!SM_Prefix
    mCommodityID = TmpRs!CommodityID
    mTFlag = TmpRs!TxnType
    mBags = TmpRs!NoofBags
    mQty = TmpRs!Qty
    If GetDataInGrid(mSM_Prefix, mCommodityID, mTFlag, mBags, mQty) = False Then
    End If
    TmpRs.MoveNext
Next

tmp = "Select SM_Prefix,CommodityID,Sum(BalanceWeight)BalanceWeight"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " Where Status='E' And SpillageFlag='0'"
tmp = tmp & " Group By SM_Prefix,CommodityID"
tmp = tmp & " Order By SM_Prefix,CommodityID"

Call TmpTable
For I = 1 To TmpRs.RecordCount
    mSM_Prefix = TmpRs!SM_Prefix
    mCommodityID = TmpRs!CommodityID
    For x = 2 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(x, 0) = "" Then Exit For
        MSHFlexGrid1.TextMatrix(x, 12) = 0
        If LCase(MSHFlexGrid1.TextMatrix(x, 14)) = LCase(mSM_Prefix_) And MSHFlexGrid1.TextMatrix(x, 2) = mCommodityID_ Then
           'MSHFlexGrid1.TextMatrix(x, 5) = Val(MSHFlexGrid1.TextMatrix(x, 5)) + Val(mBags_)
           MSHFlexGrid1.TextMatrix(x, 12) = (TmpRs!BalanceWeight) * (-1)
           Exit For
        End If
    Next
    TmpRs.MoveNext
Next

For I = 2 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(I, 7) = (Val(MSHFlexGrid1.TextMatrix(I, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 5))) - Val(MSHFlexGrid1.TextMatrix(I, 6)) '(a + b) - c
    MSHFlexGrid1.TextMatrix(I, 7) = Format(MSHFlexGrid1.TextMatrix(I, 7), "########0.00")
    MSHFlexGrid1.TextMatrix(I, 11) = (Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 9))) - Val(MSHFlexGrid1.TextMatrix(I, 10)) '"Balance Qty(in MT)"
    MSHFlexGrid1.TextMatrix(I, 11) = Format(MSHFlexGrid1.TextMatrix(I, 11), "########0.000")
    MSHFlexGrid1.TextMatrix(I, 13) = Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 12))
    MSHFlexGrid1.TextMatrix(I, 13) = Format(MSHFlexGrid1.TextMatrix(I, 13), "########0.000")
Next



MSHFlexGrid1.TextMatrix(0, 4) = 0
MSHFlexGrid1.TextMatrix(0, 5) = 0
MSHFlexGrid1.TextMatrix(0, 6) = 0
MSHFlexGrid1.TextMatrix(0, 7) = 0
MSHFlexGrid1.TextMatrix(0, 8) = 0
MSHFlexGrid1.TextMatrix(0, 9) = 0
MSHFlexGrid1.TextMatrix(0, 10) = 0
MSHFlexGrid1.TextMatrix(0, 11) = 0
MSHFlexGrid1.TextMatrix(0, 12) = 0
MSHFlexGrid1.TextMatrix(0, 13) = 0

For I = 1 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
    MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
    MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
    MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
    MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 9))
    MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 10))
    MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 11))
    MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 12))
    MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(I, 13))
Next

LblTotalRecord.Caption = TmpRs.RecordCount
If MSHFlexGrid1.Rows <= 3 Then
   MsgBox "No Records Found !!!"
Else
   MsgBox "Done!!!"
End If

End Sub
Private Function GetDataInGrid(mSM_Prefix_ As Variant, mCommodityID_ As Variant, mTFlag_ As Variant, mBags_ As Variant, mQty_ As Variant) As Boolean
Dim Retval As Boolean
Retval = True
For x = 2 To MSHFlexGrid1.Rows - 1

    If MSHFlexGrid1.TextMatrix(x, 0) = "" Then
       Retval = False
       mRow = x
       Exit For
    End If
    If LCase(MSHFlexGrid1.TextMatrix(x, 14)) = LCase(mSM_Prefix_) And MSHFlexGrid1.TextMatrix(x, 2) = mCommodityID_ Then
       If LCase(mTFlag_) = "d" Then
          MSHFlexGrid1.TextMatrix(x, 5) = Val(MSHFlexGrid1.TextMatrix(x, 5)) + Val(mBags_)
          MSHFlexGrid1.TextMatrix(x, 9) = Val(MSHFlexGrid1.TextMatrix(x, 9)) + Val(mQty_)
       Else
          MSHFlexGrid1.TextMatrix(x, 6) = Val(MSHFlexGrid1.TextMatrix(x, 6)) + Val(mBags_)
          MSHFlexGrid1.TextMatrix(x, 10) = Val(MSHFlexGrid1.TextMatrix(x, 10)) + Val(mQty_)
       End If
       Exit For
    End If
Next

GetDataInGrid = Retval

End Function

Private Function GetDeposited_WithdrawalDetails(tSM_Prefix As Variant, tCommodityID As Variant, tFlag As Variant, tType As Variant)
Dim Retval As Variant
Dim TempBags, TempQty As Variant
Retval = 0
TmpRs1.Filter = "SM_Prefix='" & tSM_Prefix & "' And CommodityID=" & tCommodityID & "  And TxnType='" & tFlag & "'"
TmpRs3.Filter = "SM_Prefix='" & tSM_Prefix & "' And CommodityID=" & tCommodityID & ""
TmpRs4.Filter = "SM_Prefix='" & tSM_Prefix & "' And CommodityID=" & tCommodityID & ""
TmpRs5.Filter = "SM_Prefix='" & tSM_Prefix & "' And CommodityID=" & tCommodityID & ""
If TmpRs1.RecordCount > 0 Then
   If tType = "B" Then
      Retval = TmpRs1!NoofBags
   ElseIf tType = "Q" Then
      Retval = TmpRs1!Qty
   End If

End If

If tFlag = "D" Then
    If TmpRs3.RecordCount > 0 Then
       If tType = "B" Then
          Retval = Retval + TmpRs3!NoofBags
       ElseIf tType = "Q" Then
          Retval = Retval + TmpRs3!Qty
       End If
    End If
    If TmpRs4.RecordCount > 0 Then
       If tType = "B" Then
          Retval = Retval + TmpRs4!NoofBags
       ElseIf tType = "Q" Then
          Retval = Retval + TmpRs4!Qty
       End If
    End If
ElseIf tFlag = "W" Then
    If TmpRs5.RecordCount > 0 Then
       If tType = "B" Then
          Retval = Retval + TmpRs5!NoofBags
       ElseIf tType = "Q" Then
          Retval = Retval + TmpRs5!Qty
       End If
    End If
End If
TmpRs1.Filter = ""
TmpRs3.Filter = ""
TmpRs4.Filter = ""
TmpRs5.Filter = ""
GetDeposited_WithdrawalDetails = Retval
End Function

Private Sub Form_Load()
Call Grid_Head
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
TxtFromDate.Value = Date
TxtToDate.Value = Date
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2

MSHFlexGrid1.Cols = 15
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.AllowUserResizing = flexResizeColumns
MSHFlexGrid1.TextMatrix(1, 0) = "State Name"
MSHFlexGrid1.TextMatrix(1, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(1, 2) = "CommodityID"
MSHFlexGrid1.TextMatrix(1, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(1, 4) = "Opening Balance For Bags"
MSHFlexGrid1.TextMatrix(1, 5) = "Deposited Bags"
MSHFlexGrid1.TextMatrix(1, 6) = "Withdrawn Bags"
MSHFlexGrid1.TextMatrix(1, 7) = "Balance Bags"
MSHFlexGrid1.TextMatrix(1, 8) = "Opening Balance For Quantity"
MSHFlexGrid1.TextMatrix(1, 9) = "Deposited Quantity (in MT)"
MSHFlexGrid1.TextMatrix(1, 10) = "Withdrawn Quantity (in MT)"
MSHFlexGrid1.TextMatrix(1, 11) = "Balance Quantity (in MT)"
MSHFlexGrid1.TextMatrix(1, 12) = "Weight Gain/Loss"
MSHFlexGrid1.TextMatrix(1, 13) = "Net Balance"
MSHFlexGrid1.TextMatrix(1, 14) = "SM_Prefix"



MSHFlexGrid1.ColWidth(0) = 2000
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 0
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 1700
MSHFlexGrid1.ColWidth(5) = 1700
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1700
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 1500
MSHFlexGrid1.ColWidth(10) = 1500
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 0

MSHFlexGrid1.RowHeight(1) = 800
MSHFlexGrid1.Font.Size = 10
End Sub
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
'----
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
    CmbLocationID.AddItem RsMst_Location!Locationname
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
Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodity.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID

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

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_StockSummaryReport.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_StockSummaryReport.xls")
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
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

