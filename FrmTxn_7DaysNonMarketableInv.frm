VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_7DaysNonMarketableInv 
   Caption         =   "7 Days Non Marketable Invoice Generation"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   9525
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14970
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
         Left            =   7140
         TabIndex        =   4
         Top             =   405
         Width           =   1845
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
         Left            =   150
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   420
         Width           =   3210
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Generate Invoice"
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
         Left            =   5295
         TabIndex        =   3
         Top             =   405
         Width           =   1845
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
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
         Left            =   3450
         TabIndex        =   2
         Top             =   405
         Width           =   1845
      End
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   885
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8175
         Left            =   120
         TabIndex        =   6
         Top             =   1215
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   14420
         _Version        =   393216
         FixedCols       =   0
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
      Begin VB.Label Label24 
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
         Left            =   9345
         TabIndex        =   9
         Top             =   465
         Width           =   2535
      End
      Begin VB.Label LblTotalRecord 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12075
         TabIndex        =   8
         Top             =   510
         Width           =   75
      End
      Begin VB.Label Label34 
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
         Height          =   225
         Left            =   1320
         TabIndex        =   7
         Top             =   165
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmTxn_7DaysNonMarketableInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSM_Prefix, mLocationID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 17

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Selection"
MSHFlexGrid1.TextMatrix(0, 1) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 2) = "GSL ID."
MSHFlexGrid1.TextMatrix(0, 3) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 4) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 5) = "CMP"
MSHFlexGrid1.TextMatrix(0, 6) = "Godown Address"
MSHFlexGrid1.TextMatrix(0, 7) = "GSL No."
MSHFlexGrid1.TextMatrix(0, 8) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 9) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 10) = "Weight in MT"
MSHFlexGrid1.TextMatrix(0, 11) = "ClientIDRow"
MSHFlexGrid1.TextMatrix(0, 12) = "ClientID"
MSHFlexGrid1.TextMatrix(0, 13) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 14) = "KR No"
MSHFlexGrid1.TextMatrix(0, 15) = "KR Date"
MSHFlexGrid1.TextMatrix(0, 16) = "Rate"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 0
MSHFlexGrid1.ColWidth(2) = 0
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1800
MSHFlexGrid1.ColWidth(6) = 2200
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1800
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 0
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 2000
MSHFlexGrid1.ColWidth(14) = 1200
MSHFlexGrid1.ColWidth(15) = 1400
MSHFlexGrid1.ColWidth(16) = 1200

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True

End Sub

Private Sub ChkCheckAll_Click()
If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "7DaysNonMarketable_GSL.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "7DaysNonMarketable_GSL.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdSave_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Please select Location!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If MSHFlexGrid1.Rows < 3 Then
   MsgBox " No Location populated in Grid  !!!"
   Exit Sub
End If

If ChkSelection = False Then
   MsgBox " No GSL selected from Grid  !!!"
   Exit Sub
End If

For Ln = 1 To MSHFlexGrid1.Rows - 1

    If MSHFlexGrid1.TextMatrix(Ln, 0) = strChecked Then

      Call TableTxn_CDTFBOE(" Where SM_Prefix='" & mSM_Prefix & "' And CDTFNo=" & Val(MSHFlexGrid1.TextMatrix(Ln, 3)) & " And InvoiceDate='" & Date & "'")
      
      Call TableTxn_CDTFBOE_GSL(" Where SM_Prefix='" & mSM_Prefix & "'")

       tmp = "Select TCD.CxTFDate,TCD.CxTFNo From Txn_CxTF_Details TCD"
       tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType "
       tmp = tmp & " Where TCD.TxnType = 'D' and TCD.ExchangeTypeID = 2 And TCG.GSLID=" & MSHFlexGrid1.TextMatrix(Ln, 2) & " And TCD.SM_Prefix='" & mSM_Prefix & "' And TCD.CxTFNo>" & Val(MSHFlexGrid1.TextMatrix(Ln, 3))
       
       Call Tmp5Table
       
       If TmpRs5.RecordCount <= 0 Then
          mRate = MSHFlexGrid1.TextMatrix(Ln, 16)
          tmp = "Select DepositeWeight,DepositeBags,CDTFDate,LCDTF_Date,KRDate From Mst_GSL Where GSLID=" & Val(MSHFlexGrid1.TextMatrix(Ln, 2)) & " And SM_Prefix='" & mSM_Prefix & "'"  ''CDTFDate<>LCDTF_Date And
          Call Tmp4Table
          If TmpRs4.RecordCount > 0 Then
             If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
             RsTxn_CDTFBOE.Find "ClientIDRow=" & Val(MSHFlexGrid1.TextMatrix(Ln, 11))
             If RsTxn_CDTFBOE.EOF = True Then
                Call AddCDTFBOE(Val(MSHFlexGrid1.TextMatrix(Ln, 11)), Ln)
             End If
             mTotDays = DateDiff("d", TmpRs4!CDTFDate, DateAdd("d", 7, TmpRs4!LCDTF_Date)) '+ 1
             mTotDays1 = DateDiff("d", TmpRs4!CDTFDate, TmpRs4!KRDate) + 1
             If mTotDays > mTotDays1 Then
                Call SaveCDTFBOEGSL(Val(MSHFlexGrid1.TextMatrix(Ln, 2)), TmpRs4!DepositeBags, TmpRs4!DepositeWeight, mTotDays, mRate, Val(MSHFlexGrid1.TextMatrix(Ln, 3)))
             Else
                Call SaveCDTFBOEGSL(Val(MSHFlexGrid1.TextMatrix(Ln, 2)), TmpRs4!DepositeBags, TmpRs4!DepositeWeight, mTotDays1, mRate, Val(MSHFlexGrid1.TextMatrix(Ln, 3)))
             End If
          End If
       End If
    End If
Next
MsgBox "Done!!"
Call Grid_Head
CmbLocationID.Text = ""
LblTotalRecord.Caption = "-"
End Sub
Private Sub AddCDTFBOE(mClientID_ As Variant, mROw_ As Variant)

RsTxn_CDTFBOE.AddNew
RsTxn_CDTFBOE!BOEID = getCDTFBOE_Max
RsTxn_CDTFBOE!LocationID = mLocationID
RsTxn_CDTFBOE!SM_Prefix = mSM_Prefix
RsTxn_CDTFBOE!CdtfNo = MSHFlexGrid1.TextMatrix(mROw_, 3)
RsTxn_CDTFBOE!CDTFDate = CDate(MSHFlexGrid1.TextMatrix(mROw_, 4))
RsTxn_CDTFBOE!ClientIDRow = mClientID_
RsTxn_CDTFBOE!InvoiceDate = Date
RsTxn_CDTFBOE!BillingMonth = CDate(Year(CDate(MSHFlexGrid1.TextMatrix(mROw_, 15))) & "-" & Month(CDate(MSHFlexGrid1.TextMatrix(mROw_, 15))) & "-01")
RsTxn_CDTFBOE!Discount = 0
RsTxn_CDTFBOE!Premium = 0
RsTxn_CDTFBOE!BatchID = 0
RsTxn_CDTFBOE!RBatchID = 0
RsTxn_CDTFBOE!Flag = "P"
RsTxn_CDTFBOE!G_UID = GetGUID
RsTxn_CDTFBOE!CreatedDate = Now
RsTxn_CDTFBOE!CreatedBy = Gen_UserLoginID
RsTxn_CDTFBOE!TxnType = "D"
RsTxn_CDTFBOE.Update

End Sub
Private Sub SaveCDTFBOEGSL(mGSLID_ As Variant, mNoofBags_ As Variant, mDematWeight_ As Variant, mNoOfDays_ As Variant, mRate_ As Variant, mCDTFNo_ As Variant)

RsTxn_CDTFBOE_GSL.AddNew
RsTxn_CDTFBOE_GSL!BOEGSLID = getCDTFBOE_GSL_Max
RsTxn_CDTFBOE_GSL!BOEID = RsTxn_CDTFBOE!BOEID
RsTxn_CDTFBOE_GSL!SM_Prefix = mSM_Prefix
RsTxn_CDTFBOE_GSL!GSLID = mGSLID_
RsTxn_CDTFBOE_GSL!CDTFBags = mNoofBags_
RsTxn_CDTFBOE_GSL!CDTFQty = mDematWeight_
RsTxn_CDTFBOE_GSL!NoOfDays = mNoOfDays_
RsTxn_CDTFBOE_GSL!Rate = mRate_ 'TmpRs2!Rate
RsTxn_CDTFBOE_GSL!BillingUnit = "Q" 'TmpRs2!Unit
RsTxn_CDTFBOE_GSL!BillingCycleID = 1 ''TmpRs2!BillingCycleID
RsTxn_CDTFBOE_GSL!Amount = Format(mNoOfDays_ * mDematWeight_ * mRate_, "#####0.00")
RsTxn_CDTFBOE_GSL.Update

RsTxn_CDTFBOE!Amount = IIf(IsNull(RsTxn_CDTFBOE!Amount), 0, RsTxn_CDTFBOE!Amount) + RsTxn_CDTFBOE_GSL!Amount
RsTxn_CDTFBOE.Update

tmp = "Update Txn_CxTF_GSL Set BillFlag='I' Where SM_Prefix='" & mSM_Prefix & "' And CxTFNo=" & mCDTFNo_ & " And TxnType='D' And GSLID=" & mGSLID_ & " And BillFlag is null"
Call Tmp5Table

End Sub

Private Sub CmdSearch_Click()
If CmbLocationID.Text = "" Then
   MsgBox "Please select Location!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = "select CMP.LocationID,CMp.CMPname,god.GodownNo 'GodownNo',gsl.StackNo,gsl.LotNo,TCD.CxtfNo,TCD.CxtfDate,MC.Commodity,"
tmp = tmp & " sum(tgsl.NoofBags)NoofBags,sum(tgsl.GrossWeight)GrossWeight,sum(tgsl.AdjNClientWeight)AdjNClientWeight,"
tmp = tmp & " sum(tgsl.AdjNCMSLWeight)AdjNCMSLWeight,sum(tgsl.GunnyWeight)GunnyWeight,sum(tgsl.StdDeduction)StdDeduction,"
tmp = tmp & " sum(tgsl.DematWeight)DematWeight,c.ClientName + ' ' + d.DepositoryName ClientName,tgsl.Remarks,tgsl.GSLID,"
tmp = tmp & " god.CloseDate , god.GodownID,MC.CommodityID, tgsl.ClientIDRow, gsl.ISINID, gsl.FGSLID, god.Address,gsl.KRNo,GSL.KRDate,C.ClientID"
tmp = tmp & " From Txn_CxTF_GSL tgsl"
tmp = tmp & " inner join Txn_CxTF_details TCD on TCD.CxtfNo= tgsl.CxTfNo and TCD.SM_Prefix = tgsl.SM_Prefix And TCD.TxnType=tgsl.TxnType"
tmp = tmp & " inner join mst_gsl gsl on gsl.GSLID = tgsl.GSLID and gsl.SM_Prefix = tgsl.SM_Prefix"
tmp = tmp & " inner join Mst_Godown god on gsl.GodownID = god.GodownID"
tmp = tmp & " inner join Mst_CMP CMP on CMP.CMPID = god.CMPID"
tmp = tmp & " inner join mst_client c on tgsl.ClientIDRow = c.ClientIDRow"
tmp = tmp & " inner join mst_depository d on d.DepositoryID = c.DepositoryID  and d.ExchangeTypeID =2"
tmp = tmp & " Inner join Mst_Commodity MC On MC.CommodityID=gsl.CommodityID"
tmp = tmp & " Inner join Mst_CommodityDetail CMD On CMD.CommodityID=gsl.CommodityID AND CMD.ExchangeTypeID=GSL.ExchangeTypeID"
tmp = tmp & " Where tgsl.TxnType = 'D' And gsl.ExchangeTypeID=2 And tgsl.SM_Prefix = '" & mSM_Prefix & "' And isnull(gsl.KRNo,0)>0 And gsl.KRDate is not null"
tmp = tmp & " AND (gsl.DepositeWeight<(CMD.LotSize - CMD.LotSize*CMD.Variation/100))"
tmp = tmp & " And gsl.GSLID not in (Select Distinct TCG.GSLID  from  Txn_CDTFBOE_GSL  TCG Inner join Txn_CDTFBOE TC on TCG.BOEID=TC.BOEID Where TCG.SM_Prefix='" & mSM_Prefix & "' And TC.TxnType='D')"
tmp = tmp & " Group By CMP.LocationID,CMp.CMPname,god.GodownNo,MC.Commodity,gsl.StackNo,gsl.LotNo,c.ClientName ,d.DepositoryName,TCD.CxtfNo,TCD.CxtfDate,"
tmp = tmp & " tgsl.Remarks,tgsl.GSLID,god.CloseDate,god.GodownID,tgsl.ClientIDRow,gsl.ISINID,gsl.FGSLID,C.ClientID,"
tmp = tmp & " god.Address,gsl.KRNo,GSL.KRDate,MC.CommodityID"
tmp = tmp & " Order by TCD.CxtfNo,TCD.CxtfDate,tgsl.GSLID"

Call Tmp1Table

Call Grid_Head
If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs1!LocationID
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs1!GSLID
       MSHFlexGrid1.TextMatrix(I, 3) = TmpRs1!CxTFNO
       MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs1!CxTFDate, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 5) = TmpRs1!CMPName
       MSHFlexGrid1.TextMatrix(I, 6) = TmpRs1!Address
       MSHFlexGrid1.TextMatrix(I, 7) = TmpRs1!GodownNo & "~" & TmpRs1!StackNo & "~" & TmpRs1!LotNo
       MSHFlexGrid1.TextMatrix(I, 8) = TmpRs1!Commodity
       MSHFlexGrid1.TextMatrix(I, 9) = TmpRs1!NoofBags
       MSHFlexGrid1.TextMatrix(I, 10) = TmpRs1!DematWeight
       MSHFlexGrid1.TextMatrix(I, 11) = TmpRs1!ClientIDRow
       MSHFlexGrid1.TextMatrix(I, 12) = TmpRs1!ClientID
       MSHFlexGrid1.TextMatrix(I, 13) = TmpRs1!ClientName
       MSHFlexGrid1.TextMatrix(I, 14) = TmpRs1!KRNo
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs1!KRDate), "", Format(TmpRs1!KRDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 16) = GetStorageRate(TmpRs1!CommodityID, TmpRs1!GodownID)
       
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.Text = strUnChecked
         
       TmpRs1.MoveNext
       
   Next
   LblTotalRecord.Caption = TmpRs1.RecordCount
Else
  MsgBox " No Record found!!"
End If

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame0.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head

Frame0.Enabled = True
CmbLocationID.Text = ""
LblTotalRecord.Caption = "-"

End Sub


Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

Call TableMst_Location(" Where Collect7Days=1")

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
    
End Sub

'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 0
        If .TextMatrix(.row, .Col) = strUnChecked Then
           If .TextMatrix(.row, 3) <> "" Then
                 mPINo = .TextMatrix(.row, 3)
                 .TextMatrix(.row, .Col) = strChecked
                 For x = 1 To MSHFlexGrid1.Rows - 1
                     If mPINo = .TextMatrix(x, 3) Then
                        .TextMatrix(x, .Col) = strChecked
                     End If
                 Next
           End If
        Else
           .TextMatrix(.row, .Col) = strUnChecked
           mPINo = .TextMatrix(.row, 3)
           For x = 1 To MSHFlexGrid1.Rows - 1
               If mPINo = .TextMatrix(x, 3) Then
                  .TextMatrix(x, .Col) = strUnChecked
               End If
           Next
        End If
     End If
End With


End Sub

Private Function GetStorageRate(mCommodityID_ As Variant, mGodownID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0

tmp = "Select MNRC.GodownID,MNRC.Rate from Mst_NCDEXStorageRateCardDetail MNRC"
tmp = tmp & " inner Join Mst_NCDEXStorageRateCard MNR on MNRC.NSRCID=MNR.NSRCID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MNR.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Commodity MComm On MCG.CommodityGroupID = MComm.CommodityGroupID"
tmp = tmp & " Where MComm.CommodityID=" & mCommodityID_ & "  And MNR.LocationID=" & mLocationID
tmp = tmp & " And  MNRC.GodownID=" & mGodownID_

Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   Retval = TmpRs2!Rate
End If
GetStorageRate = Retval
End Function

Function getCDTFBOE_Max() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(BOEID) from Txn_CDTFBOE"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCDTFBOE_Max = Retval
End Function

Function getCDTFBOE_GSL_Max() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(BOEGSLID) from Txn_CDTFBOE_GSL"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getCDTFBOE_GSL_Max = Retval
End Function

Private Function ChkSelection() As Boolean
Dim Retval As Boolean
Retval = False

For k = 1 To MSHFlexGrid1.Rows
    If MSHFlexGrid1.TextMatrix(k, 1) = "" Then Exit For
    If MSHFlexGrid1.TextMatrix(k, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
ChkSelection = Retval
End Function

