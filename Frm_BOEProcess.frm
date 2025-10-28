VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_BOEProcess 
   Caption         =   "BOE Process"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   9000
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   4695
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Process"
         Height          =   375
         Left            =   2865
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109379585
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109379585
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
         Height          =   1485
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2619
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   263
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_BOEProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityGroupId, mLocationID, mCxTFNo, mCommodityID, mSM_Prefix As Variant
Dim wc As Variant


Private Sub CmdProcess_Click()

MsgBox "Please wait till next message!!"
CmdProcess.Enabled = False

tmp = "Select TCD.CxTFNo,TCD.SM_Prefix,TCD.CMPID,CMP.LocationID,CM.CommodityGroupID,TCD.CommodityID,"
tmp = tmp & " TCB.BOEAmount from  Txn_CxTF_Details TCD"
tmp = tmp & " left Join Txn_CxTF_BOE TCB on TCD.SM_Prefix=TCB.SM_Prefix And TCD.CxTFNo=TCB.CxTFNo"
tmp = tmp & " inner Join Mst_CMP CMP on TCD.CMPID=CMP.CMPID"
tmp = tmp & " inner Join Mst_Commodity CM on TCD.CommodityID=CM.CommodityID"
tmp = tmp & " Where TCD.CxtfDate>='" & Format(SFrom.Value, Gen_DatFormat) & "' And  ExchangeTypeID=1 And TxnType='W' And (TCB.BOEAmount is null or TCB.BOEAmount=0)"

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   For r = 1 To TmpRs4.RecordCount
       DoEvents
       mCommodityGroupId = IIf(IsNull(TmpRs4!CommodityGroupID), 0, TmpRs4!CommodityGroupID)
       mLocationID = TmpRs4!LocationID
       mCommodityID = TmpRs4!CommodityID
       mCxTFNo = TmpRs4!CxTFNO
       mSM_Prefix = TmpRs4!SM_Prefix
       Call FillBOE(TmpRs4!CxTFNO, TmpRs4!SM_Prefix)
       Call SaveCxTFBOE
       TmpRs4.MoveNext
   Next
   
   MsgBox "Done!!"
Else
   MsgBox "No record found!!"
End If
End Sub
Private Sub SaveCxTFBOE()
 
For I = 1 To MSHFlexGrid6.Rows - 2
    If MSHFlexGrid6.TextMatrix(I, 0) = "" Then Exit For
    Call TableTxn_CxTF_BOE(" where CxTFNo=" & Val(mCxTFNo) & " And LocationID=" & mLocationID)
    If RsTxn_CxTF_BOE.RecordCount > 0 Then
       RsTxn_CxTF_BOE.MoveFirst
       RsTxn_CxTF_BOE.Find "ClientID= " & Val(MSHFlexGrid6.TextMatrix(I, 3))
    End If
    If RsTxn_CxTF_BOE.EOF = False Then
       'RsTxn_CxTF_BOE!CxTFDate = MSHFlexGrid6.TextMatrix(i, 3)
       RsTxn_CxTF_BOE!WithdrawalBags = IIf(IsNull(RsTxn_CxTF_BOE!WithdrawalBags), 0, RsTxn_CxTF_BOE!WithdrawalBags) + Val(MSHFlexGrid6.TextMatrix(I, 8))
       RsTxn_CxTF_BOE!WithdrawalQty = IIf(IsNull(RsTxn_CxTF_BOE!WithdrawalQty), 0, RsTxn_CxTF_BOE!WithdrawalQty) + Val(MSHFlexGrid6.TextMatrix(I, 9))
       RsTxn_CxTF_BOE!BOEAmount = IIf(IsNull(RsTxn_CxTF_BOE!BOEAmount), 0, RsTxn_CxTF_BOE!BOEAmount) + Val(MSHFlexGrid6.TextMatrix(I, 16))
       RsTxn_CxTF_BOE.Update
    Else
       RsTxn_CxTF_BOE.AddNew
       RsTxn_CxTF_BOE!BOEID = getMaxBOEID
       RsTxn_CxTF_BOE!CxTFNO = mCxTFNo
       RsTxn_CxTF_BOE!LocationID = mLocationID
       RsTxn_CxTF_BOE!CommodityID = mCommodityID
       RsTxn_CxTF_BOE!SM_Prefix = mSM_Prefix
       RsTxn_CxTF_BOE!ClientIDRow = Val(MSHFlexGrid6.TextMatrix(I, 3))
       RsTxn_CxTF_BOE!CxTFDate = MSHFlexGrid6.TextMatrix(I, 0)
       RsTxn_CxTF_BOE!WithdrawalBags = Val(MSHFlexGrid6.TextMatrix(I, 8))
       RsTxn_CxTF_BOE!WithdrawalQty = Val(MSHFlexGrid6.TextMatrix(I, 9))
       RsTxn_CxTF_BOE!BOEAmount = Val(MSHFlexGrid6.TextMatrix(I, 16))
       RsTxn_CxTF_BOE!G_UID = GetGUID
       RsTxn_CxTF_BOE.Update
    End If
Next

End Sub

Private Function getMaxBOEID() As Double
Dim Retval As Double
tmp = "Select max(BOEID) from Txn_CxTF_BOE"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxBOEID = Retval
End Function

Private Sub Form_Load()
If LCase(Gen_UserRole) <> "power" Then
   CmdProcess.Enabled = False
   MsgBox "Access Denied!!"
   Exit Sub
End If

If Gen_UserLoginID <> "2047" And Gen_UserLoginID <> "1076" Then
   CmdProcess.Enabled = False
   Exit Sub
End If

CmdProcess.Enabled = True
SFrom = "04/01/2011"
STo = Date
End Sub

Private Sub Grid_Head_Bill()

MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 2
MSHFlexGrid6.FixedRows = 1
MSHFlexGrid6.Cols = 17

MSHFlexGrid6.TextMatrix(0, 0) = "CIDF Date"
MSHFlexGrid6.TextMatrix(0, 1) = "CDTF Date"
MSHFlexGrid6.TextMatrix(0, 2) = "No. of days"
MSHFlexGrid6.TextMatrix(0, 3) = "ClientRowID"
MSHFlexGrid6.TextMatrix(0, 4) = "Client Name"
MSHFlexGrid6.TextMatrix(0, 5) = "GSLID"
MSHFlexGrid6.TextMatrix(0, 6) = "G-S-L"
MSHFlexGrid6.TextMatrix(0, 7) = "Lot Type"
MSHFlexGrid6.TextMatrix(0, 8) = "No of Bags withdrawn"
MSHFlexGrid6.TextMatrix(0, 9) = "Withdrawn Qty"
MSHFlexGrid6.TextMatrix(0, 10) = "Bag Size"
MSHFlexGrid6.TextMatrix(0, 11) = "NContractID"
MSHFlexGrid6.TextMatrix(0, 12) = "Billing Cycle"
MSHFlexGrid6.TextMatrix(0, 13) = "Billing Unit"
MSHFlexGrid6.TextMatrix(0, 14) = "No. of Cycle"
MSHFlexGrid6.TextMatrix(0, 15) = "Rate per Cycle"
MSHFlexGrid6.TextMatrix(0, 16) = "Storage Charges Amount"

MSHFlexGrid6.ColWidth(0) = 1200
MSHFlexGrid6.ColWidth(1) = 1200
MSHFlexGrid6.ColWidth(2) = 1000
MSHFlexGrid6.ColWidth(3) = 0
MSHFlexGrid6.ColWidth(4) = 2400
MSHFlexGrid6.ColWidth(5) = 0
MSHFlexGrid6.ColWidth(6) = 1400
MSHFlexGrid6.ColWidth(7) = 1200
MSHFlexGrid6.ColWidth(8) = 1200
MSHFlexGrid6.ColWidth(9) = 1200
MSHFlexGrid6.ColWidth(10) = 1000
MSHFlexGrid6.ColWidth(11) = 0
MSHFlexGrid6.ColWidth(12) = 1200
MSHFlexGrid6.ColWidth(13) = 1200
MSHFlexGrid6.ColWidth(14) = 1200
MSHFlexGrid6.ColWidth(15) = 1200
MSHFlexGrid6.ColWidth(16) = 1600
MSHFlexGrid6.RowHeight(0) = 600

End Sub


Private Sub FillBOE(CITFNo_ As Variant, SM_Prefix_ As Variant)

Dim I As Variant

Call Grid_Head_Bill

'---For Spillage with bags=0
If CITFNo_ = 133 And SM_Prefix_ = "KBAGH" Then
   Exit Sub
End If

tmp = "Select TC.CxTFDate,GSL.CDTFDate,datediff(day,GSL.CDTFDate,TC.CxTFDate)+ 1 NoOfDay,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownNo+'~'+"
tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID"
tmp = tmp & " From  Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTFNo=TC.CxTFNo And TCG.TxnType=TC.TxnType"
tmp = tmp & " left Join Mst_GSL GSL on TCG.GSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Where TC.ExchangeTypeID=1 And TC.CxTFNo=" & Val(CITFNo_) & " And TC.TxnType='W' And TC.SM_Prefix='" & SM_Prefix_ & "'"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid6.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid6.TextMatrix(I, 0) = Format(MSHFlexGrid6.TextMatrix(I, 0), "dd-MMM-yyyy")
       MSHFlexGrid6.TextMatrix(I, 1) = Format(MSHFlexGrid6.TextMatrix(I, 1), "dd-MMM-yyyy")
       If MSHFlexGrid6.TextMatrix(I, 7) = "1" Then
          MSHFlexGrid6.TextMatrix(I, 7) = "Spillage"
       Else
          MSHFlexGrid6.TextMatrix(I, 7) = "Normal"
          'Call FillBillingDetails(i)
          Call FillBillingDetails(I, TmpRs!CDTFDate, TmpRs!CxTFDate)
       End If
       If MSHFlexGrid6.TextMatrix(I, 13) = "Bag" Then
          MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 8)), "#####0.00")
       Else
          MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
       End If
       TmpRs.MoveNext
       MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
   Next
End If

End Sub
Private Sub FillBillingDetails(Row_ As Variant, mFromDate As Date, mToDate As Date)

If Val(MSHFlexGrid6.TextMatrix(Row_, 11)) <> 0 Then Exit Sub

tmp = "Select SPRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,BillingUnit,Amount"
tmp = tmp & " From Mst_SpecialRateCard MS"
tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID"
tmp = tmp & " Where ClientIDRow = " & Val(MSHFlexGrid6.TextMatrix(Row_, 3)) & " And CommodityGroupID = " & mCommodityGroupId & " And LocationID =" & mLocationID & "  And ExchangeTypeID = 1"

Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid6.TextMatrix(Row_, 12) = TmpRs2!BillingCycle
   MSHFlexGrid6.TextMatrix(Row_, 13) = IIf(TmpRs2!BillingUnit = "B", "Bag", "Qty")
    
    'MSHFlexGrid6.TextMatrix(Row_, 14) = Val(MSHFlexGrid6.TextMatrix(Row_, 2)) / TmpRs2!NoofDays
    'MSHFlexGrid6.TextMatrix(Row_, 14) = CDbl(CLng(Val(MSHFlexGrid6.TextMatrix(Row_, 14)) + 0.5))
   
   MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoOfDays)

   If TmpRs2!BillingUnit = "Q" Then
      MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs2!Amount
   Else
'      tmp = "Select RatePerBag From Mst_SpecialRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SPRCID=" & TmpRs2!SPRCID
'      Call Tmp1Table
'      If TmpRs1.RecordCount > 0 Then
'         MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
'      End If

      If DateDiff("M", mFromDate, mToDate) = 0 Then
         MSHFlexGrid6.TextMatrix(Row_, 15) = 0
      Else
         tmp = "Select RatePerBag From Mst_SpecialRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SPRCID=" & TmpRs2!SPRCID
         Call Tmp1Table
         If TmpRs1.RecordCount > 0 Then
            MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
         End If
      End If
      
   End If
Else
    tmp = "Select SRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,MS.Flag,RateOnQty"
    tmp = tmp & " From Mst_StorageRateCard MS"
    tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID"
    tmp = tmp & " Where CommodityGroupID = " & mCommodityGroupId & " And LocationID =" & mLocationID & "  And ExchangeTypeID = 1 And DefaultRate = 1 "
    
    Call Tmp2Table
    
    If TmpRs2.RecordCount > 0 Then
       MSHFlexGrid6.TextMatrix(Row_, 12) = TmpRs2!BillingCycle
       MSHFlexGrid6.TextMatrix(Row_, 13) = IIf(TmpRs2!Flag = "B", "Bag", "Qty")
        'MSHFlexGrid6.TextMatrix(Row_, 14) = Val(MSHFlexGrid6.TextMatrix(Row_, 2)) / TmpRs2!NoOfDays
        'MSHFlexGrid6.TextMatrix(Row_, 14) = CDbl(CLng(Val(MSHFlexGrid6.TextMatrix(Row_, 14)) + 0.5))
       MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoOfDays)
       If TmpRs2!Flag = "Q" Then
          MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs2!RateOnQty
       Else
'          tmp = "Select RatePerBag From Mst_StorageRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
'          Call Tmp1Table
'          If TmpRs1.RecordCount > 0 Then
'             MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
'          End If


         If DateDiff("M", mFromDate, mToDate) = 0 Then
            MSHFlexGrid6.TextMatrix(Row_, 15) = 0
         Else
            tmp = "Select RatePerBag From Mst_StorageRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
               MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
            End If
         End If
         
       End If
    End If
End If
End Sub

Private Function CalculateNoOfCycle(mFromDate_ As Date, mToDate_ As Date, mBillingCycleID_ As Variant, mNoOfDays_ As Variant)
Dim mCycleCount As Variant
Dim MonthCount, DepositMonthCycles, WithdrawalMonthCycles As Variant
'D=1 Daily
If TmpRs2!BillingCycleID = 1 Then
   mCycleCount = DateDiff("d", mFromDate, mToDate) + 1
'w=2 'Weekly
ElseIf mBillingCycleID_ = 2 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) - 1
'   DepositMonthCycles = 4 - Left((DatePart("d", mFromDate_) / mNoOfDays_), 2)
   
   If Format(DatePart("d", mFromDate_) / mNoOfDays_) = 1 Then
      DepositMonthCycles = 4 '- 0 'Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
   ElseIf Format(DatePart("d", mFromDate_) / mNoOfDays_) > 4 Then
      DepositMonthCycles = 4 - (Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) - 1)
   Else
      DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
   End If
   
   WithdrawalMonthCycles = RoundUp(DatePart("d", mToDate_) / mNoOfDays_)
   mCycleCount = (MonthCount * 4) + DepositMonthCycles + WithdrawalMonthCycles
'F=3 'Fortnightly
ElseIf mBillingCycleID_ = 3 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) - 1
   DepositMonthCycles = IIf(DatePart("d", mFromDate_) - mNoOfDays_ >= 0, 1, 2)
   WithdrawalMonthCycles = IIf(DatePart("d", mToDate_) - mNoOfDays_ >= 0, 2, 1)
   mCycleCount = (MonthCount * 2) + DepositMonthCycles + WithdrawalMonthCycles
'M=4 'Monthly
ElseIf TmpRs2!BillingCycleID = 4 Then
   mCycleCount = DateDiff("M", mFromDate, mToDate) + 1
End If
CalculateNoOfCycle = mCycleCount
End Function
