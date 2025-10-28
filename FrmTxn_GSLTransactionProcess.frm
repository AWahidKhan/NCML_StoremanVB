VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_GSLTransactionProcess 
   Caption         =   "Transaction GSL Process"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17055
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15240
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13680
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   1
         Top             =   5400
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker TxtMonthYear 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   113115139
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label1 
         Caption         =   "Billing Month-Year"
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
         TabIndex        =   4
         Top             =   300
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmTxn_GSLTransactionProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonthYear As Variant, mTxnFromDate As Variant
Dim mR As Double, mC As Double, mRow As Double, mCol As Double
Dim mMonthYearStartDate As Variant, mMonthYearEndDate As Variant

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdProcess_Click()

Dim mGSLID As Double, mLocationID As Double, mClientIDRow As Double, mGodownID As Double, mCommodityID As Double, mTansactionNo As Double, mExchangeTypeID As Double
Dim mSM_Prefix As Variant, mContractID As Variant, mBags As Double, mQty As Double, mWBags As Double, mWQty As Double

Dim mOldGSLID As Double, mNewGSLID As Double, mOldLocationID As Double, mNewLocationID As Double
Dim mOldSM_Prefix As Variant, mNewSM_Prefix As Variant

Dim mTSM_Prefix As Variant
Dim mTGSLID As Double, mTLocationID As Double, mTGodownID As Double, mTCommodityID As Double, mTClientIDRow As Double, mTExchangeTypeID As Double, mTTansactionNo As Double, mTBags As Double, mTQty As Double
Dim mTransferMadeupFlag As Boolean

If CDate(Format(TxtMonthYear.Value, "yyyy-MM-01")) >= CDate(Format(Date, "yyyy-MM-01")) Then
    MsgBox "Future Month-Year and Current Month-Year Not allow."
    Exit Sub
End If

MsgBox "Wait for next message."

mMonthYear = Format(TxtMonthYear, "yyyy-MM-01")
mTxnFromDate = Format(DateAdd("m", 1, mMonthYear), "yyyy-MM-dd")
mMonthYearStartDate = mMonthYear
mMonthYearEndDate = Format(DateAdd("m", 1, mMonthYear) - 1, "yyyy-MM-dd")

tmp = " IF OBJECT_ID('dbo.TmpInvMonGSLDetail', 'U') IS NOT NULL"
tmp = tmp & " DROP TABLE dbo.TmpInvMonGSLDetail "
Call TmpTable

tmp = " Select A.GSLID,A.SM_Prefix,ML.LocationID,A.CommodityID,A.GodownID,A.StackNo,A.LotNo,A.BalanceBags,A.BalanceWeight,A.ExchangeTypeID,A.ClientIDRow,A.ContractID,A.YearMonth,A.CDTFNo,A.CDTFDate,A.CxTFNO,A.CxTFDate,A.TFlag,A.MadeUpFlag"
tmp = tmp & " Into TmpInvMonGSLDetail"
tmp = tmp & " from Txn_InvMonGSLDetail A Inner Join Mst_Location ML On A.SM_Prefix = ML.SM_Prefix Where A.YearMonth  = '" & mMonthYear & "'"
'tmp = tmp & " And A.SM_Prefix = 'ADLBD' And A.GSLID = 2685 "

Call TmpTable

''-- Call Txn_GSL
Call TableTxn_GSL

''---------------------------------------------------------COmmented by Parag on 29th Nov 2018 as O flag added in query specified below
''---------------------------------------------------------
''---- Process for Opening Balance
'ProgressBar1.Value = 0

'tmp = " Select * From TmpInvMonGSLDetail As A "
'tmp = tmp & " Where A.YearMonth = '" & mMonthYear & "' And A.TFlag = 'O' And A.BalanceWeight != 0 "
'tmp = tmp & " Order By A.GSLID Asc,A.SM_Prefix Asc,A.CommodityID Asc,A.GodownID Asc,A.ClientIDRow Asc,Case A.TFlag When 'O' Then 1 When 'D' Then 2 When 'W' Then 3 Else 4 End"
'
'Call TmpTable
'
'ProgressBar1.Max = TmpRs.RecordCount
'
'For mR = 1 To TmpRs.RecordCount
'    If TmpRs!GSLID = 2685 Then
'       a = a
'    End If
'    mGSLID = TmpRs!GSLID
'    mSM_Prefix = TmpRs!SM_Prefix
'
'    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
'    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
'    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
'    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
'    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
'    mTansactionNo = IIf(IsNull(TmpRs!CdtfNo), 0, TmpRs!CdtfNo)
'
'    mBags = TmpRs!BalanceBags
'    mQty = TmpRs!BalanceWeight
'
'    Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "OPENING BALANCE", CDate(mMonthYear), mBags, mQty, 0, 0, Date, TmpRs!CDTFDate, "D")
'
'    ProgressBar1.Value = mR
'    TmpRs.MoveNext
'Next mR
''---------------------------------------------------------

''---------------------------------------------------------
''--- Process for Opening Balance from Deposit
ProgressBar1.Value = 0
''-- Call Withdrawal from (O/P)
tmp = " Select A.GSLID,A.SM_Prefix,A.CommodityID,A.GodownID,Sum(A.BalanceBags)BalanceBags,Sum(A.BalanceWeight)BalanceWeight,A.ExchangeTypeID,A.ClientIDRow,A.LocationID,A.YearMonth,A.TFlag,A.CDTFNo,A.CDTFDate,A.MadeUpFlag,A.ContractID  "
tmp = tmp & " From TmpInvMonGSLDetail As A "
tmp = tmp & " Where A.YearMonth = '" & mMonthYear & "' And A.TFlag = 'W' "
tmp = tmp & " Group By A.GSLID,A.SM_Prefix,A.CommodityID,A.GodownID,A.ExchangeTypeID,A.ClientIDRow,A.LocationID,A.YearMonth,A.TFlag,A.CDTFNo,A.CDTFDate,A.MadeUpFlag,A.ContractID "
tmp = tmp & " Order By A.GSLID Asc,A.SM_Prefix Asc,A.CommodityID Asc,A.GodownID Asc,A.ClientIDRow Asc,Case A.TFlag When 'O' Then 1 When 'D' Then 2 When 'W' Then 3 Else 4 End "
Call Tmp1Table

''-- Call Transfer Transaction
tmp = " Select Trns.*,ML.LocationID,MLL.LocationID[NewLocationID] from Txn_GSL_Transfer Trns Inner Join Mst_Location ML On Trns.SM_Prefix = ML.SM_Prefix Inner Join Mst_Location MLL On Trns.NewSM_Prefix = MLL.SM_Prefix Where (Trns.GSLTransferDate >= '" & mMonthYearStartDate & "' ANd Trns.GSLTransferDate <= '" & mMonthYearEndDate & "') "
Call Tmp2Table

''-- Call Deposit from (O/P)
tmp = " Select A.GSLID,A.SM_Prefix,A.CommodityID,A.GodownID,Sum(A.BalanceBags)BalanceBags,Sum(A.BalanceWeight)BalanceWeight,A.ExchangeTypeID,A.ClientIDRow,A.LocationID,A.YearMonth,A.TFlag,A.CDTFNo,A.CDTFDate,A.MadeUpFlag,A.ContractID  "
tmp = tmp & " From TmpInvMonGSLDetail As A "
tmp = tmp & " Where A.YearMonth = '" & mMonthYear & "' And A.TFlag in ('D','O')"
tmp = tmp & " Group By A.GSLID,A.SM_Prefix,A.CommodityID,A.GodownID,A.ExchangeTypeID,A.ClientIDRow,A.LocationID,A.YearMonth,A.TFlag,A.CDTFNo,A.CDTFDate,A.MadeUpFlag,A.ContractID  "
tmp = tmp & " Order By A.GSLID Asc,A.SM_Prefix Asc,A.CommodityID Asc,A.GodownID Asc,A.ClientIDRow Asc,Case A.TFlag When 'O' Then 1 When 'D' Then 2 When 'W' Then 3 Else 4 End"
Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 1
For mR = 1 To TmpRs.RecordCount
    mGSLID = TmpRs!GSLID
    mSM_Prefix = TmpRs!SM_Prefix
    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mTansactionNo = IIf(IsNull(TmpRs!CdtfNo), 0, TmpRs!CdtfNo)
    mContractID = IIf(IsNull(TmpRs!ContractID), 0, TmpRs!ContractID)
    mBags = TmpRs!BalanceBags
    mQty = TmpRs!BalanceWeight
    mWBags = 0
    mWQty = 0
    
    ''-- If Deposit Bags are Transfer In Same Month Then Check the entry in Transfer
    ''-- If transfer is Full then omit Deposit Entry and Make the O/P Entry agaisnt Transfer GSL
    ''-- If transfer is partial then O/P Deposit entry for remain Bags as well as O/P Deposit for Transfer Bags
        
    TmpRs2.Filter = ""
    TmpRs2.MoveFirst
    TmpRs2.Filter = " OLDGSLID = " & mGSLID & " And SM_Prefix = '" & mSM_Prefix & "' "
    If TmpRs2.EOF = False Then
        For mRow = 1 To TmpRs2.RecordCount
            mWBags = mWBags + TmpRs2!TransferedBags
            mWQty = mWQty + TmpRs2!TransferedWeight
            TmpRs2.MoveNext
        Next mRow
        If mBags - mWBags > 0 Then
            Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "OPENING BALANCE", CDate(mMonthYear), (mBags - mWBags), (mQty - mWQty), 0, 0, Date, TmpRs!CDTFDate, "D")
        End If
        ''-- Add the Entry for Transfer GSLs
        TmpRs2.MoveFirst
        For mRow = 1 To TmpRs2.RecordCount
            
            mTGSLID = TmpRs2!NewGSLID
            mTSM_Prefix = TmpRs2!NewSM_Prefix
            mTLocationID = TmpRs2!NewLocationID
            mTTansactionNo = TmpRs2!GSLTransferID
            
            Call TableMst_GSL(" Where GSLID = " & mTGSLID & " And SM_Prefix = '" & mTSM_Prefix & "' And Flag <> 'Z' ")
            If RsMst_GSL.RecordCount > 0 Then
                mTGodownID = IIf(IsNull(RsMst_GSL!GodownID), 0, RsMst_GSL!GodownID)
                mTCommodityID = IIf(IsNull(RsMst_GSL!CommodityID), 0, RsMst_GSL!CommodityID)
                mTClientIDRow = IIf(IsNull(RsMst_GSL!ClientIDRow), 0, RsMst_GSL!ClientIDRow)
                mTExchangeTypeID = IIf(IsNull(RsMst_GSL!ExchangeTypeID), 0, RsMst_GSL!ExchangeTypeID)
               
                mTBags = RsMst_GSL!BalanceBags
                mTQty = RsMst_GSL!BalanceWeight
                Call Add_TxnGSL(mTGSLID, mTLocationID, mTSM_Prefix, mTTansactionNo, mTGodownID, mTCommodityID, mTClientIDRow, mTExchangeTypeID, "OPENING BALANCE", CDate(mMonthYear), mTBags, mTQty, 0, 0, Date, RsMst_GSL!CDTFDate, "D")
            End If
            
            ''-- Check the Transfer GSL Having Any Madeup GSL or Not if available the Deposit in O/P
            If IsNull(TmpRs2!MadeupOLDGSLID) = False And IsNull(TmpRs2!MadeupNewGSLID) = False Then
                mTransferMadeupFlag = True
            Else
                mTransferMadeupFlag = False
            End If
            
            If mTransferMadeupFlag = True Then
                mTGSLID = TmpRs2!MadeupNewGSLID
                mTSM_Prefix = TmpRs2!NewSM_Prefix
                mTLocationID = TmpRs2!NewLocationID
                mTTansactionNo = TmpRs2!GSLTransferID
                
                Call TableMst_GSL(" Where GSLID = " & mTGSLID & " And SM_Prefix = '" & mTSM_Prefix & "' And Flag <> 'Z' ")
                If RsMst_GSL.RecordCount > 0 Then
                    mTGodownID = IIf(IsNull(RsMst_GSL!GodownID), 0, RsMst_GSL!GodownID)
                    mTCommodityID = IIf(IsNull(RsMst_GSL!CommodityID), 0, RsMst_GSL!CommodityID)
                    mTClientIDRow = IIf(IsNull(RsMst_GSL!ClientIDRow), 0, RsMst_GSL!ClientIDRow)
                    mTExchangeTypeID = IIf(IsNull(RsMst_GSL!ExchangeTypeID), 0, RsMst_GSL!ExchangeTypeID)
                   
                    mTBags = RsMst_GSL!BalanceBags
                    mTQty = RsMst_GSL!BalanceWeight
                    Call Add_TxnGSL(mTGSLID, mTLocationID, mTSM_Prefix, mTTansactionNo, mTGodownID, mTCommodityID, mTClientIDRow, mTExchangeTypeID, "OPENING BALANCE", CDate(mMonthYear), mTBags, mTQty, 0, 0, Date, RsMst_GSL!CDTFDate, "D")
                End If
            End If
            ''---------------------------------
            
            TmpRs2.MoveNext
        Next mRow
        ''------------
    Else
        ''-- Check for Withdarwal entry if Transfer not done against Deposit
        TmpRs1.Filter = ""
        TmpRs1.MoveFirst
        tmp = " ExchangeTypeID = " & mExchangeTypeID & " And GSLID = " & mGSLID & " And LocationID = " & mLocationID & " And CommodityID = " & mCommodityID & " And ClientIDRow = " & mClientIDRow & " "
        TmpRs1.Filter = tmp
        If TmpRs1.EOF = False Then
            mWBags = TmpRs1!BalanceBags
            mWQty = TmpRs1!BalanceWeight
        Else
            mWBags = 0
            mWQty = 0
        End If
        
        ''-- Calculate the Balance Bags & Qty with withdrawal
        ''-- if Balance Bags is zero then goto next entry but if ContractID is available then check with Balance Bags if difference is avaible then make entry other wise next
        
        If mBags - mWBags > 0 Then
            Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "OPENING BALANCE", CDate(mMonthYear), (mBags - mWBags), (mQty - mWQty), 0, 0, Date, TmpRs!CDTFDate, "D")
        Else
            If mContractID <> 0 And mQty <> mWQty Then
                Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "OPENING BALANCE", CDate(mMonthYear), (mBags - mWBags), (mQty - mWQty), 0, 0, Date, TmpRs!CDTFDate, "D")
            End If
        End If
        ''------------------------
    End If
    
    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR
''---------------------------------------------------------

''---------------------------------------------------------
''-- All Deposit Trasaction Entry from Next Month
ProgressBar1.Value = 0
tmp = " Select DGSL.GSLID,DGSL.SM_Prefix,ML.LocationID,DGSL.CxTFNo,DD.CxTFDate,DGSL.TxnType,DGSL.ClientIDRow,DD.CommodityID,DD.ExchangeTypeID,Sum(DGSL.NoOfBags)Bags,Sum(DGSL.DematWeight)Qty,GSL.CDTFDate,GSL.GodownID"
tmp = tmp & " from Txn_CxTF_GSL DGSL"
tmp = tmp & " Inner Join Txn_CxTF_Details DD On DGSL.CxTFNo = DD.CxTFNo And DGSL.SM_Prefix = DD.SM_Prefix And DGSL.TxnType = DD.TxnType"
tmp = tmp & " Inner Join Mst_Location ML On DGSL.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL GSL On DGSL.GSLID = GSL.GSLID ANd DGSL.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Where DD.TxnType = 'D' And CxTFDate >= '" & mTxnFromDate & "'"
tmp = tmp & " Group By DGSL.GSLID,DGSL.SM_Prefix,ML.LocationID,DGSL.CxTFNo,DD.CxTFDate,DGSL.TxnType,DGSL.ClientIDRow,DD.CommodityID,DD.ExchangeTypeID,GSL.CDTFDate,GSL.GodownID"
tmp = tmp & " Order By DD.CxTFDate,DGSL.GSLID,DGSL.SM_Prefix,ML.LocationID,DGSL.CxTFNo,DGSL.ClientIDRow,DD.CommodityID,DD.ExchangeTypeID,GSL.CDTFDate"
Call TmpTable
ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
    mGSLID = TmpRs!GSLID
    mSM_Prefix = TmpRs!SM_Prefix
    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mTansactionNo = IIf(IsNull(TmpRs!CxTFNO), 0, TmpRs!CxTFNO)
    mBags = TmpRs!Bags
    mQty = TmpRs!Qty
    
    Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "CDTF", TmpRs!CxTFDate, mBags, mQty, 0, 0, Date, TmpRs!CDTFDate, "D")

    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR
''---------------------------------------------------------

''---------------------------------------------------------
''-- All Withdrawal Trasaction Entry from Next month
ProgressBar1.Value = 0
tmp = " Select DGSL.GSLID,DGSL.SM_Prefix,ML.LocationID,DGSL.CxTFNo,DD.CxTFDate,DGSL.TxnType,DGSL.ClientIDRow,DD.CommodityID,DD.ExchangeTypeID,Sum(DGSL.NoOfBags)Bags,Sum(DGSL.DematWeight)Qty,GSL.CDTFDate,GSL.GodownID"
tmp = tmp & " from Txn_CxTF_GSL DGSL"
tmp = tmp & " Inner Join Txn_CxTF_Details DD On DGSL.CxTFNo = DD.CxTFNo And DGSL.SM_Prefix = DD.SM_Prefix And DGSL.TxnType = DD.TxnType"
tmp = tmp & " Inner Join Mst_Location ML On DGSL.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL GSL On DGSL.GSLID = GSL.GSLID ANd DGSL.SM_Prefix = GSL.SM_Prefix"
'tmp = tmp & " Where DD.TxnType = 'W' And CxTFDate >= '" & mTxnFromDate & "'"
tmp = tmp & " Where DD.TxnType = 'W' And DD.CreatedDate >= '" & mTxnFromDate & "'"
tmp = tmp & " Group By DGSL.GSLID,DGSL.SM_Prefix,ML.LocationID,DGSL.CxTFNo,DD.CxTFDate,DGSL.TxnType,DGSL.ClientIDRow,DD.CommodityID,DD.ExchangeTypeID,GSL.CDTFDate,GSL.GodownID"
tmp = tmp & " Order By DD.CxTFDate,DGSL.GSLID,DGSL.SM_Prefix,ML.LocationID,DGSL.CxTFNo,DGSL.ClientIDRow,DD.CommodityID,DD.ExchangeTypeID,GSL.CDTFDate"
Call TmpTable
ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
    mGSLID = TmpRs!GSLID
    mSM_Prefix = TmpRs!SM_Prefix
    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mTansactionNo = IIf(IsNull(TmpRs!CxTFNO), 0, TmpRs!CxTFNO)
    mBags = TmpRs!Bags
    mQty = TmpRs!Qty
        
    Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "CITF", TmpRs!CxTFDate, mBags * -1, mQty * -1, 0, 0, Date, TmpRs!CDTFDate, "W")
    
    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR
''---------------------------------------------------------

''---------------------------------------------------------
''-- Adjustment Withdrawal Trasaction Entry
ProgressBar1.Value = 0
tmp = " Select AdjGSL.TxnID,Adj.TxnDate,AdjGSL.GSLID,Adj.SM_Prefix,ML.LocationID,Adj.CommodityID,Adj.TxnType,AdjGSL.NoOfBags Bags,AdjGSL.Quantity Qty,AdjGSL.SpillageFlag,GSL.CDTFDate,GSL.GodownID,GSL.ClientIDRow,GSL.ExchangeTypeID"
tmp = tmp & " from Txn_GSLAdjTransaction AdjGSL"
tmp = tmp & " Inner Join Txn_AdjTransaction Adj On AdjGSL.TxnID = Adj.TxnID"
tmp = tmp & " Inner Join Mst_Location ML On Adj.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL GSL On AdjGSL.GSLID = GSL.GSLID ANd Adj.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Where Adj.TxnDate >= '" & mTxnFromDate & "'"
tmp = tmp & " Order By AdjGSL.GSLID,Adj.SM_Prefix,Case Adj.TxnType When 'AG' Then 1 When 'AR' Then 2 When 'SA' Then 3 When 'Th' Then 4 When 'Fi' Then 5 Else 6 End"
Call TmpTable
ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
    mGSLID = TmpRs!GSLID
    mSM_Prefix = TmpRs!SM_Prefix
    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mTansactionNo = IIf(IsNull(TmpRs!TxnID), 0, TmpRs!TxnID)
    mBags = TmpRs!Bags
    mQty = TmpRs!Qty
        
    If LCase("AG") = LCase(TmpRs!TxnType) Then
        Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "Adjustment Given", TmpRs!TxnDate, mBags * -1, mQty * -1, 0, 0, Date, TmpRs!CDTFDate, "W")
    ElseIf LCase("AR") = LCase(TmpRs!TxnType) Then
        Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "Adjustment Received", TmpRs!TxnDate, mBags, mQty, 0, 0, Date, TmpRs!CDTFDate, "D")
    ElseIf LCase("SA") = LCase(TmpRs!TxnType) Then
        Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "SALE", TmpRs!TxnDate, mBags * -1, mQty * -1, 0, 0, Date, TmpRs!CDTFDate, "W")
    ElseIf LCase("TH") = LCase(TmpRs!TxnType) Then
        Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "THEFT", TmpRs!TxnDate, mBags * -1, mQty * -1, 0, 0, Date, TmpRs!CDTFDate, "W")
    ElseIf LCase("FI") = LCase(TmpRs!TxnType) Then
        Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "FIRE", TmpRs!TxnDate, mBags * -1, mQty * -1, 0, 0, Date, TmpRs!CDTFDate, "W")
    End If
    
    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR
''---------------------------------------------------------

''---------------------------------------------------------
''--  GSL Transfer from GSL and To GSL Entry
ProgressBar1.Value = 0
tmp = " Select A.GSLTransferID,A.GSLTransferDate,A.OLDGSLID,A.SM_Prefix,ML.LocationID,A.TransferedBags,A.TransferedWeight,GSL.CDTFDate,GSL.GodownID,NewGSL.GodownID[NewGodownID],GSL.CommodityID,GSL.ClientIDRow,NewGSL.ClientIDRow[NewClientIDRow],GSL.ExchangeTypeID,A.NewGSLID,A.NewSM_Prefix,MLL.LocationID[NewLocatoinID],A.MadeupOLDGSLID,GSL.ConnectedGSLID,GSL.MadeUpFlag "
tmp = tmp & " from Txn_GSL_Transfer A"
tmp = tmp & " Inner Join Mst_Location ML On A.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location MLL On A.NewSM_Prefix = MLL.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL GSL On A.OLDGSLID = GSL.GSLID ANd A.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL NewGSL On A.NewGSLID = NewGSL.GSLID ANd A.NewSM_Prefix = NewGSL.SM_Prefix"
tmp = tmp & " Where A.GSLTransferDate >= '" & mTxnFromDate & "'  Order By A.GSLTransferDate"

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
    mOldGSLID = TmpRs!OldGSLID
    mOldSM_Prefix = TmpRs!SM_Prefix
    mOldLocationID = TmpRs!LocationID
    mNewGSLID = TmpRs!NewGSLID
    mNewSM_Prefix = TmpRs!NewSM_Prefix
    mNewLocationID = TmpRs!NewLocatoinID
    mBags = TmpRs!TransferedBags
    mQty = TmpRs!TransferedWeight
    mTansactionNo = IIf(IsNull(TmpRs!GSLTransferID), 0, TmpRs!GSLTransferID)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    
    ''---  Add Entry in Txn_GSL for Withdrawals from Trasfer
    Call Add_TxnGSL(mOldGSLID, mOldLocationID, mOldSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "TRANSFER", TmpRs!GSLTransferDate, mBags * -1, mQty * -1, mNewGSLID, 0, Date, TmpRs!CDTFDate, "W")
    
    If IsNull(TmpRs!MadeupOLDGSLID) = True Then
       If TmpRs!MadeupFlag = 0 Then
          If TmpRs!ConnectedGSLID > 0 Then
             tmp = "Select * from Mst_GSL Where MadeupFlag = '1' And GSLID = " & TmpRs!ConnectedGSLID & " And SM_Prefix = '" & mOldSM_Prefix & "' And Flag = 'Z'"
             Call Tmp2Table
             If TmpRs2.RecordCount > 0 Then
                Call Add_TxnGSL(TmpRs2!GSLID, mOldLocationID, mOldSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "TRANSFER", TmpRs!GSLTransferDate, TmpRs2!DepositeBags * -1, TmpRs2!DepositeWeight * -1, TmpRs2!FGSLID, 0, Date, TmpRs!CDTFDate, "W")
                Call Add_TxnGSL(TmpRs2!FGSLID, mNewLocationID, mNewSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "TRANSFER", TmpRs!GSLTransferDate, TmpRs2!DepositeBags, TmpRs2!DepositeWeight, 0, TmpRs2!GSLID, Date, TmpRs!CDTFDate, "D")
             End If
          End If
       End If
    End If
    
    mGodownID = IIf(IsNull(TmpRs!NewGodownID), 0, TmpRs!NewGodownID)
    mClientIDRow = IIf(IsNull(TmpRs!NewClientIDRow), 0, TmpRs!NewClientIDRow)
    ''---  Add Entry in Txn_GSL for Deposit from Trasfer
    Call Add_TxnGSL(mNewGSLID, mNewLocationID, mNewSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "TRANSFER", TmpRs!GSLTransferDate, mBags, mQty, 0, mOldGSLID, Date, TmpRs!CDTFDate, "D")
    
    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR
''---------------------------------------------------------

''---------------------------------------------------------
''--  GSL Transfer Madeup Bags from GSL and To GSL Entry
ProgressBar1.Value = 0
tmp = " Select A.GSLTransferID,A.GSLTransferDate,A.OLDGSLID,A.SM_Prefix,ML.LocationID,A.MadeupOLDGSLID,A.MadeupBags,0.00 [Weight],GSL.CDTFDate,GSL.GodownID,NewGSL.GodownID[NewGodownID],GSL.CommodityID,GSL.ClientIDRow,NewGSL.ClientIDRow[NewClientIDRow],GSL.ExchangeTypeID,A.NewGSLID,A.NewSM_Prefix,MLL.LocationID[NewLocatoinID],A.MadeupNewGSLID"
tmp = tmp & " from Txn_GSL_Transfer A"
tmp = tmp & " Inner Join Mst_Location ML On A.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_Location MLL On A.NewSM_Prefix = MLL.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL GSL On A.MadeupOLDGSLID = GSL.GSLID ANd A.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL NewGSL On A.MadeupNewGSLID = NewGSL.GSLID ANd A.NewSM_Prefix = NewGSL.SM_Prefix"
tmp = tmp & " Where A.GSLTransferDate >= '" & mTxnFromDate & "'  And A.MadeupOLDGSLID is Not Null Order By A.GSLTransferDate"
Call TmpTable
ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
    mOldGSLID = TmpRs!MadeupOLDGSLID
    mOldSM_Prefix = TmpRs!SM_Prefix
    mOldLocationID = TmpRs!LocationID
    mNewGSLID = TmpRs!MadeupNewGSLID
    mNewSM_Prefix = TmpRs!NewSM_Prefix
    mNewLocationID = TmpRs!NewLocatoinID
    mBags = TmpRs!MadeupBags
    mQty = TmpRs!Weight
    mTansactionNo = IIf(IsNull(TmpRs!GSLTransferID), 0, TmpRs!GSLTransferID)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    
    ''---  Add Entry in Txn_GSL for Withdrawals Madeup from Trasfer
    Call Add_TxnGSL(mOldGSLID, mOldLocationID, mOldSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "TRANSFER", TmpRs!GSLTransferDate, mBags * -1, mQty * -1, mNewGSLID, 0, Date, TmpRs!CDTFDate, "W")
    
    mGodownID = IIf(IsNull(TmpRs!NewGodownID), 0, TmpRs!NewGodownID)
    mClientIDRow = IIf(IsNull(TmpRs!NewClientIDRow), 0, TmpRs!NewClientIDRow)
    ''---  Add Entry in Txn_GSL for Deposit Madeup from Trasfer
    Call Add_TxnGSL(mNewGSLID, mNewLocationID, mNewSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "TRANSFER", TmpRs!GSLTransferDate, mBags, mQty, 0, mOldGSLID, Date, TmpRs!CDTFDate, "D")
    
    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR

''---------------------------------------------------------
'' -- Madeup
ProgressBar1.Value = 0
tmp = " Select A.SpillageID,A.SM_Prefix,ML.LocationID,A.SpillageDate,A.GSLID,A.CommodityID,A.SpillageNoOfBags,A.SpillageQuantity,A.Flag,A.NewGSLID,GSL.CDTFDate,GSL.GodownID,GSL.ClientIDRow,GSL.ExchangeTypeID"
tmp = tmp & " from Txn_Spillage AS A"
tmp = tmp & " Inner Join Mst_Location ML On A.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Left Join Mst_GSL GSL On A.NewGSLID = GSL.GSLID ANd A.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Where A.SpillageDate >= '" & mTxnFromDate & "' And A.Flag = 'M' Order By A.SpillageID,A.SpillageDate"
Call TmpTable
ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
    mGSLID = TmpRs!NewGSLID
    mSM_Prefix = TmpRs!SM_Prefix
    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mTansactionNo = IIf(IsNull(TmpRs!SpillageID), 0, TmpRs!SpillageID)
    mBags = TmpRs!SpillageNoOfBags
    mQty = TmpRs!SpillageQuantity
    
    Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "MADEUP", TmpRs!SpillageDate, mBags, mQty, 0, 0, Date, TmpRs!CDTFDate, "D")

    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR
''---------------------------------------------------------

''---------------------------------------------------------
'' -- Spillage
ProgressBar1.Value = 0
tmp = " Select A.SpillageID,A.SM_Prefix,ML.LocationID,A.SpillageDate,A.GSLID,A.CommodityID,A.SpillageNoOfBags,A.SpillageQuantity,A.Flag,A.NewGSLID,GSL.CDTFDate,GSL.GodownID,GSL.ClientIDRow,GSL.ExchangeTypeID"
tmp = tmp & " from Txn_Spillage AS A"
tmp = tmp & " Inner Join Mst_Location ML On A.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Left Join Mst_GSL GSL On A.GSLID = GSL.GSLID ANd A.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Where A.SpillageDate >= '" & mTxnFromDate & "' And A.Flag = 'S' Order By A.SpillageID,A.SpillageDate"
Call TmpTable
ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
    mGSLID = TmpRs!GSLID
    mSM_Prefix = TmpRs!SM_Prefix
    mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
    mGodownID = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
    mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
    mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
    mExchangeTypeID = IIf(IsNull(TmpRs!ExchangeTypeID), 0, TmpRs!ExchangeTypeID)
    mTansactionNo = IIf(IsNull(TmpRs!SpillageID), 0, TmpRs!SpillageID)
    mBags = TmpRs!SpillageNoOfBags
    mQty = TmpRs!SpillageQuantity
    
    Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "SPILLAGE", TmpRs!SpillageDate, mBags, mQty, 0, 0, Date, TmpRs!CDTFDate, "D")

    ProgressBar1.Value = mR
    TmpRs.MoveNext
Next mR
''---------------------------------------------------------


tmp = "Select SM_Prefix,GSLID,Sum(TransactionBags) 'Bag',Sum(TransactionQty) 'Qty' From Txn_GSL Group by SM_Prefix,GSLID Order by SM_Prefix,GSLID "
Call TmpTable



tmp = "Select * from Mst_Location"
Call Tmp2Table

ProgressBar1.Max = TmpRs.RecordCount + 1

For x = 1 To TmpRs.RecordCount
    tmp = "Select SM_Prefix +'-'+ Convert(Varchar,GSLID), * from Mst_GSL Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = " & TmpRs!GSLID & " "
    Call Tmp1Table
    'TmpRs1.MoveFirst
    'TmpRs1.Find "Key = '" & TmpRs!SM_Prefix & "-" & TmpRs!GSLID & "' "
    'If TmpRs1.EOF = False Then
    If TmpRs1.RecordCount > 0 Then
       If TmpRs!Bag <> TmpRs1!BalanceBags Then
          mGSLID = TmpRs1!GSLID
          mSM_Prefix = TmpRs1!SM_Prefix
          TmpRs2.MoveFirst
          TmpRs2.Find "SM_Prefix = '" & TmpRs1!SM_Prefix & "' "
          mLocationID = TmpRs2!LocationID
          mGodownID = IIf(IsNull(TmpRs1!GodownID), 0, TmpRs1!GodownID)
          mCommodityID = IIf(IsNull(TmpRs1!CommodityID), 0, TmpRs1!CommodityID)
          mClientIDRow = IIf(IsNull(TmpRs1!ClientIDRow), 0, TmpRs1!ClientIDRow)
          mExchangeTypeID = IIf(IsNull(TmpRs1!ExchangeTypeID), 0, TmpRs1!ExchangeTypeID)
          If TmpRs!Bag - TmpRs1!BalanceBags > 0 Then
             mBags = -(TmpRs!Bag - TmpRs1!BalanceBags)
             mQty = -(TmpRs!Qty - TmpRs1!BalanceWeight)
          ElseIf TmpRs!Bag - TmpRs1!BalanceBags < 0 Then
             mBags = Abs((TmpRs!Bag - TmpRs1!BalanceBags))
             mQty = Abs((TmpRs!Qty - TmpRs1!BalanceWeight))
          End If
          Call Add_TxnGSL(mGSLID, mLocationID, mSM_Prefix, mTansactionNo, mGodownID, mCommodityID, mClientIDRow, mExchangeTypeID, "AdjustmetDuringOpeningBalance", Date, mBags, mQty, 0, 0, Date, Date, "D")
       End If
    End If
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Max
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done."

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
TxtMonthYear.Value = Date
End Sub

Public Function Add_TxnGSL(mGSLID_ As Double, mLocationID_ As Double, mSM_Prefix_ As Variant, mTansactionNo_ As Double, mGodownID_ As Double, mCommodityID_ As Double, mClientIDRow_ As Double, mExchangeTypeID_ As Double, mTransactionType_ As Variant, mTransactionDate_ As Variant, mBags_ As Double, mQty_ As Double, mToGSL_ As Double, mFromGSL_ As Double, mTransactionEntryDate_ As Variant, mCDTFDate_ As Variant, mFlag_ As Variant)
    ''---  Add Entry in Txn_GSL
    RsTxn_GSL.AddNew
    ''RsTxn_GSL!Txn_GSL_ID = IDENTITY
    'RsTxn_GSL!Mst_GSL_ID = mGSLID_
    RsTxn_GSL!GSLID = mGSLID_
    RsTxn_GSL!LocationID = mLocationID_
    RsTxn_GSL!SM_Prefix = mSM_Prefix_
    If LCase(mTransactionType_) = LCase("CDTF") Then
       RsTxn_GSL!CxTFNO = mTansactionNo_ 'RsTxn_GSL!CxTF_DetailsID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("CITF") Then
       RsTxn_GSL!CxTFNO = mTansactionNo_ 'RsTxn_GSL!CxTF_DetailsID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("Adjustment Given") Then
       RsTxn_GSL!TxnID = mTansactionNo_ 'RsTxn_GSL!TxnGSLAdjID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("Adjustment Received") Then
       RsTxn_GSL!TxnID = mTansactionNo_ 'RsTxn_GSL!TxnGSLAdjID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("SALE") Then
       RsTxn_GSL!TxnID = mTansactionNo_ 'RsTxn_GSL!TxnGSLAdjID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("THEFT") Then
       RsTxn_GSL!TxnID = mTansactionNo_ 'RsTxn_GSL!TxnGSLAdjID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("FIRE") Then
       RsTxn_GSL!TxnID = mTansactionNo_ 'RsTxn_GSL!TxnGSLAdjID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("TRANSFER") Then
       RsTxn_GSL!GSLTransferID = mTansactionNo_ 'RsTxn_GSL!GSLTransferID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("MADEUP") Then
       RsTxn_GSL!SpillageID = mTansactionNo_ 'RsTxn_GSL!Txn_Spillage_ID = mTansactionNo_
    End If
    If LCase(mTransactionType_) = LCase("SPILLAGE") Then
       RsTxn_GSL!SpillageID = mTansactionNo_ 'RsTxn_GSL!Txn_Spillage_ID = mTansactionNo_
    End If
    
    RsTxn_GSL!GodownID = IIf(mGodownID_ = 0, Null, mGodownID_)
    RsTxn_GSL!CommodityID = IIf(mCommodityID_ = 0, Null, mCommodityID_)
    RsTxn_GSL!ClientID = IIf(mClientIDRow_ = 0, Null, mClientIDRow_)
    RsTxn_GSL!ExchangeTypeID = IIf(mExchangeTypeID_ = 0, Null, mExchangeTypeID_)
    RsTxn_GSL!TransactionType = mTransactionType_
    RsTxn_GSL!TransactionDate = IIf(IsNull(mTransactionDate_), Null, Format(mTransactionDate_, "yyyy-MM-dd"))
    RsTxn_GSL!TransactionBags = mBags_
    RsTxn_GSL!TransactionQty = mQty_
    RsTxn_GSL!ToGSLID = IIf(mToGSL_ = 0, Null, mToGSL_) 'RsTxn_GSL!TransferToMst_GSL_ID = IIf(mToGSL_ = 0, Null, mToGSL_)
    RsTxn_GSL!FromGSLID = IIf(mFromGSL_ = 0, Null, mFromGSL_) 'RsTxn_GSL!TransferFromMst_GSL_ID = IIf(mFromGSL_ = 0, Null, mFromGSL_)
    RsTxn_GSL!TransactionEntryDate = IIf(IsNull(mTransactionEntryDate_), Null, Format(mTransactionEntryDate_, "yyyy-MM-dd"))
    RsTxn_GSL!CDTFDate = IIf(IsNull(mCDTFDate_), Null, Format(mCDTFDate_, "yyyy-MM-dd"))
    RsTxn_GSL!TransactionFlag = mFlag_
    RsTxn_GSL!G_UID = GetGUID
    ''RsTxn_GSL!Concurrency = Now
    If IsNull(RsTxn_GSL!CreatedBy) = True Or RsTxn_GSL!CreatedBy = "" Then
       RsTxn_GSL!CreatedBy = Gen_UserLoginID
       RsTxn_GSL!CreatedDate = Now
    Else
       RsTxn_GSL!UpdatedBy = Gen_UserLoginID
       RsTxn_GSL!UpdatedDate = Now
    End If
    RsTxn_GSL.Update
End Function
