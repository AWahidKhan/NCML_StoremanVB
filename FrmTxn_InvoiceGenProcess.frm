VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_InvoiceGenProcess 
   Caption         =   "Invoice Generation Process"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1485
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbNonNcdexClient 
         Height          =   315
         Left            =   6195
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   6105
      End
      Begin VB.OptionButton OptMonthWise 
         Caption         =   "Month Wise"
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
         Left            =   2520
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.OptionButton OptTxnWise 
         Caption         =   "Client Wise"
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
         Left            =   150
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   13590
         TabIndex        =   8
         Top             =   1065
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   495
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Left            =   12585
         TabIndex        =   7
         Top             =   1065
         Width           =   1000
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
         Left            =   11580
         TabIndex        =   6
         Top             =   1065
         Width           =   1000
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   3135
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   3000
      End
      Begin VB.CommandButton CmdShow 
         Caption         =   "Show"
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
         Left            =   10575
         TabIndex        =   5
         Top             =   1065
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1650
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   585
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
         Format          =   78249985
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   585
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
         Format          =   78249985
         CurrentDate     =   36494
      End
      Begin VB.Label Label2 
         Caption         =   "Client Name"
         Height          =   225
         Left            =   8640
         TabIndex        =   21
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label4 
         Caption         =   "Storage Rate Not Availabe in Storage Rate Master "
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
         Left            =   690
         TabIndex        =   18
         Top             =   1020
         Width           =   5535
      End
      Begin VB.Label Label36 
         Caption         =   "Location"
         Height          =   300
         Left            =   4290
         TabIndex        =   16
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   390
         TabIndex        =   15
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   1950
         TabIndex        =   14
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8265
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   15015
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   195
         TabIndex        =   9
         Top             =   195
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12726
         _Version        =   393216
         FixedCols       =   0
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
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   7800
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Caption         =   "Total Record Found :"
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
         Left            =   9120
         TabIndex        =   12
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label LblTotalRecord 
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
         Left            =   11520
         TabIndex        =   11
         Top             =   120
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmTxn_InvoiceGenProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mExchangeTypeID, mLocationID As Double
Dim mClientIDRow As Double
Dim mTxnType, mSM_Prefix As Variant

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 35
    .WordWrap = True
    
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Select"
      
    .TextMatrix(0, 1) = "GSL ID"
    .TextMatrix(0, 2) = "SM_Prefix"
    .TextMatrix(0, 3) = "Txn Type"
    .TextMatrix(0, 4) = "CxTF No"
    .TextMatrix(0, 5) = "CxTF Date"
    .TextMatrix(0, 6) = "LocationID"
    .TextMatrix(0, 7) = "Location Name"
    .TextMatrix(0, 8) = "CMPID"
    .TextMatrix(0, 9) = "CMP Name"
    .TextMatrix(0, 10) = "CommodityID"
    .TextMatrix(0, 11) = "Commodity"
    .TextMatrix(0, 12) = "Commodity GroupID"
    .TextMatrix(0, 13) = "Commodity Group"
    .TextMatrix(0, 14) = "Exchange TypeID"
    .TextMatrix(0, 15) = "Exchange Type"
    .TextMatrix(0, 16) = "ClientIDRow"
    .TextMatrix(0, 17) = "Client Name"
    .TextMatrix(0, 18) = "No Of Bags"
    .TextMatrix(0, 19) = "Gross Weight"
    .TextMatrix(0, 20) = "Demat Weight"
    .TextMatrix(0, 21) = "Flag"
    .TextMatrix(0, 22) = "Flag1"
    .TextMatrix(0, 23) = "Gross Wt"
    .TextMatrix(0, 24) = "Bag Size"
    .TextMatrix(0, 25) = "T Flag"
    .TextMatrix(0, 26) = "Billing CycleID"
    .TextMatrix(0, 27) = "Billing Cycle"
    .TextMatrix(0, 28) = "Billing Unit"
    .TextMatrix(0, 29) = "Rate On Qty"
    .TextMatrix(0, 30) = "Default Flag"
    .TextMatrix(0, 31) = "SRCID"
    .TextMatrix(0, 32) = "Rate on Bags"
    
    .TextMatrix(0, 33) = "CDTF No"
    .TextMatrix(0, 34) = "CDTF Date"
    
    
    .ColWidth(0) = 700
    .ColWidth(1) = 800
    .ColWidth(2) = 0
    .ColWidth(3) = 800
    .ColWidth(4) = 1000
    .ColWidth(5) = 1400
    .ColWidth(6) = 0
    .ColWidth(7) = 1400
    .ColWidth(8) = 0
    .ColWidth(9) = 1400
    .ColWidth(10) = 0
    .ColWidth(11) = 1200
    .ColWidth(12) = 0
    .ColWidth(13) = 1200
    .ColWidth(14) = 0
    .ColWidth(15) = 0 '1500
    .ColWidth(16) = 0
    .ColWidth(17) = 2600
    .ColWidth(18) = 1200
    .ColWidth(19) = 1200
    .ColWidth(20) = 1200
    .ColWidth(21) = 800
    .ColWidth(22) = 800
    .ColWidth(23) = 1000
    .ColWidth(24) = 1000
    .ColWidth(25) = 800
    .ColWidth(26) = 0
    .ColWidth(27) = 1200
    .ColWidth(28) = 1000
    .ColWidth(29) = 800
    .ColWidth(30) = 0
    .ColWidth(31) = 0
    .ColWidth(32) = 1000
    .ColWidth(33) = 1000
    .ColWidth(34) = 1000
    .RowHeight(0) = 500
    
End With
End Sub

Private Sub ChkCheckAll_Click()

If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.Col = 0
'       If MSHFlexGrid1.CellBackColor <> vbRed And MSHFlexGrid1.CellBackColor <> vbYellow Then
'          MSHFlexGrid1.Row = i
          MSHFlexGrid1.TextMatrix(I, 0) = strChecked
'       End If
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.Row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbTxnType_LostFocus()
If CmbTxnType.Text = "" Then Exit Sub

If LCase(CmbTxnType.Text) <> "deposit" And LCase(CmbTxnType.Text) <> "withdrawal" Then
   MsgBox "Invalid selection!!!"
   CmbTxnType.SetFocus
   Exit Sub
End If

If LCase(CmbTxnType.Text) = "deposit" Then
   mTxnType = "D"
ElseIf LCase(CmbTxnType.Text) = "withdrawal" Then
   mTxnType = "W"
End If

End Sub

Private Sub CmdClear_Click()
SFrom.Value = Date - 365
STo.Value = Date
CmbPreFix.Text = ""
LblTotalRecord.Caption = ""
ChkCheckAll.Value = 0
Call Grid_Head
OptTxnWise.Value = True
ProgressBar1.Value = 0
Call OptTxnWise_Click

End Sub

Private Sub CmdSave_Click()
Dim mTFlag As Variant

MsgBox "Wait till next message !!!!"
ProgressBar1.Value = 0
If MSHFlexGrid1.Rows <= 2 Then
   MsgBox "No data in grid!!!"
   Exit Sub
End If
tmp = ""
TYearMonth = Year(SFrom) & "-" & IIf(Month(SFrom) < 10, "0" & Month(SFrom), Month(SFrom)) & "-01"
If CmbPreFix.Text <> "" Then
   Call TableTxn_InvMonGSLDetail(" Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & TYearMonth & "' And (TFlag = 'D' Or TFlag = 'W') ")
   Call TableTxn_CxTF_GSL(" Where SM_Prefix = '" & mSM_Prefix & "' ")
Else
   Call TableTxn_InvMonGSLDetail(" Where YearMonth = '" & TYearMonth & "' And (TFlag = 'D' Or TFlag = 'W') ")
   Call TableTxn_CxTF_GSL
End If
tmp = "Select GSLID,SM_Prefix,GodownID,StackNo,LotNo,KRDate,NContractID,G_UID from Mst_GSL"
If CmbPreFix.Text <> "" Then
   tmp = tmp & " Where SM_PreFix='" & mSM_Prefix & "'"
End If
Call TmpTable

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows
For m = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(m, 0) = strChecked Then
       If LCase(MSHFlexGrid1.TextMatrix(m, 3)) = "deposit" Then
          mTFlag = "D"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(m, 3)) = "withdrawal" Then
          mTFlag = "W"
       End If
   
       With MSHFlexGrid1
            TGSLID = .TextMatrix(m, 1)
            tSM_Prefix = .TextMatrix(m, 2)
            TCMPID = .TextMatrix(m, 8)
            tCommodityID = .TextMatrix(m, 10)
            TBalanceBags = .TextMatrix(m, 18)
            TBalanceWeight = .TextMatrix(m, 20)
            TExchangeTypeID = .TextMatrix(m, 14)
            TClientIDRow = .TextMatrix(m, 16)
            tFlag = .TextMatrix(m, 21)
            TFlag1 = .TextMatrix(m, 22)
            TGrossWt = .TextMatrix(m, 19)
            TBagSize = .TextMatrix(m, 24)
            TYearMonth = Year(SFrom) & "-" & IIf(Month(SFrom) < 10, "0" & Month(SFrom), Month(SFrom)) & "-01"
            TCxTFNo = .TextMatrix(m, 4)
            TCxTFDate = CDate(.TextMatrix(m, 5))
            TCDTFNo = .TextMatrix(m, 33)
            TCDTFDate = CDate(.TextMatrix(m, 34))
       End With
       If TmpRs.RecordCount > 0 Then
          TmpRs.MoveFirst
       End If
       If CmbPreFix.Text <> "" Then
          TmpRs.Find "GSLID = " & TGSLID & ""
       Else
          TmpRs.Filter = "GSLID = " & TGSLID & " And SM_Prefix = '" & tSM_Prefix & "'"
       End If
       If TmpRs.EOF = False Then
          TGodownID = IIf(IsNull(TmpRs!godownid), "Null", TmpRs!godownid)
          TStackNo = IIf(IsNull(TmpRs!StackNo), "Null", TmpRs!StackNo)
          TLotNo = IIf(IsNull(TmpRs!LotNo), "Null", TmpRs!LotNo)
          TKRDate = TmpRs!KRDate
          TContractID = TmpRs!NContractID
          TG_UID = TmpRs!G_UID
       End If
        
       If CmbPreFix.Text = "" Then
          TmpRs.Filter = ""
       End If
       
       RsTxn_InvMonGSLDetail.AddNew
       RsTxn_InvMonGSLDetail!GSLID = TGSLID
       RsTxn_InvMonGSLDetail!SM_Prefix = tSM_Prefix
       RsTxn_InvMonGSLDetail!CMPID = TCMPID
       RsTxn_InvMonGSLDetail!CommodityID = tCommodityID
       RsTxn_InvMonGSLDetail!godownid = TGodownID
       RsTxn_InvMonGSLDetail!StackNo = TStackNo
       RsTxn_InvMonGSLDetail!LotNo = TLotNo
       RsTxn_InvMonGSLDetail!BalanceBags = Val(TBalanceBags)
       RsTxn_InvMonGSLDetail!BalanceWeight = Val(TBalanceWeight)
       RsTxn_InvMonGSLDetail!ExchangeTypeID = TExchangeTypeID
       RsTxn_InvMonGSLDetail!ClientIDRow = TClientIDRow
       RsTxn_InvMonGSLDetail!Flag = tFlag
       RsTxn_InvMonGSLDetail!Flag1 = TFlag1
       RsTxn_InvMonGSLDetail!tFlag = mTFlag
       RsTxn_InvMonGSLDetail!GrossWt = TGrossWt
       If IsNull(TKRDate) = False Then
          RsTxn_InvMonGSLDetail!KRDate = TKRDate
       End If
       If IsNull(TContractID) = False Then
          RsTxn_InvMonGSLDetail!ContractID = TContractID
       End If
       RsTxn_InvMonGSLDetail!BagSize = Val(TBagSize)
       RsTxn_InvMonGSLDetail!YearMonth = TYearMonth
       RsTxn_InvMonGSLDetail!G_UID = TG_UID
       RsTxn_InvMonGSLDetail!CxTFDate = TCxTFDate
       RsTxn_InvMonGSLDetail!CdtfNo = TCDTFNo
       RsTxn_InvMonGSLDetail!CDTFDate = TCDTFDate
       RsTxn_InvMonGSLDetail!CxTFNo = TCxTFNo
       RsTxn_InvMonGSLDetail.Update
         
       RsTxn_InvMonGSLDetail.Filter = ""
          
       If RsTxn_CxTF_GSL.RecordCount > 0 Then
          RsTxn_CxTF_GSL.MoveFirst
       End If
        
       If CmbPreFix.Text <> "" Then
          RsTxn_CxTF_GSL.Filter = "GSLID = " & TGSLID & " And TxnType = '" & mTFlag & "' And CxTFNo = " & TCxTFNo & ""
       Else
          RsTxn_CxTF_GSL.Filter = "GSLID = " & TGSLID & " And SM_Prefix = '" & tSM_Prefix & "'"
       End If
        
       For x = 1 To RsTxn_CxTF_GSL.RecordCount
           RsTxn_CxTF_GSL!BillFlag = "I"
           RsTxn_CxTF_GSL.Update
           RsTxn_CxTF_GSL.MoveNext
       Next
       
       'If CmbPreFix.Text = "" Then
       RsTxn_CxTF_GSL.Filter = ""
       'End If
   End If
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Data Saved Successfully !!!"


Call Grid_Head

End Sub

Private Sub CmdShow_Click()
Dim mCMPLockedDate, mMonth, mYear, mTotalDays As Variant
Dim iu As Double

Call Grid_Head
ChkCheckAll.Value = 0

If STo.Value < SFrom.Value Then
   MsgBox "To date must be always greater than from date!!!"
   STo.SetFocus
   Exit Sub
End If

'If CmbPreFix.Text = "" Then
'   MsgBox "Please select atleast one location!!!"
'   CmbPreFix.SetFocus
'   Exit Sub
'End If

MsgBox "Wait till next message !!!!"

'If OptMonthWise.Value = True Then
'   tmp = "Select Max(CMPLockedDate) as 'CMPLockedDate' From Mst_CMP Where LocationID = " & mLocationID & ""
'   Call TmpTable
'
'   If TmpRs.RecordCount > 0 Then
'      mCMPLockedDate = TmpRs!CMPLockedDate
'      mMonth = Month(mCMPLockedDate)
'      mYear = Year(mCMPLockedDate)
'      If mMonth = 1 Then
'         mMonth = 12
'         mYear = mYear - 1
'      Else
'         mMonth = mMonth - 1
'      End If
'      mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
'      SFrom.Value = mMonth & "/01/" & mYear
'      STo.Value = mMonth & "/" & mTotalDays & "/" & mYear
'   End If
'
'End If


'''tmp = "Select TCG.GSLID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTFNo,TCD.CxTFDate,MCMP.LocationID,ML.LocationName,TCD.CMPID,MCMP.CMPName, "
'''tmp = tmp & " TCD.CommodityID,MC.Commodity,MC.CommodityGroupID,MCG.CommodityGroup,"
'''tmp = tmp & " TCD.ExchangeTypeID,ME.ExchangeType,MClient.ClientIDRow,MClient.ClientName,TCG.NoOfBags,TCG.GrossWeight,TCG.DematWeight,MGSL.Flag,MGSL.Flag1,MGSL.GrossWt,"
'''tmp = tmp & " round((TCG.DematWeight/Nullif(TCG.NoOfBags,0))*1000,2) as 'BagSize','O' TFlag,MSRC.BillingCycleID,MBC.BillingCycle,MSRC.Flag,MSRC.RateOnQty,MSRC.DefaultRate,MSRC.SRCID,MGSL.CDTFNo,MGSL.CDTFDate"
'''tmp = tmp & " From Txn_CxTF_GSL  TCG"
'''tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
'''tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'''tmp = tmp & " Inner Join Mst_Commodity MC On TCD.CommodityID = MC.CommodityID"
'''tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
'''tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
'''tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
'''tmp = tmp & " Inner Join Mst_Client MClient On MGSL.ClientIDRow = MClient.ClientIDRow"
'''tmp = tmp & " Inner Join Mst_ExchangeType ME On TCD.ExchangeTypeID = ME.ExchangeTypeID"
'''tmp = tmp & " Left Join Mst_StorageRateCard MSRC On  TCD.ExchangeTypeID = MSRC.ExchangeTypeID And MC.CommodityGroupID = MSRC.CommodityGroupID And MCMP.LocationID = MSRC.LocationID"
'''tmp = tmp & " Left Join Mst_BillingCycle MBC On MSRC.BillingCycleID = MBC.BillingCycleID "

ProgressBar1.Value = 0

tmp = "Select TCG.GSLID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTFNo,TCD.CxTFDate,MCMP.LocationID,"
tmp = tmp & " ML.LocationName,TCD.CMPID,MCMP.CMPName,TCD.CommodityID,MC.Commodity,"
tmp = tmp & " MC.CommodityGroupID,MCG.CommodityGroup,TCD.ExchangeTypeID,ME.ExchangeType,"
tmp = tmp & " MClient.ClientIDRow,MClient.ClientName,TCG.NoOfBags,TCG.GrossWeight,"
tmp = tmp & " TCG.DematWeight,MGSL.Flag,MGSL.Flag1,MGSL.GrossWt,"
'tmp = tmp & " round((TCG.DematWeight/Nullif(TCG.NoOfBags,0))*1000,2) as 'BagSize','O' TFlag,MGSL.CDTFNo,MGSL.CDTFDate"
tmp = tmp & " MGSL.BagSize,'O' TFlag,MGSL.CDTFNo,MGSL.CDTFDate"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_Commodity MC On TCD.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MClient On MGSL.ClientIDRow = MClient.ClientIDRow"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TCD.ExchangeTypeID = ME.ExchangeTypeID"

wc = " Where MGSL.SpillageFlag = 0   "
wc = wc & "And TCD.ExchangeTypeID = 1 And (TCG.BillFlag is Null or TCG.BillFlag = '') "

If Gen_DBType = "MDB" Then
   wc = wc & " and TCD.CxTFDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   wc = wc & " and TCD.CxTFDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If

If CmbPreFix.Text <> "" Then
   wc = wc & " And ML.LocationName='" & CmbPreFix.Text & "' "
End If

If CmbNonNcdexClient.Text <> "" Then
   wc = wc & " And MClient.ClientName = '" & CmbNonNcdexClient.Text & "'"
End If

tmp1 = " Order By TCG.CxTFNo,TCG.SM_Prefix,TCD.CxTFDate,TCG.TxnType"

tmp = tmp & " " & wc & " " & tmp1


Call TmpTable

Call Grid_Head

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 1

For iu = 1 To TmpRs.RecordCount
    
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    
    For c = 1 To 25 'MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c - 1)), "", TmpRs.Fields(c - 1))
    Next
    
    MSHFlexGrid1.TextMatrix(iu, 5) = Format(MSHFlexGrid1.TextMatrix(iu, 5), "dd-mmm-yyyy")
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 3)) = "d" Then
       MSHFlexGrid1.TextMatrix(iu, 3) = "Deposit"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 3)) = "w" Then
       MSHFlexGrid1.TextMatrix(iu, 3) = "Withdrawal"
    End If
    
    
    
''' MSHFlexGrid1.TextMatrix(iu, 31) = IIf(IsNull(TmpRs!SRCID), "", TmpRs!SRCID)
    
    
    tmp = "Select MBC.BillingCycleID,MSPRC.BillingUnit,MSPRCD.SPRCID,MSPRC.SRCID,MSPRC.RateOnQty,MSPRC.BaseRate,"
    tmp = tmp & " MSPRC.AbsoluteDiscount, MSPRC.Amount, MSPRCD.FromBagSize, MSPRCD.ToBagSize, MSPRCD.RatePerBag,"
    tmp = tmp & " MBC.BillingCycle,MBC.NoofDays"
    tmp = tmp & " from Mst_SpecialRateCard MSPRC"
    tmp = tmp & " Left Join Mst_SpecialRateCardDetail MSPRCD On MSPRC.SPRCID=MSPRCD.SPRCID"
    tmp = tmp & " Inner Join Mst_BillingCycle MBC on MSPRC.BillingCycleID= MBC.BillingCycleID"
    tmp = tmp & " Where MSPRC.LocationID= " & TmpRs!LocationID & " And MSPRC.CommodityGroupID=" & TmpRs!CommodityGroupID & " "
    tmp = tmp & " And MSPRC.ExchangeTypeID=1 And MSPRC.ClientIDRow=" & TmpRs!ClientIDRow & " "
    
    Call Tmp1Table
    
    If TmpRs1.RecordCount > 0 Then
    
       MSHFlexGrid1.TextMatrix(iu, 26) = IIf(IsNull(TmpRs1!BillingCycleID), "", TmpRs1!BillingCycleID)
       MSHFlexGrid1.TextMatrix(iu, 27) = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
       MSHFlexGrid1.TextMatrix(iu, 28) = IIf(IsNull(TmpRs1!BillingUnit), "", TmpRs1!BillingUnit)
       MSHFlexGrid1.TextMatrix(iu, 31) = IIf(IsNull(TmpRs1!SRCID), "", TmpRs1!SRCID)
    
       If TmpRs1!BillingUnit = "B" And TmpRs1!BaseRate = "n" Then
          MSHFlexGrid1.TextMatrix(iu, 32) = TmpRs1!RatePerBag
       ElseIf (TmpRs1!BillingUnit = "B" And TmpRs1!BaseRate = "y") Then
          ' Find rate from Storage rate Card Detail
          If RsMst_StorageRateCardDetail.RecordCount > 0 Then
             RsMst_StorageRateCardDetail.MoveFirst
          End If
              
          RsMst_StorageRateCardDetail.Find "SRCDID=" & TmpRs1!SRCDID
             
          If Not RsMst_StorageRateCardDetail.EOF Then
             MSHFlexGrid1.TextMatrix(iu, 32) = RsMst_StorageCardRateDetail!RatePerBag
             If TmpRs1!AbsoluteDiscount = "Percentage" Then
                MSHFlexGrid1.TextMatrix(iu, 32) = Val(MSHFlexGrid1.TextMatrix(iu, 32)) - (Val(MSHFlexGrid1.TextMatrix(iu, 32)) * (TmpRs1!Amount / 100))
             ElseIf TmpRs1!AbsoluteDiscount = "Actual" Then
                MSHFlexGrid1.TextMatrix(iu, 32) = Val(MSHFlexGrid1.TextMatrix(iu, 32)) - Val(TmpRs1!Amount)
             End If
          End If
          ''For Quantity
          ElseIf TmpRs1!BillingUnit = "Q" And TmpRs1!BaseRate = "n" Then
              MSHFlexGrid1.TextMatrix(iu, 29) = TmpRs1!Amount
          ElseIf TmpRs1!BillingUnit = "Q" And TmpRs!BaseRate = "y" Then
              If TmpRs1!AbsoluteDiscount = "Actual" Then
                 MSHFlexGrid1.TextMatrix(iu, 29) = TmpRs1!RateOnQty - TmpRs1!Amount
              ElseIf TmpRs!AbsoluteDiscount = "Percentage" Then
                 MSHFlexGrid1.TextMatrix(iu, 29) = TmpRs1!RateOnQty - (TmpRs1!RateOnQty * (TmpRs1!Amount / 100))
              End If
          End If
   Else
      
      '2. Check in Storage Rate Card
       tmp = "Select MSRC.SRCID,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag,MSRC.Flag,MBC.BillingCycle,"
       tmp = tmp & " MBC.NoofDays,MSRC.RateOnQty,MBC.BillingCycleID"
       tmp = tmp & " From Mst_StorageRateCard MSRC"
       tmp = tmp & " Left Join Mst_StoragerateCardDetail MSRCD on MSRC.SRCID= MSRCD.SRCID"
       tmp = tmp & " Inner Join Mst_CommodityGroup MCG on MCG.CommodityGroupID=MSRC.CommodityGroupID"
       tmp = tmp & " Inner Join Mst_BillingCycle MBC on MSRC.BillingCycleID= MBC.BillingCycleID"
       tmp = tmp & " Where   MSRC.ExchangeTypeID = 1 And MSRC.DefaultRate=1 "
       tmp = tmp & " And  MSRC.LocationID=" & TmpRs!LocationID & "  "
       tmp = tmp & " And MCG.CommodityGroupID=" & TmpRs!CommodityGroupID & " "
       Call Tmp1Table

       If TmpRs1.RecordCount > 0 Then
          
          If TmpRs1!Flag = "B" Then
             MSHFlexGrid1.TextMatrix(iu, 32) = GetRate(TmpRs1!SRCID, TmpRs!BagSize)
          ElseIf TmpRs1!Flag = "Q" Then
             MSHFlexGrid1.TextMatrix(iu, 29) = TmpRs1!RateOnQty
          End If
             
          MSHFlexGrid1.TextMatrix(iu, 26) = IIf(IsNull(TmpRs1!BillingCycleID), "", TmpRs1!BillingCycleID)
          MSHFlexGrid1.TextMatrix(iu, 27) = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
          MSHFlexGrid1.TextMatrix(iu, 28) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
          MSHFlexGrid1.TextMatrix(iu, 31) = IIf(IsNull(TmpRs1!SRCID), "", TmpRs1!SRCID)
       End If

   End If
    
'''    If LCase(MSHFlexGrid1.TextMatrix(iu, 28)) = "bags" Then
'''       Call AddStorageRateForBags(TmpRs!SRCID, IIf(IsNull(TmpRs!BagSize), 0, TmpRs!BagSize), iu)
'''    Else
'''       MSHFlexGrid1.TextMatrix(iu, 32) = 0
'''    End If
'''    If MSHFlexGrid1.TextMatrix(iu, 32) = "" Then
'''       MSHFlexGrid1.TextMatrix(iu, 32) = 0
'''    End If
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 28)) = "b" Then
       MSHFlexGrid1.TextMatrix(iu, 28) = "Bags"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 28)) = "q" Then
       MSHFlexGrid1.TextMatrix(iu, 28) = "Quantity"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 28)) = "s" Then
       MSHFlexGrid1.TextMatrix(iu, 28) = "Sq Ft"
    End If
    
    MSHFlexGrid1.TextMatrix(iu, 33) = TmpRs!CdtfNo
    MSHFlexGrid1.TextMatrix(iu, 34) = TmpRs!CDTFDate
    MSHFlexGrid1.TextMatrix(iu, 34) = Format(MSHFlexGrid1.TextMatrix(iu, 34), "dd-mmm-yyyy")
    
    MSHFlexGrid1.Row = iu
    MSHFlexGrid1.Col = 0
    MSHFlexGrid1.CellFontName = "Wingdings"
    MSHFlexGrid1.CellFontSize = 12
    MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
    MSHFlexGrid1.Text = strUnChecked
    If MSHFlexGrid1.TextMatrix(iu, 27) = "" Then
       MSHFlexGrid1.Row = iu
       For o = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.Col = o
           MSHFlexGrid1.CellBackColor = vbRed
       Next
    End If
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max

LblTotalRecord.Refresh
LblTotalRecord.Caption = TmpRs.RecordCount
MsgBox "Done!!!"

End Sub
Private Function GetRate(mSRCID_ As Integer, mBagSize_ As Double)
Dim Retval As Double
Retval = 0
tmp = " Select RatePerBag from Mst_StorageRateCardDetail "
tmp = tmp & " Where SRCID = " & mSRCID_ & " And  FromBagSize <= " & mBagSize_ & " "
tmp = tmp & "And ToBagSize >= " & mBagSize_ & " "
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   Retval = TmpRs2!RatePerBag
End If
GetRate = Retval
End Function
Private Sub AddStorageRateForBags(mSRCID_ As Variant, mBagSize_ As Double, x1 As Double)

tmp = " Select RatePerBag from Mst_StorageRateCardDetail "
tmp = tmp & " Where SRCID = " & mSRCID_ & " And  FromBagSize <= " & mBagSize_ & " "
tmp = tmp & "And ToBagSize >= " & mBagSize_ & " "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For c1 = 32 To 32
       MSHFlexGrid1.TextMatrix(x1, c1) = IIf(IsNull(TmpRs1.Fields(c1 - 32)), 0, TmpRs1.Fields(c1 - 32))
   Next
End If
TmpRs1.Close
End Sub
Private Sub Form_Load()

'Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   'Call GButtonLockAD(Me)
'   Frame0.Enabled = False
'   Frame2.Enabled = False
'   MsgBox "Access denied !!!!!!!!!"
'   Exit Sub
'End If


Frame0.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click

End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   mLocationID = 0
   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbPreFix.SetFocus
   Exit Sub
End If

If tmp <> CmbPreFix.Text Then
'   CmbSCmp.Text = ""
   mCMPID = Null
End If
mLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!LocationName
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
    If .Col = 0 Then  '.Row = 1 And
       If .TextMatrix(.Row, .Col) = "" Then Exit Sub
          .Col = 0
       If (.CellBackColor) <> vbRed And (.CellBackColor) <> vbYellow Then
          If .TextMatrix(.Row, .Col) = strUnChecked Then
             .TextMatrix(.Row, .Col) = strChecked
          Else
             .TextMatrix(.Row, .Col) = strUnChecked
          End If
       Else
          If .CellBackColor = vbRed Then
             MsgBox "Storage Rate Not Availabe in Storage Rate Master !!!"
          ElseIf .CellBackColor = vbYellow Then
             MsgBox "Storage Rate is not Default Rate!!!"
          End If
       End If
    End If
End With

End Sub
Private Sub OptMonthWise_Click()
Call OptTxnWise_Click
End Sub

Private Sub OptTxnWise_Click()
If OptTxnWise.Value = True Then
   SFrom.Enabled = True
   STo.Enabled = True
   SFrom.Value = Date - 365
   STo.Value = Date
   CmbNonNcdexClient.Text = ""
   CmbNonNcdexClient.Enabled = True
ElseIf OptMonthWise.Value = True Then
   SFrom.Enabled = False
   STo.Enabled = False
   CmbNonNcdexClient.Text = ""
   CmbNonNcdexClient.Enabled = False
End If
Call ClearClientData
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
'----
Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   mExchangeTypeID = 0
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeType.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   Call ClearClientData
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbNonNcdexClient.Width = 10095
   CmbNonNcdexClient.Top = TxtClientID.Top
   CmbNonNcdexClient.Left = TxtClientID.Left
   Label9.Visible = False
   Label10.Visible = False
   Label11.Visible = False
   CmbNonNcdexClient.Visible = True
Else
   Label9.Visible = True
   Label10.Visible = True
   Label11.Visible = True
   CmbNonNcdexClient.Visible = False
End If

End Sub


Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub
Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow

End Sub

Private Sub ClearClientData()

'TxtClientID = ""
'TxtClientName = ""
'TxtDPID = ""
'TxtDPName = ""
'CmbClientDP.Text = ""
'CmbClientDP.Visible = False

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "InvoiceGenProcess.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "InvoiceGenProcess.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    x = 0
    For c = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, x + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, x + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, x + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
        
        MSHFlexGrid1.Row = I
        MSHFlexGrid1.Col = x
        
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, x + 1), oWS.Cells(mRow, x + 1)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, x + 1), oWS.Cells(mRow, x + 1)).Interior.Color = vbYellow
        End If
        x = x + 1
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
