VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_NAVOutstandInvoices 
   Caption         =   "NAV Outstanding Invoices"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   9120
      Left            =   120
      TabIndex        =   3
      Top             =   735
      Width           =   15030
      Begin TabDlg.SSTab SSTab1 
         Height          =   8775
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   15478
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Mismatch Invoices"
         TabPicture(0)   =   "FrmRpt_NAVOutstandInvoices.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Invoices not found in Store-Man"
         TabPicture(1)   =   "FrmRpt_NAVOutstandInvoices.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Invoices not foud in NAV"
         TabPicture(2)   =   "FrmRpt_NAVOutstandInvoices.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSHFlexGrid3"
         Tab(2).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   8175
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   14420
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   8175
            Left            =   -74880
            TabIndex        =   6
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   14420
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   8175
            Left            =   -74880
            TabIndex        =   7
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   14420
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   1335
         TabIndex        =   2
         Top             =   165
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
         Left            =   120
         TabIndex        =   1
         Top             =   165
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmRpt_NAVOutstandInvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.Cols = 16

MSHFlexGrid1.TextMatrix(1, 0) = "Entry No"
MSHFlexGrid1.TextMatrix(1, 1) = "DocumentNo"
MSHFlexGrid1.TextMatrix(1, 2) = "VoucherNo"
MSHFlexGrid1.TextMatrix(1, 3) = "PostingDate"
MSHFlexGrid1.TextMatrix(1, 4) = "ClientNAV Code"
MSHFlexGrid1.TextMatrix(1, 5) = "Client Name"
MSHFlexGrid1.TextMatrix(1, 6) = "NAV Invoice Amount"
MSHFlexGrid1.TextMatrix(1, 7) = "NAV Balance Amount"
MSHFlexGrid1.TextMatrix(1, 8) = "SM Invoice Amount"
MSHFlexGrid1.TextMatrix(1, 9) = "Rec Amount"
MSHFlexGrid1.TextMatrix(1, 10) = "Rec TDS"
MSHFlexGrid1.TextMatrix(1, 11) = "Rec Other Deduction"
MSHFlexGrid1.TextMatrix(1, 12) = "Rec Amount (Batch Generated)"
MSHFlexGrid1.TextMatrix(1, 13) = "Rec TDS (Batch Generated)"
MSHFlexGrid1.TextMatrix(1, 14) = "Rec Other Deduction (Batch Generated)"
MSHFlexGrid1.TextMatrix(1, 15) = "SM Balance Amount"

MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 2200
MSHFlexGrid1.ColWidth(5) = 2200
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 2200
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1400
MSHFlexGrid1.ColWidth(15) = 1400

MSHFlexGrid1.RowHeight(1) = 600
MSHFlexGrid1.ColAlignment(9) = vbLeftJustify
MSHFlexGrid1.WordWrap = True
End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
Call Grid_Head1
Call Grid_Head2

tmp = "Select * from Txn_NAVInvoiceOutStanding Where CustomerNo Like '%CSWC%' Or"
tmp = tmp & " CustomerNo Like '%CSWD%' OR CustomerNo Like '%CSWO%'"  'CustomerPostingGroup in ('DP','SC','SO')"

Call Tmp1Table

tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0)+isnull(TI.AdjAmountFAG,0)RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And (TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo >0))RecTDS,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID and(TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo > 0))RecOtherDed"
tmp = tmp & " ,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
''tmp = tmp & " Where TI.EntryNo in (select distinct EntryNo from Txn_NAVInvoiceOutStanding Where CustomerNo Like '%CSWC%' Or"
''tmp = tmp & " CustomerNo Like '%CSWD%' OR CustomerNo Like '%CSWO%')"
tmp = tmp & " Order By TI.TIDate,TI.NAVCode"

Call TmpTable

tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,"
tmp = tmp & " isnull(Sum(TCRT.RecAmount),0) as RecAmount,"
tmp = tmp & " isnull(Sum(TCRT.RecTDS),0) as RecTDS,"
tmp = tmp & " isnull(Sum(TCRT.RecOtherDed), 0) As RecOtherDed,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName,TCRT.CRID,TCRT.SM_Prefix"
tmp = tmp & " From  Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " inner Join Txn_TIDetails TI on TI.TIID=TCRT.TIID"
tmp = tmp & " inner Join Txn_CRAdvanceAllocation CRAA on CRAA.CRAAID=TCRT.CRAAID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
tmp = tmp & " Where isnull(TCRT.EntryNo,0)=0 And TCR.ReturnFlag<>'Y' And isnull(CRAA.BatchID,0)>0"
tmp = tmp & " Group By TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode"
tmp = tmp & " ,MC.ClientID,MC.ClientName ,MD.DepositoryID,MD.DepositoryName,TCRT.CRID,TCRT.SM_Prefix"

Call Tmp2Table

tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.BatchID>0 And isnull(TCRT.EntryNo,0)= 0 And isnull(TCRT.CRAAID,0)=0)+isnull(TI.AdjAmountFAG,0)RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.BatchID>0 And isnull(TCRT.EntryNo,0) = 0 And isnull(TCRT.CRAAID,0)=0) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R' And TCR.BatchID>0 And TIID=TI.TIID And (isnull(TCRT.TDSEntryNo,0) = 0 And  isnull(TCRT.OtherDedEntryNo,0)=0))RecTDS,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.BatchID>0 And TIID=TI.TIID And isnull(TCRT.EntryNo,0) = 0 And isnull(TCRT.CRAAID,0)=0)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R' And TCR.BatchID>0 And TIID=TI.TIID and(isnull(TCRT.TDSEntryNo,0) = 0 and  isnull(TCRT.OtherDedEntryNo,0) = 0))RecOtherDed"
tmp = tmp & " ,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
''tmp = tmp & " Where TI.EntryNo in (select distinct EntryNo from Txn_NAVInvoiceOutStanding Where CustomerNo Like '%CSWC%' Or"
''tmp = tmp & " CustomerNo Like '%CSWD%' OR CustomerNo Like '%CSWO%')"
tmp = tmp & " Order By TI.TIDate,TI.NAVCode"

Call Tmp3Table

''tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,"
''tmp = tmp & " isnull(Sum(TCRT.RecAmount),0) as RecAmount,"
''tmp = tmp & " isnull(Sum(TCRT.RecTDS),0) as RecTDS,"
''tmp = tmp & " isnull(Sum(TCRT.RecOtherDed), 0) As RecOtherDed,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName,TCRT.CRID,TCRT.SM_Prefix"
''tmp = tmp & " From  Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " inner Join Txn_TIDetails TI on TI.TIID=TCRT.TIID"
''tmp = tmp & " inner Join Txn_CRAdvanceAllocation CRAA on CRAA.CRAAID=TCRT.CRAAID"
''tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
''tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
''tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
''tmp = tmp & " Where isnull(TCRT.EntryNo,0)=0 And TCR.ReturnFlag<>'Y' And isnull(CRAA.BatchID,0)>0"
''tmp = tmp & " Group By TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode"
''tmp = tmp & " ,MC.ClientID,MC.ClientName ,MD.DepositoryID,MD.DepositoryName,TCRT.CRID,TCRT.SM_Prefix"
''
''Call Tmp4Table

For I = 1 To TmpRs1.RecordCount
    
    If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
    TmpRs.Find "EntryNo=" & TmpRs1!EntryNo
    If TmpRs.EOF = False Then
       mTotRecAmt = TmpRs!RecAmount
       mTotRecTDS = TmpRs!RecTDS
       mTotRecOtherDed = TmpRs!RecOtherDed
       If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
       TmpRs2.Find "EntryNo=" & TmpRs1!EntryNo
       If TmpRs2.EOF = False Then
          mTotRecAmt = mTotRecAmt + TmpRs2!RecAmount
          mTotRecTDS = mTotRecTDS + TmpRs2!RecTDS
          mTotRecOtherDed = mTotRecOtherDed + TmpRs2!RecOtherDed
       End If
    
       If TmpRs1!RemainingAmount <> TmpRs!TIAmount - (mTotRecAmt + mTotRecTDS + mTotRecOtherDed) Then
          r = MSHFlexGrid1.Rows - 1
          MSHFlexGrid1.TextMatrix(r, 0) = TmpRs1!EntryNo
          MSHFlexGrid1.TextMatrix(r, 1) = TmpRs1!DocumentNo
          MSHFlexGrid1.TextMatrix(r, 2) = TmpRs1!VoucherNo
          MSHFlexGrid1.TextMatrix(r, 3) = TmpRs1!PostingDate
          MSHFlexGrid1.TextMatrix(r, 4) = TmpRs1!CustomerNo
          MSHFlexGrid1.TextMatrix(r, 5) = TmpRs1!CustomerName
          MSHFlexGrid1.TextMatrix(r, 6) = TmpRs1!InvoiceAMount
          MSHFlexGrid1.TextMatrix(r, 7) = TmpRs1!RemainingAmount
          MSHFlexGrid1.TextMatrix(r, 8) = TmpRs!TIAmount
          MSHFlexGrid1.TextMatrix(r, 9) = mTotRecAmt
          MSHFlexGrid1.TextMatrix(r, 10) = mTotRecTDS
          MSHFlexGrid1.TextMatrix(r, 11) = mTotRecOtherDed
          
          If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
          TmpRs3.Find "EntryNo=" & TmpRs1!EntryNo
          If TmpRs3.EOF = False Then
             MSHFlexGrid1.TextMatrix(r, 12) = TmpRs3!RecAmount
             MSHFlexGrid1.TextMatrix(r, 13) = TmpRs3!RecTDS
             MSHFlexGrid1.TextMatrix(r, 14) = TmpRs3!RecOtherDed
          End If
         
          MSHFlexGrid1.TextMatrix(r, 15) = Format(Val(MSHFlexGrid1.TextMatrix(r, 8)) - (Val(MSHFlexGrid1.TextMatrix(r, 9)) + Val(MSHFlexGrid1.TextMatrix(r, 10)) + Val(MSHFlexGrid1.TextMatrix(r, 11))), "#####0.00")
          MSHFlexGrid1.TextMatrix(r, 15) = Format(Val(MSHFlexGrid1.TextMatrix(r, 15)) - (Val(MSHFlexGrid1.TextMatrix(r, 12)) + Val(MSHFlexGrid1.TextMatrix(r, 13)) + Val(MSHFlexGrid1.TextMatrix(r, 14))), "#####0.00")
          
          MSHFlexGrid1.TextMatrix(0, 6) = Format(Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(r, 6)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 7) = Format(Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(r, 7)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 8) = Format(Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(r, 8)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 9) = Format(Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(r, 9)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 10) = Format(Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(r, 10)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 11) = Format(Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(r, 11)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(r, 12)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 13) = Format(Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(r, 13)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 14) = Format(Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(r, 14)), "#####0.00")
          MSHFlexGrid1.TextMatrix(0, 15) = Format(Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(r, 15)), "#####0.00")
          
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       End If
    Else
          If TmpRs1!RemainingAmount > 0 Then
            r = MSHFlexGrid2.Rows - 1
            MSHFlexGrid2.TextMatrix(r, 0) = TmpRs1!EntryNo
            MSHFlexGrid2.TextMatrix(r, 1) = TmpRs1!DocumentNo
            MSHFlexGrid2.TextMatrix(r, 2) = TmpRs1!VoucherNo
            MSHFlexGrid2.TextMatrix(r, 3) = TmpRs1!PostingDate
            MSHFlexGrid2.TextMatrix(r, 4) = TmpRs1!CustomerNo
            MSHFlexGrid2.TextMatrix(r, 5) = TmpRs1!CustomerName
            MSHFlexGrid2.TextMatrix(r, 6) = TmpRs1!InvoiceAMount
            MSHFlexGrid2.TextMatrix(r, 7) = TmpRs1!RemainingAmount
            'MSHFlexGrid2.TextMatrix(r, 8) = TmpRs!TIAmount
            'MSHFlexGrid2.TextMatrix(r, 9) = TmpRs!RecAmount
            'MSHFlexGrid2.TextMatrix(r, 10) = TmpRs!RecTDS
            'MSHFlexGrid2.TextMatrix(r, 11) = TmpRs!RecOtherDed
            'MSHFlexGrid2.TextMatrix(r, 12) = Val(MSHFlexGrid1.TextMatrix(r, 8)) - (Val(MSHFlexGrid1.TextMatrix(r, 9)) + Val(MSHFlexGrid1.TextMatrix(r, 10)) + Val(MSHFlexGrid1.TextMatrix(r, 11)))
            
            MSHFlexGrid2.TextMatrix(0, 6) = Format(Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(r, 6)), "#####0.00")
            MSHFlexGrid2.TextMatrix(0, 7) = Format(Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(r, 7)), "#####0.00")
            
            MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
         End If
    End If
    
    TmpRs1.MoveNext
Next

If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
For I = 1 To TmpRs.RecordCount
    
    mTotRecAmt = TmpRs!RecAmount
    mTotRecTDS = TmpRs!RecTDS
    mTotRecOtherDed = TmpRs!RecOtherDed
    If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
    TmpRs2.Find "EntryNo=" & TmpRs!EntryNo
    If TmpRs2.EOF = False Then
       mTotRecAmt = mTotRecAmt + TmpRs2!RecAmount
       mTotRecTDS = mTotRecTDS + TmpRs2!RecTDS
       mTotRecOtherDed = mTotRecOtherDed + TmpRs2!RecOtherDed
    End If
    If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
    TmpRs1.Find "EntryNo=" & TmpRs!EntryNo
       
    If TmpRs1.EOF = True Then
       If Format(TmpRs!TIAmount - (mTotRecAmt + mTotRecTDS + mTotRecOtherDed), "#####0.00") <> 0 Then
          r = MSHFlexGrid3.Rows - 1
          MSHFlexGrid3.TextMatrix(r, 0) = TmpRs!EntryNo
          MSHFlexGrid3.TextMatrix(r, 1) = TmpRs!PostInvoiceNo
          MSHFlexGrid3.TextMatrix(r, 2) = TmpRs!TINo
          MSHFlexGrid3.TextMatrix(r, 3) = TmpRs!TIDate
          MSHFlexGrid3.TextMatrix(r, 4) = TmpRs!NAVCode
          MSHFlexGrid3.TextMatrix(r, 5) = IIf(IsNull(TmpRs!ClientName), IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName), TmpRs!ClientName)
          MSHFlexGrid3.TextMatrix(r, 6) = TmpRs!TIAmount
          MSHFlexGrid3.TextMatrix(r, 7) = mTotRecAmt
          MSHFlexGrid3.TextMatrix(r, 8) = mTotRecTDS
          MSHFlexGrid3.TextMatrix(r, 9) = mTotRecOtherDed
          MSHFlexGrid3.TextMatrix(r, 10) = Format(Val(MSHFlexGrid3.TextMatrix(r, 6)) - (Val(MSHFlexGrid3.TextMatrix(r, 7)) + Val(MSHFlexGrid3.TextMatrix(r, 8)) + Val(MSHFlexGrid3.TextMatrix(r, 9))), "#####0.00")
        
          MSHFlexGrid3.TextMatrix(0, 6) = Format(Val(MSHFlexGrid3.TextMatrix(0, 6)) + Val(MSHFlexGrid3.TextMatrix(r, 6)), "#####0.00")
          MSHFlexGrid3.TextMatrix(0, 7) = Format(Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(r, 7)), "#####0.00")
          MSHFlexGrid3.TextMatrix(0, 8) = Format(Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(r, 8)), "#####0.00")
          MSHFlexGrid3.TextMatrix(0, 9) = Format(Val(MSHFlexGrid3.TextMatrix(0, 9)) + Val(MSHFlexGrid3.TextMatrix(r, 9)), "#####0.00")
          MSHFlexGrid3.TextMatrix(0, 10) = Format(Val(MSHFlexGrid3.TextMatrix(0, 10)) + Val(MSHFlexGrid3.TextMatrix(r, 10)), "#####0.00")
          
          MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       End If
    End If
    
    TmpRs.MoveNext
Next
CmdExcel.Enabled = True
End Sub

Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 13

MSHFlexGrid2.TextMatrix(1, 0) = "Entry No"
MSHFlexGrid2.TextMatrix(1, 1) = "DocumentNo"
MSHFlexGrid2.TextMatrix(1, 2) = "VoucherNo"
MSHFlexGrid2.TextMatrix(1, 3) = "PostingDate"
MSHFlexGrid2.TextMatrix(1, 4) = "ClientNAV Code"
MSHFlexGrid2.TextMatrix(1, 5) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 6) = "NAV Invoice Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "NAV Balance Amount"
MSHFlexGrid2.TextMatrix(1, 8) = "SM Invoice Amount"
MSHFlexGrid2.TextMatrix(1, 9) = "Rec Amount"
MSHFlexGrid2.TextMatrix(1, 10) = "Rec TDS"
MSHFlexGrid2.TextMatrix(1, 11) = "Rec Other Deduction"
MSHFlexGrid2.TextMatrix(1, 12) = "SM Balance Amount"

MSHFlexGrid2.ColWidth(0) = 1800
MSHFlexGrid2.ColWidth(1) = 1500
MSHFlexGrid2.ColWidth(2) = 1400
MSHFlexGrid2.ColWidth(3) = 1400
MSHFlexGrid2.ColWidth(4) = 2200
MSHFlexGrid2.ColWidth(5) = 2200
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 1400
MSHFlexGrid2.ColWidth(8) = 1400
MSHFlexGrid2.ColWidth(9) = 2200
MSHFlexGrid2.ColWidth(10) = 1400
MSHFlexGrid2.ColWidth(11) = 1400
MSHFlexGrid2.ColWidth(12) = 1400

MSHFlexGrid2.RowHeight(1) = 600
MSHFlexGrid2.ColAlignment(9) = vbLeftJustify
MSHFlexGrid2.WordWrap = True
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NAVOutstandingInvoices.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NAVOutstandingInvoices.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Mismatch Invoices"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Invoices Not found in SM"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid2.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "Invoices Not found in NAV"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid3.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
        End If
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Grid_Head2()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2
MSHFlexGrid3.Cols = 11

MSHFlexGrid3.TextMatrix(1, 0) = "Entry No"
MSHFlexGrid3.TextMatrix(1, 1) = "DocumentNo"
MSHFlexGrid3.TextMatrix(1, 2) = "VoucherNo"
MSHFlexGrid3.TextMatrix(1, 3) = "PostingDate"
MSHFlexGrid3.TextMatrix(1, 4) = "ClientNAV Code"
MSHFlexGrid3.TextMatrix(1, 5) = "Client Name"
MSHFlexGrid3.TextMatrix(1, 6) = "SM Invoice Amount"
MSHFlexGrid3.TextMatrix(1, 7) = "Rec Amount"
MSHFlexGrid3.TextMatrix(1, 8) = "Rec TDS"
MSHFlexGrid3.TextMatrix(1, 9) = "Rec Other Deduction"
MSHFlexGrid3.TextMatrix(1, 10) = "SM Balance Amount"

MSHFlexGrid3.ColWidth(0) = 1800
MSHFlexGrid3.ColWidth(1) = 1500
MSHFlexGrid3.ColWidth(2) = 1400
MSHFlexGrid3.ColWidth(3) = 1400
MSHFlexGrid3.ColWidth(4) = 2200
MSHFlexGrid3.ColWidth(5) = 2200
MSHFlexGrid3.ColWidth(6) = 1500
MSHFlexGrid3.ColWidth(7) = 1400
MSHFlexGrid3.ColWidth(8) = 1400
MSHFlexGrid3.ColWidth(9) = 2200
MSHFlexGrid3.ColWidth(10) = 1400

MSHFlexGrid3.RowHeight(1) = 600
MSHFlexGrid3.ColAlignment(9) = vbLeftJustify
MSHFlexGrid3.WordWrap = True

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
Call Grid_Head
Call Grid_Head1
Call Grid_Head2

End Sub
