VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_ControlReport 
   Caption         =   "Control Report"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   14895
      Begin VB.OptionButton OptGSLvsTxn 
         Caption         =   "GSL Vs Txn: Exchange Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3375
      End
      Begin VB.OptionButton OptCDCGSLDetail 
         Caption         =   "CDC Vs GSL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   3
         Top             =   120
         Width           =   2295
      End
      Begin VB.OptionButton OptRRN 
         Caption         =   "RRN Vs CITF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
      Begin VB.OptionButton OptDeliveryQty 
         Caption         =   "WHR Vs CITF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   120
         Width           =   2295
      End
      Begin VB.OptionButton OptGrossWt 
         Caption         =   "TTD-GCTD Difference"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   13515
         TabIndex        =   7
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   12300
         TabIndex        =   6
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   11085
         TabIndex        =   5
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   9870
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8295
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   14631
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_ControlReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClear_Click()
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()
If MSHFlexGrid1.TextMatrix(1, 0) <> "" Then
   Gen_mTimeStamp = GetTimeStamp
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ControlReport.xls")
   Call Xls_Detail_Rpt
Else
   MsgBox "No Record To Export"
End If
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ControlReport.xls")
Set oWS = oWB.Worksheets("Sheet1")
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
        
        MSHFlexGrid1.Row = I
        MSHFlexGrid1.Col = c
        
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbRed
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetReport_Click()
Call Grid_Head
If OptGrossWt.Value = True Then
   tmp = "Select A.SM_Prefix,A.CxTFNO,A.txnType, A.[GSL-Gross-Wt], B.[TTD-Gross-Wt]" & _
         " From Vw_CxTFGSL A" & _
         " INNER JOIN Vw_TruckDetail B ON A.SM_Prefix = B.SM_Prefix AND A.CxTFNO = B.CxTFNO AND A.txnType = B.txnType" & _
         " Where a.[GSL-Gross-Wt] <> b.[TTD-Gross-Wt]" & _
         " Order BY  A.SM_Prefix,A.CxTFNO,A.txnType"
   Call TmpTable
ElseIf OptDeliveryQty.Value = True Then
   tmp = "Select  a.SM_Prefix,MC.ClientName,b.WHRNO,a.TxnQty,a.TxnBags,b.DeliveredQuantity,b.DeliveredNoOfBags" & _
         " From Vw_CxTFDetails a" & _
         " Inner Join Vw_Whr_detail b on a.SM_Prefix=b.SM_Prefix and a.clientIDRow=b.ClientIDRow and a.RRNID=b.WHRNO" & _
         " inner Join Mst_Client MC on a.ClientIDRow=MC.ClientIDRow " & _
         " Where a.TxnQty <> b.DeliveredQuantity Or a.TxnBags <> b.DeliveredNoOfBags" & _
         " Order by a.SM_Prefix,a.ClientIDRow,a.RRNID,b.WHRNO"
   Call TmpTable
ElseIf OptRRN.Value = True Then
        tmp = "Select b.SM_Prefix,MC.ClientName,MR.RRNNO,MR.ISINID,b.TxnQty,a.[Del Qty]" & _
        " From Vw_RRN a" & _
        " Inner Join Vw_ClientDetail b on a.RRNID=b.RRNID and a.ClientIDROW=b.ClientIDRow" & _
        " Inner Join Mst_RRN MR on a.RRNID=MR.RRNID" & _
        " inner Join Mst_Client MC on a.ClientIDRow=MC.ClientIDRow " & _
        " Where b.TxnQty <> a.[Del Qty]" & _
        " Order By  b.SM_Prefix,a.RRNID,a.ClientIDRow"
   Call TmpTable
ElseIf OptCDCGSLDetail.Value = True Then
   tmp = "Select b.SM_Prefix,MC.ClientName,a.CDCNO,b.TxnQty,a.[Del Qty]" & _
         " from Vw_GSLDetail a" & _
         " Inner Join Vw_ClientDetailForGslDetail b on  a.ClientIDROW=b.ClientIDRow and a.SM_Prefix=b.SM_Prefix" & _
          " inner Join Mst_Client MC on a.ClientIDRow=MC.ClientIDRow " & _
         " Where b.TxnQty <> a.[Del Qty]" & _
         " Order By  b.SM_Prefix,a.CDCNO,a.ClientIDRow"
         '
   Call TmpTable
ElseIf OptGSLvsTxn.Value = True Then
'   tmp = "Select a.GSLID,a.SM_prefix,c.CxTFNo,c.TXnType,a.flag,a.balancebags,a.balanceweight,a.ExchangeTypeID as 'GSLExchange', " & _
'         " c.ExchangeTypeID as 'Txn Exchange', a.CommodityID as 'GSL Commodity',c.CommodityID as 'Txn Commodity',  " & _
'         " a.CMPID as 'GSL CMP', c.CMPID as 'Txn CMP'  " & _
'         " From Mst_GSL a  " & _
'         " Inner Join Txn_CxTF_GSL b On a.GSLID = b.GSLID And a.SM_Prefix = b.SM_Prefix  " & _
'         " Inner Join Txn_CxTF_Details c On b.CxTFNo = c.CxTFNo And b.SM_Prefix = c.SM_prefix And b.TxnType = c.TxnType  " & _
'         " Where (a.ExchangeTypeID <> c.ExchangeTypeID Or a.CommodityID <> c.CommodityID Or a.CMPID <> c.CMPID ) And SpillageFlag=0 " & _
'         " Order By a.GSLID,a.SM_Prefix,c.CxTFNo,c.TxnType "
        
        tmp = "Select a.GSLID,a.SM_prefix,c.CxTFNo,c.TXnType,a.flag,a.balancebags,a.balanceweight,MEx.ExchangeType as 'GSLExchange'," & _
        " MEx.ExchangeType as 'Txn Exchange', MC.Commodity as 'GSL Commodity',MC1.Commodity as 'Txn Commodity'," & _
        " CMP.CMPName as 'GSL CMP', CMP1.CMPName as 'Txn CMP'" & _
        " From Mst_GSL a " & _
        " Inner Join Txn_CxTF_GSL b On a.GSLID = b.GSLID And a.SM_Prefix = b.SM_Prefix" & _
        " Inner Join Txn_CxTF_Details c On b.CxTFNo = c.CxTFNo And b.SM_Prefix = c.SM_prefix And b.TxnType = c.TxnType" & _
        " Inner Join Mst_ExchangeType MEx on a.ExchangeTypeID=MEx.ExchangeTypeID" & _
        " Inner Join Mst_ExchangeType MEx1 on c.ExchangeTypeID=MEx1.ExchangeTypeID" & _
        " Inner Join Mst_Commodity MC on a.CommodityID=MC.CommodityID" & _
        " Inner Join Mst_Commodity MC1 on c.CommodityID=MC1.CommodityID" & _
        " Inner Join Mst_CMP CMP on a.CMPID=CMP.CMPID" & _
        " Inner Join Mst_CMP CMP1 on c.CMPID=CMP1.CMPID" & _
        " Where (a.ExchangeTypeID <> c.ExchangeTypeID Or a.CommodityID <> c.CommodityID Or a.CMPID <> c.CMPID) And SpillageFlag = 0" & _
        " Order By a.GSLID,a.SM_Prefix,c.CxTFNo,c.TxnType"

   Call TmpTable
End If

If TmpRs.RecordCount > 0 Then
   For r = 1 To TmpRs.RecordCount
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(r, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       If OptDeliveryQty.Value = True Then
          MSHFlexGrid1.TextMatrix(r, 2) = GetWHRNo(MSHFlexGrid1.TextMatrix(r, 2))
       End If
       If OptGSLvsTxn.Value = True Then
          If MSHFlexGrid1.TextMatrix(r, 7) <> MSHFlexGrid1.TextMatrix(r, 8) Then '= "GSL Exchange"
             MSHFlexGrid1.Col = 7
             MSHFlexGrid1.Row = r
             MSHFlexGrid1.CellBackColor = vbRed
          ElseIf MSHFlexGrid1.TextMatrix(r, 9) <> MSHFlexGrid1.TextMatrix(r, 10) Then '= "GSL Commodity"
             MSHFlexGrid1.Col = 9
             MSHFlexGrid1.Row = r
             MSHFlexGrid1.CellBackColor = vbRed
          ElseIf MSHFlexGrid1.TextMatrix(r, 11) <> MSHFlexGrid1.TextMatrix(r, 12) Then ' = "GSL CMP"
             MSHFlexGrid1.Col = 11
             MSHFlexGrid1.Row = r
             MSHFlexGrid1.CellBackColor = vbRed
          End If
       End If
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
Else
   MsgBox "Record Not Found"
End If
End Sub
Private Function GetWHRNo(mWhrNo_ As Variant) As String
Dim Retval As String

Retval = ""
tmp = "Select CMPID,WHRNo from Txn_WHR_Detail Where WHRNo=" & mWhrNo_ & ""
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs1!CMPID), 0, TmpRs1!CMPID) & "-" & TmpRs1!WHRNO - (IIf(IsNull(TmpRs1!CMPID), 0, TmpRs1!CMPID) * 1000000)
End If
GetWHRNo = Retval
End Function

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .WordWrap = True
    If OptGrossWt.Value = True Then
       .Cols = 5
       .TextMatrix(0, 0) = "Location"
       .TextMatrix(0, 1) = "CxTF NO"
       .TextMatrix(0, 2) = "Transaction Type"
       .TextMatrix(0, 3) = "GSL-Gross-Wt"
       .TextMatrix(0, 4) = "TruckDetail-Gross-Wt"
       .ColWidth(0) = 1500
       .ColWidth(1) = 1500
       .ColWidth(2) = 2000
       .ColWidth(3) = 2000
       .ColWidth(4) = 2500
    ElseIf OptDeliveryQty.Value = True Then
       .Cols = 7
       .TextMatrix(0, 0) = "Location"
       .TextMatrix(0, 1) = "Client Name"
       .TextMatrix(0, 2) = "WHR No"
       .TextMatrix(0, 3) = "Transaction Qty"
       .TextMatrix(0, 4) = "Transaction Bags"
       .TextMatrix(0, 5) = "Delivered Quantity"
       .TextMatrix(0, 6) = "Delivered NoOfBags"
       .ColWidth(0) = 1500
       .ColWidth(1) = 2500
       .ColWidth(2) = 1400
       .ColWidth(3) = 1400
       .ColWidth(4) = 1400
       .ColWidth(5) = 2000
       .ColWidth(6) = 2000
    ElseIf OptRRN.Value = True Then
       .Cols = 6
       .TextMatrix(0, 0) = "Location"
       .TextMatrix(0, 1) = "Client Name"
       .TextMatrix(0, 2) = "RRN No"
        .TextMatrix(0, 3) = "ISIN ID"
       .TextMatrix(0, 4) = "Transaction Qty"
       .TextMatrix(0, 5) = "RRN Delivered Qty"
       .ColWidth(0) = 1500
       .ColWidth(1) = 2500
       .ColWidth(2) = 1500
       .ColWidth(3) = 1500
       .ColWidth(4) = 1500
       .ColWidth(5) = 1500
    ElseIf OptCDCGSLDetail.Value = True Then
       .Cols = 5
       .TextMatrix(0, 0) = "Location"
       .TextMatrix(0, 1) = "Client Name"
       .TextMatrix(0, 2) = "CDC NO"
       .TextMatrix(0, 3) = "Transaction Qty"
       .TextMatrix(0, 4) = "Delivered Qty"
       .ColWidth(0) = 1500
       .ColWidth(1) = 2500
       .ColWidth(2) = 1500
       .ColWidth(3) = 1500
       .ColWidth(4) = 1500
    ElseIf OptGSLvsTxn.Value = True Then
       .Cols = 13
       
       .TextMatrix(0, 0) = "GSL ID"
       .TextMatrix(0, 1) = "SM_prefix"
       .TextMatrix(0, 2) = "CxTF No"
       .TextMatrix(0, 3) = "Txn Type"
       .TextMatrix(0, 4) = "Flag"
       .TextMatrix(0, 5) = "Balance Bags"
       .TextMatrix(0, 6) = "Balance Weight"
       .TextMatrix(0, 7) = "GSL Exchange"
       .TextMatrix(0, 8) = "Txn Exchange"
       .TextMatrix(0, 9) = "GSL Commodity"
       .TextMatrix(0, 10) = "Txn Commodity"
       .TextMatrix(0, 11) = "GSL CMP"
       .TextMatrix(0, 12) = "Txn CMP"
       
       .ColWidth(0) = 1500
       .ColWidth(1) = 1500
       .ColWidth(2) = 1500
       .ColWidth(3) = 800
       .ColWidth(4) = 800
       .ColWidth(5) = 800
       .ColWidth(6) = 800
       .ColWidth(7) = 1000
       .ColWidth(8) = 1000
       .ColWidth(9) = 1200
       .ColWidth(10) = 1200
       .ColWidth(11) = 1200
       .ColWidth(12) = 1200
    
    End If
    .RowHeight(0) = 500
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
End Sub






