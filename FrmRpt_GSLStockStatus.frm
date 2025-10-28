VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GSLStockStatus 
   Caption         =   "GSL Stock Status Report"
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
      Height          =   9645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      Begin VB.CommandButton CmdGO 
         Caption         =   "GO"
         Height          =   350
         Left            =   9240
         TabIndex        =   7
         Top             =   960
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   3525
      End
      Begin VB.ComboBox CmbSGodown 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   2565
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   5400
         TabIndex        =   4
         Top             =   600
         Width           =   3765
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   10200
         TabIndex        =   3
         Top             =   960
         Width           =   1000
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   165
         TabIndex        =   1
         Top             =   600
         Width           =   1605
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8055
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   14208
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox CmbClientName 
         Height          =   315
         ItemData        =   "FrmRpt_GSLStockStatus.frx":0000
         Left            =   5400
         List            =   "FrmRpt_GSLStockStatus.frx":0007
         TabIndex        =   13
         Top             =   960
         Width           =   3765
      End
      Begin VB.Label lblRecCount 
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
         Left            =   11640
         TabIndex        =   15
         Top             =   600
         Width           =   90
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
         Left            =   9240
         TabIndex        =   14
         Top             =   600
         Width           =   2445
      End
      Begin VB.Label Label2 
         Caption         =   "CMP"
         Height          =   300
         Left            =   3240
         TabIndex        =   12
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   5400
         TabIndex        =   11
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Godown"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "Location ID"
         Height          =   300
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Client Name"
         Height          =   300
         Left            =   4440
         TabIndex        =   8
         Top             =   960
         Width           =   1320
      End
   End
End
Attribute VB_Name = "FrmRpt_GSLStockStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCMPName, mCommodityID, mGodownID, mExchangeTypeID, mClientID, mISINID, mClientIDRow As Double
Dim msSM_prefix As Variant


Private Sub CmdGo_Click()
Dim wc As String
Call Grid_Head

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, mCMPName, "Mst_CMP.CMPID", "N", "N")
End If

If CmbSGodown.Text <> "" Then
   wc = GenWc(wc, mGodownID, "Mst_Godown.GodownID", "N", "N")
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Mst_Commodity.CommodityID", "N", "N")
End If

If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, msSM_prefix, "Mst_GSL.sm_prefix", "S", "N")
End If

If CmbClientName.Text <> "" Then
   wc = GenWc(wc, CmbClientName, "Mst_Client.ClientName", "S", "L")
End If



'tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.sm_prefix , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo, Mst_GSL.StackNo,"
'tmp = tmp & " Mst_GSL.LotNo , Mst_GSL.DepositeBags, Mst_GSL.CDTFBags, Mst_GSL.BalanceBags, Mst_GSL.DepositeWeight,"
'tmp = tmp & " Mst_GSL.CDTFWeight , Mst_GSL.BalanceWeight, Mst_ExchangeType.ExchangeType, Mst_Client.ClientName,"
'tmp = tmp & " Mst_GSL.Flag , Mst_GSL.ISINID, Mst_GSL.CDTFNo, Mst_GSL.CDTFDate, Mst_GSL.PledgeNo, Mst_GSL.PledgeBags,"
'tmp = tmp & " Mst_GSL.PledgeWeight , Mst_GSL.SpillageFlag , Mst_GSL.DepositAdjWt,Mst_GSL.WithdrawalAdjWt,mst_gsl.g_uid, "
'tmp = tmp & " Mst_GSL.KRNo, Mst_GSL.GrossWt, "
'tmp = tmp & " Mst_GSL.BGSLID , Mst_GSL.FGSLID, mst_gsl.WHRBags,mst_gsl.WHRWeight"
'tmp = tmp & " FROM (Mst_Client RIGHT JOIN ((Mst_Godown RIGHT JOIN (Mst_Commodity RIGHT JOIN (Mst_CMP RIGHT JOIN Mst_GSL ON"
'tmp = tmp & " Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Commodity.CommodityID = Mst_GSL.CommodityID) ON Mst_Godown.GodownID ="
'tmp = tmp & " Mst_GSL.GodownID) LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID)"
'tmp = tmp & " ON Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow) LEFT JOIN Mst_ISIN ON Mst_GSL.ISINID = Mst_ISIN.ISINID"
'tmp = tmp & " " & wc
'tmp = tmp & " ORDER BY mst_gsl.sm_prefix,Mst_CMP.CMPName,Mst_GSL.GSLID;"


tmp = "SELECT  Mst_Client.ClientName,Mst_ExchangeType.ExchangeType, Mst_GSL.sm_prefix , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo,"
tmp = tmp & " sum(Mst_GSL.DepositeBags)DepositeBags, sum(Mst_GSL.CDTFBags)WithdrawnBags ,sum(Mst_GSL.BalanceBags)BalanceBags,sum(Mst_GSL.DepositeWeight)DepositeQty,sum(Mst_GSL.CDTFWeight)WithdrawnQty,sum(Mst_GSL.BalanceWeight)BalanceQty"
tmp = tmp & " FROM (Mst_Client RIGHT JOIN ((Mst_Godown RIGHT JOIN (Mst_Commodity RIGHT JOIN (Mst_CMP RIGHT JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) ON Mst_Commodity.CommodityID = Mst_GSL.CommodityID)"
tmp = tmp & " ON Mst_Godown.GodownID = Mst_GSL.GodownID) LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID) ON Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow)"
tmp = tmp & " " & wc
tmp = tmp & " group by Mst_Client.ClientName,Mst_ExchangeType.ExchangeType, Mst_GSL.sm_prefix , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo"
tmp = tmp & " ORDER BY mst_gsl.sm_prefix,Mst_CMP.CMPName"



Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    MSHFlexGrid1.TextMatrix(I, 6) = Format(MSHFlexGrid1.TextMatrix(I, 6), "########0")
    MSHFlexGrid1.TextMatrix(I, 7) = Format(MSHFlexGrid1.TextMatrix(I, 7), "########0")
    MSHFlexGrid1.TextMatrix(I, 8) = Format(MSHFlexGrid1.TextMatrix(I, 8), "########0")
    MSHFlexGrid1.TextMatrix(I, 9) = Format(MSHFlexGrid1.TextMatrix(I, 9), "########0")
    MSHFlexGrid1.TextMatrix(I, 10) = Format(MSHFlexGrid1.TextMatrix(I, 10), "########0.000")
    MSHFlexGrid1.TextMatrix(I, 11) = Format(MSHFlexGrid1.TextMatrix(I, 11), "########0.000")
    
    TmpRs.MoveNext
Next
LblRecCount.Refresh
LblRecCount.Caption = TmpRs.RecordCount

'RsMst_GSL.Close
'Call TableMst_GSL(" Order By sm_prefix,CMPID,GodownID,StackNo,LotNo ")
End Sub


Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12
MSHFlexGrid1.TextMatrix(0, 0) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 1) = "Exchange Type  "
MSHFlexGrid1.TextMatrix(0, 2) = "Location Name "
MSHFlexGrid1.TextMatrix(0, 3) = "CMP "
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity "
MSHFlexGrid1.TextMatrix(0, 5) = "Godown "
MSHFlexGrid1.TextMatrix(0, 6) = "Deposited Bags "
MSHFlexGrid1.TextMatrix(0, 7) = "Withdrawn Bags "
MSHFlexGrid1.TextMatrix(0, 8) = "Balance Bags"
MSHFlexGrid1.TextMatrix(0, 9) = "Deposited Qty(in MT)"
MSHFlexGrid1.TextMatrix(0, 10) = "Withdrawn Qty(in MT)"
MSHFlexGrid1.TextMatrix(0, 11) = "Balance Qty(in MT)"

MSHFlexGrid1.ColWidth(1) = 1200




MSHFlexGrid1.RowHeight(0) = 500

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GSLStockStatus.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLStockStatus.xls")
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
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
Call TableMst_StoreManLocations("")

CmbPreFix.Clear
If RsMst_StoreManLocations.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_StoreManLocations.RecordCount
    CmbPreFix.AddItem RsMst_StoreManLocations!StoreManLocations
    RsMst_StoreManLocations.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "StoreManLocations = '" & CmbPreFix.Text & "'"
If RsMst_StoreManLocations.EOF Then
   MsgBox "Invalid PreFix Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
msSM_prefix = RsMst_StoreManLocations!SM_Prefix

End Sub

Private Sub CmbSCmp_GotFocus()
tmp = CmbSCmp.Text
Call TableMst_CMP
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp
End Sub
Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
End Sub
Private Sub CmbSGodown_GotFocus()
If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPName)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbStack_GotFocus()

If CmbSGodown.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where CMPID = " & mCMPName & " And GodownID = " & mGodownID, "D")
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next

CmbStack.Text = tmp
End Sub
'----
Private Sub CmbStack_LostFocus()
If CmbStack.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "StackNo = '" & CmbStack.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbStack.SetFocus
   Exit Sub
End If

mCmbStack = CmbStack.Text
End Sub
Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbSCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbSCommodity.Text = tmp
End Sub

Private Sub CmbSCommodity_LostFocus()
If CmbSCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID

End Sub
Private Sub CmbSExchangeType_GotFocus()
tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeType.Text = tmp
End Sub

Private Sub CmbSExchangeType_LostFocus()
If CmbSExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub
Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   'Call GButtonLockAD(Me)
'   Frame0.Enabled = False
'   MsgBox "Access denied !!!!!!!!!"
'   Exit Sub
'End If
Frame0.Enabled = True
CmbClientName.SelText = "RELIANCE"
CmbClientName.Enabled = False

End Sub


