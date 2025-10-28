VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_CxTF 
   Caption         =   "Detail Listing"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   10350
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbTxnType 
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
         ItemData        =   "FrmRpt_CxTF.frx":0000
         Left            =   960
         List            =   "FrmRpt_CxTF.frx":000A
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "GSL Detail"
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
         Left            =   9240
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Truck Detail"
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
         Left            =   7440
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Client tab detail"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CITF"
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
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "WHR Details"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   1575
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
         Height          =   495
         Left            =   13695
         TabIndex        =   5
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
         Height          =   495
         Left            =   11265
         TabIndex        =   4
         Top             =   165
         Width           =   1215
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
         Height          =   495
         Left            =   12480
         TabIndex        =   3
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "TxnType"
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
         TabIndex        =   12
         Top             =   293
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13996
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_CxTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbTxnType_LostFocus()
If CmbTxnType.Text = "" Then Exit Sub

'If LCase(CmbTxnType.Text) <> "deposit" Or LCase(CmbTxnType.Text) <> "withdrawal" Then
'   MsgBox "Invalid Selection!!!"
'   CmbTxnType.SetFocus
'   Exit Sub
'End If

End Sub

Private Sub CmdGetReport_Click()
Dim Wc As Variant

Call Grid_Head

If LCase(CmbTxnType.Text) = "deposit" Then
   Wc = "D"
ElseIf LCase(CmbTxnType.Text) = "withdrawal" Then
   Wc = "W"
Else
   Wc = "W"
End If

If Option1.Value = True Then
   tmp = "SELECT Txn_WHR_Detail.WHRNo, Txn_WHR_Detail.SM_Prefix, Txn_WHR_Detail.WHRDate, Txn_WHR_Detail.ConditionID, " & _
         "Txn_WHR_Detail.StorageFrom, Txn_WHR_Detail.StorageTo, Txn_WHR_Detail.Grade, Txn_WHR_Detail.Weigher, " & _
         "Txn_WHR_Detail.Assayer, Txn_WHR_Detail.GradeByAssayer, Txn_WHR_Detail.GradeValidTill, Txn_WHR_Detail.StorageCharges, " & _
         "Txn_WHR_Detail.StorageAmount, Txn_WHR_Detail.ClientIDRow, Mst_Client.ClientName, Mst_Commodity.Commodity, " & _
         "Txn_WHR_Detail.WHRNoOfBags, Txn_WHR_Detail.WHRQuantity, Mst_Godown.GodownNo, Txn_WHR_Detail.DeliveredQuantity, " & _
         "Txn_WHR_Detail.DeliveredNoOfBags, Txn_WHR_Detail.Rate, Txn_WHR_Detail.HologramNo, Txn_WHR_Detail.PolicyNo, " & _
         "Txn_WHR_Detail.CoverNoteNo , Txn_WHR_Detail.SaidToContain, Txn_WHR_Detail.EstimatedWeight " & _
         "FROM Mst_Godown INNER JOIN Mst_CMP ON Mst_Godown.CMPID = Mst_CMP.CMPID INNER JOIN " & _
         "Txn_WHR_Detail INNER JOIN Mst_Commodity ON Txn_WHR_Detail.CommodityID = Mst_Commodity.CommodityID INNER JOIN  " & _
         "Mst_Client ON Txn_WHR_Detail.ClientIDRow = Mst_Client.ClientIDRow ON Mst_Godown.GodownID = Txn_WHR_Detail.GodownID "

ElseIf Option2.Value = True Then

       tmp = "SELECT  Txn_CxTF_Details.CxTFNo, Txn_CxTF_Details.TxnType, Txn_CxTF_Details.SM_Prefix, " & _
             "Txn_CxTF_Details.CxTFDate, Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_ExchangeType.ExchangeType, " & _
             "Txn_CxTF_Details.Flag , Txn_CxTF_Details.Agent, Txn_CxTF_Details.Remark, Txn_CxTF_Details.HologramNo " & _
             "FROM Txn_CxTF_Details INNER JOIN " & _
             "Mst_CMP ON Txn_CxTF_Details.CMPID = Mst_CMP.CMPID INNER JOIN " & _
             "Mst_Commodity ON Txn_CxTF_Details.CommodityID = Mst_Commodity.CommodityID INNER JOIN " & _
             "Mst_ExchangeType ON Txn_CxTF_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID " & _
             "WHERE (Txn_CxTF_Details.TxnType = '" & Wc & "') " & _
             "ORDER BY Txn_CxTF_Details.SM_Prefix, Txn_CxTF_Details.CxTFNo"
ElseIf Option3.Value = True Then
       
       tmp = "SELECT  Txn_CxTF_ClientDetail.SM_Prefix, Txn_CxTF_ClientDetail.CxTFNo AS [CITF No], Mst_Client.ClientName, " & _
             "Txn_CxTF_ClientDetail.TxnType, Txn_CxTF_ClientDetail.RRNID, Txn_CxTF_ClientDetail.AdjRRNQty, " & _
             "Txn_CxTF_ClientDetail.AdjRRNNoofBags , Mst_ExchangeType.ExchangeType, Mst_RRN.RRNNo " & _
             "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
             "INNER JOIN Mst_ExchangeType ON Mst_Depository.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID INNER JOIN " & _
             "Txn_CxTF_ClientDetail ON Mst_Client.ClientIDRow = Txn_CxTF_ClientDetail.ClientIDRow LEFT OUTER JOIN " & _
             "Mst_RRN ON Mst_Depository.DepositoryID = Mst_RRN.DepositoryID AND " & _
             "Txn_CxTF_ClientDetail.RRNID = Mst_RRN.RRNID " & _
             "WHERE (Txn_CxTF_ClientDetail.TxnType = '" & Wc & "') " & _
             "ORDER BY Txn_CxTF_ClientDetail.SM_Prefix, Txn_CxTF_ClientDetail.CxTFNo"
ElseIf Option4.Value = True Then
       tmp = "SELECT  SM_Prefix, CxTFNo, TxnType, TruckNo, WeighBridgeSlipno, TotalWeight, TareWeight, " & _
             "GrossWeight, Remarks, Flag " & _
             "From Txn_CxTF_TruckDetail  " & _
             "WHERE (TxnType = '" & Wc & "')  " & _
             "ORDER BY SM_Prefix, CxTFNo "
ElseIf Option5.Value = True Then
       tmp = "SELECT  Txn_CxTF_GSL.SM_Prefix, Txn_CxTF_GSL.CxTFNo, Mst_Godown.GodownNo, Mst_GSL.StackNo, " & _
             "Mst_GSL.LotNo, Txn_CxTF_GSL.NoOfBags, Txn_CxTF_GSL.GrossWeight, Txn_CxTF_GSL.DematWeight, " & _
             "Txn_CxTF_GSL.Remarks, Txn_CxTF_GSL.Flag, Txn_CxTF_GSL.GunnyWeight, Txn_CxTF_GSL.StdDeduction, " & _
             "Txn_CxTF_GSL.AdjNCMSLWeight, Txn_CxTF_GSL.AdjNClientWeight, Txn_CxTF_GSL.WeightGainLoss, " & _
             "Txn_CxTF_GSL.TxnType  " & _
             "FROM Txn_CxTF_GSL INNER JOIN " & _
             "Mst_GSL ON Txn_CxTF_GSL.GSLID = Mst_GSL.GSLID AND Txn_CxTF_GSL.SM_Prefix = Mst_GSL.SM_Prefix INNER JOIN " & _
             "Mst_Godown ON Mst_GSL.GodownID = Mst_Godown.GodownID " & _
             "WHERE (Txn_CxTF_GSL.TxnType = '" & Wc & "')  " & _
             "ORDER BY Txn_CxTF_GSL.SM_Prefix, Txn_CxTF_GSL.CxTFNo"
End If
Call TmpTable
Call dispdata

End Sub
Public Sub dispdata()
    Dim x As Double
    x = 0
    If TmpRs.RecordCount > 0 Then
       For I = 1 To TmpRs.RecordCount
           x = MSHFlexGrid1.Rows - 1
           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           For c = 0 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
           Next
           TmpRs.MoveNext
        Next
    Else
       MsgBox "No Record Found!!!"
    End If
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
Call Grid_Head
'CmdExcel.Enabled = False
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
If Option1.Value = True Then
   Call Grid_Head1
ElseIf Option2.Value = True Then
   Call Grid_Head2
ElseIf Option3.Value = True Then
   Call Grid_Head3
ElseIf Option4.Value = True Then
   Call Grid_Head4
ElseIf Option5.Value = True Then
   Call Grid_Head5
End If

End Sub
Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 27
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "WHRNo"
MSHFlexGrid1.TextMatrix(0, 1) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 2) = "WHRDate"
MSHFlexGrid1.TextMatrix(0, 3) = "ConditionID"
MSHFlexGrid1.TextMatrix(0, 4) = "StorageFrom"
MSHFlexGrid1.TextMatrix(0, 5) = "StorageTo"
MSHFlexGrid1.TextMatrix(0, 6) = "Grade"
MSHFlexGrid1.TextMatrix(0, 7) = "Weigher"
MSHFlexGrid1.TextMatrix(0, 8) = "Assayer"
MSHFlexGrid1.TextMatrix(0, 9) = "GradeByAssayer"
MSHFlexGrid1.TextMatrix(0, 10) = "GradeValidTill"
MSHFlexGrid1.TextMatrix(0, 11) = "StorageCharges"
MSHFlexGrid1.TextMatrix(0, 12) = "StorageAmount"
MSHFlexGrid1.TextMatrix(0, 13) = "ClientIDRow"
MSHFlexGrid1.TextMatrix(0, 14) = "ClientName"
MSHFlexGrid1.TextMatrix(0, 15) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 16) = "WHRNoOfBags"
MSHFlexGrid1.TextMatrix(0, 17) = "WHRQuantity"
MSHFlexGrid1.TextMatrix(0, 18) = "GodownNo"
MSHFlexGrid1.TextMatrix(0, 19) = "DeliveredQuantity"
MSHFlexGrid1.TextMatrix(0, 20) = "DeliveredNoOfBags"
MSHFlexGrid1.TextMatrix(0, 21) = "Rate"
MSHFlexGrid1.TextMatrix(0, 22) = "HologramNo"
MSHFlexGrid1.TextMatrix(0, 23) = "PolicyNo"
MSHFlexGrid1.TextMatrix(0, 24) = "CoverNoteNo"
MSHFlexGrid1.TextMatrix(0, 25) = "SaidToContain"
MSHFlexGrid1.TextMatrix(0, 26) = "EstimatedWeight"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 3000
MSHFlexGrid1.ColWidth(13) = 1500
MSHFlexGrid1.ColWidth(14) = 1000
MSHFlexGrid1.ColWidth(15) = 1000

End Sub

Private Sub Grid_Head2()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "CITF No"
MSHFlexGrid1.TextMatrix(0, 1) = "TxnType"
MSHFlexGrid1.TextMatrix(0, 2) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 3) = "CxTFDate"
MSHFlexGrid1.TextMatrix(0, 4) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 7) = "Flag"
MSHFlexGrid1.TextMatrix(0, 8) = "Agent"
MSHFlexGrid1.TextMatrix(0, 9) = "Remark"
MSHFlexGrid1.TextMatrix(0, 10) = "HologramNo"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000

End Sub

Private Sub Grid_Head3()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 1) = "CITF No"
MSHFlexGrid1.TextMatrix(0, 2) = "ClientName"
MSHFlexGrid1.TextMatrix(0, 3) = "TxnType"
MSHFlexGrid1.TextMatrix(0, 4) = "WHR"
MSHFlexGrid1.TextMatrix(0, 5) = "AdjRRNQty"
MSHFlexGrid1.TextMatrix(0, 6) = "AdjRRNNoofBags"
MSHFlexGrid1.TextMatrix(0, 7) = "ExchangeType"
MSHFlexGrid1.TextMatrix(0, 8) = "RRNNo"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 2000


End Sub

Private Sub Grid_Head4()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 10
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 1) = "CxTFNo"
MSHFlexGrid1.TextMatrix(0, 2) = "TxnType"
MSHFlexGrid1.TextMatrix(0, 3) = "TruckNo"
MSHFlexGrid1.TextMatrix(0, 4) = "WeighBridgeSlipno"
MSHFlexGrid1.TextMatrix(0, 5) = "TotalWeight"
MSHFlexGrid1.TextMatrix(0, 6) = "TareWeight"
MSHFlexGrid1.TextMatrix(0, 7) = "GrossWeight"
MSHFlexGrid1.TextMatrix(0, 8) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 9) = "Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 1000

End Sub

Private Sub Grid_Head5()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 16
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 1) = "CxTFNo"
MSHFlexGrid1.TextMatrix(0, 2) = "GodownNo"
MSHFlexGrid1.TextMatrix(0, 3) = "StackNo"
MSHFlexGrid1.TextMatrix(0, 4) = "LotNo"
MSHFlexGrid1.TextMatrix(0, 5) = "NoOfBags"
MSHFlexGrid1.TextMatrix(0, 6) = "GrossWeight"
MSHFlexGrid1.TextMatrix(0, 7) = "DematWeight"
MSHFlexGrid1.TextMatrix(0, 8) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 9) = "Flag"
MSHFlexGrid1.TextMatrix(0, 10) = "GunnyWeight"
MSHFlexGrid1.TextMatrix(0, 11) = "StdDeduction"
MSHFlexGrid1.TextMatrix(0, 12) = "AdjNCMSLWeight"
MSHFlexGrid1.TextMatrix(0, 13) = "AdjNClientWeight"
MSHFlexGrid1.TextMatrix(0, 14) = "WeightGainLoss"
MSHFlexGrid1.TextMatrix(0, 15) = "TxnType"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 3000
MSHFlexGrid1.ColWidth(13) = 1500
MSHFlexGrid1.ColWidth(14) = 1000
MSHFlexGrid1.ColWidth(15) = 1000

End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CxTF.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CxTF.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
