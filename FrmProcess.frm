VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmProcess 
   Caption         =   "Auto Process of Opening Balance"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   61276161
      CurrentDate     =   39993
   End
   Begin VB.TextBox TxtPreFix 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox CmbLocationID 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   6120
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   12726
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "Go"
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "Location ID"
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
      TabIndex        =   5
      Top             =   660
      Width           =   1695
   End
   Begin VB.Label LblLocationID 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   270
      Width           =   1440
   End
End
Attribute VB_Name = "FrmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID As Variant

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where  LocationID <> 0 ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!! "
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
   mLocationID = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
End Sub


Private Sub CmdGo_Click()
If TxtPreFix = "" Then Exit Sub
'tmp = "Delete from Mst_GSL where sm_prefix = '" & TxtPreFix & "'"
'Call TmpTable

Call TableMst_GSL(" where sm_prefix = '" & TxtPreFix & "'")
If RsMst_GSL.RecordCount > 0 Then
   MsgBox "Import Process already executed for selected location !!!!"
   Exit Sub
End If

tmp = "SELECT  Txn_GSL_OpeningBalance.GSLID, Mst_ExchangeType.ExchangeType, Mst_Commodity.Commodity, "
tmp = tmp & " Txn_GSL_OpeningBalance.ClientIDRow, Mst_CMP.CMPName, Mst_Godown.GodownNo, Txn_GSL_OpeningBalance.StackNo,"
tmp = tmp & " Txn_GSL_OpeningBalance.LotNo, Txn_GSL_OpeningBalance.DepositBags, Txn_GSL_OpeningBalance.DepositWeight,"
tmp = tmp & " Txn_GSL_OpeningBalance.SM_Prefix, Txn_GSL_OpeningBalance.CMPID, Txn_GSL_OpeningBalance.GodownID,"
tmp = tmp & " Txn_GSL_OpeningBalance.ExchangeTypeID , Txn_GSL_OpeningBalance.CommodityID, Txn_GSL_OpeningBalance.ISINID"
tmp = tmp & " FROM  Txn_GSL_OpeningBalance INNER JOIN"
tmp = tmp & " Mst_CMP ON Txn_GSL_OpeningBalance.CMPID = Mst_CMP.CMPID INNER JOIN"
tmp = tmp & " Mst_Commodity ON Txn_GSL_OpeningBalance.CommodityID = Mst_Commodity.CommodityID INNER JOIN"
tmp = tmp & " Mst_ExchangeType ON Txn_GSL_OpeningBalance.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID INNER JOIN"
tmp = tmp & " Mst_Godown ON Txn_GSL_OpeningBalance.GodownID = Mst_Godown.GodownID"
tmp = tmp & " where  Txn_GSL_OpeningBalance.sm_prefix = '" & TxtPreFix & "'"
tmp = tmp & " ORDER BY Mst_ExchangeType.ExchangeType, Mst_CMP.CMPName, Mst_Commodity.Commodity,"
tmp = tmp & " Mst_Godown.GodownNo, Txn_GSL_OpeningBalance.ClientIDRow , Txn_GSL_OpeningBalance.StackNo, Txn_GSL_OpeningBalance.LotNo"

Call TmpTable
If TmpRs.RecordCount = 0 Then
   MsgBox "No Data found !!!"
   Exit Sub
End If
Call TableMst_CommodityDetail
Call TableTxn_CxTF_ClientDetail
Call TableTxn_CxTF_Details
Call TableTxn_CxTF_GSL
Call TableTxn_CxTF_TruckDetail

MSHFlexGrid1.Cols = TmpRs.Fields.Count + 1
MSHFlexGrid1.Rows = TmpRs.RecordCount + 1

MSHFlexGrid1.TextMatrix(0, 0) = "GLS ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 2) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 3) = "ClientID Row"
MSHFlexGrid1.TextMatrix(0, 4) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 6) = "stack No"
MSHFlexGrid1.TextMatrix(0, 7) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 8) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 9) = "Weight"
MSHFlexGrid1.TextMatrix(0, 10) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 11) = "CMP ID"
MSHFlexGrid1.TextMatrix(0, 12) = "Godown ID"
MSHFlexGrid1.TextMatrix(0, 13) = "Exchange ID"
MSHFlexGrid1.TextMatrix(0, 14) = "Commodity ID"
MSHFlexGrid1.TextMatrix(0, 15) = "ISIN"
Dim CdtfNo As Variant
CNo = 0
For I = 1 To TmpRs.RecordCount
    If I = 1 Then
       CNo = CNo + 1
       Call AddCdftDetail(CNo)
    End If
    For c = o To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, c) = TmpRs.Fields(c)
    Next
    AddGSLMaster (CNo)
    Call AddGSLDetail(CNo)
    TmpRs.MoveNext
    If TmpRs.EOF Then Exit For
    If RsTxn_CxTF_Details!CMPID <> TmpRs!CMPID Or RsTxn_CxTF_Details!CommodityID <> TmpRs!CommodityID Or _
       RsTxn_CxTF_Details!ExchangeTypeID <> TmpRs!ExchangeTypeID Then
       CNo = CNo + 1
       Call AddCdftDetail(CNo)
    End If
Next


tmp = "SELECT SM_Prefix, CxTFNo, ClientIDRow, TxnType From Txn_CxTF_GSL where SM_Prefix = '" & TxtPreFix & "' GROUP BY SM_Prefix, CxTFNo, " & _
"ClientIDRow, TxnType  ORDER BY CxTFNo"
Call TmpTable

For I = 1 To TmpRs.RecordCount
    RsTxn_CxTF_ClientDetail.AddNew
    RsTxn_CxTF_ClientDetail!ClientDetailID = getCxtf_ClinetMax
    RsTxn_CxTF_ClientDetail!SM_Prefix = TmpRs!SM_Prefix
    RsTxn_CxTF_ClientDetail!CxTFNo = TmpRs!CxTFNo
    RsTxn_CxTF_ClientDetail!ClientIDRow = TmpRs!ClientIDRow
    RsTxn_CxTF_ClientDetail!Flag = ""
    RsTxn_CxTF_ClientDetail!TxnType = "D"
    RsTxn_CxTF_ClientDetail!G_UID = GetGUID
    RsTxn_CxTF_ClientDetail.Update
    TmpRs.MoveNext
Next


tmp = "SELECT  SM_Prefix, CxTFNo, TxnType, SUM(GrossWeight) AS TotQty " & _
"From Txn_CxTF_GSL where SM_Prefix = '" & TxtPreFix & "' GROUP BY SM_Prefix, CxTFNo, TxnType ORDER BY CxTFNo "
Call TmpTable
For I = 1 To TmpRs.RecordCount
    RsTxn_CxTF_TruckDetail.AddNew
    RsTxn_CxTF_TruckDetail!TruckDetailID = getCxtf_TruckMax
    RsTxn_CxTF_TruckDetail!SM_Prefix = TmpRs!SM_Prefix
    RsTxn_CxTF_TruckDetail!CxTFNo = TmpRs!CxTFNo
    RsTxn_CxTF_TruckDetail!TruckNo = "Opening"
    RsTxn_CxTF_TruckDetail!TotalWeight = TmpRs!TotQty
    RsTxn_CxTF_TruckDetail!TareWeight = 0
    RsTxn_CxTF_TruckDetail!GrossWeight = TmpRs!TotQty
    RsTxn_CxTF_TruckDetail!WeighBridgeSlipno = "Opening Balance"
    RsTxn_CxTF_TruckDetail!Remarks = "Opening Balance"
    RsTxn_CxTF_TruckDetail!Flag = ""
    RsTxn_CxTF_TruckDetail!TxnType = "D"
    RsTxn_CxTF_TruckDetail!G_UID = GetGUID
    RsTxn_CxTF_TruckDetail.Update
    TmpRs.MoveNext
Next

Call TableMst_CMP(" where CloseDate is null And LocationID =" & mLocationID)
For I = 1 To RsMst_CMP.RecordCount
    'MsgBox "Current Date Updated for CMP-->" & RsMst_CMP!CMPName
    RsMst_CMP!CurrentDate = DTPicker1.Value
    RsMst_CMP.Update
    RsMst_CMP.MoveNext
    
Next


MsgBox "Done !!!! "
End Sub
Function getCxtf_TruckMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(TruckDetailID) from Txn_CxTF_TruckDetail where SM_Prefix = '" & TmpRs!SM_Prefix & "'"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
End If
Retval = Retval + 1
TmpRs1.Close
getCxtf_TruckMax = Retval

End Function


Function getCxtf_ClinetMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(ClientDetailID) from Txn_CxTF_ClientDetail "
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
End If
Retval = Retval + 1
TmpRs1.Close
getCxtf_ClinetMax = Retval

End Function


Sub AddGSLMaster(CxTfNo_ As Variant)
RsMst_GSL.AddNew
RsMst_GSL!GSLID = TmpRs!GSLID
RsMst_GSL!SM_Prefix = TmpRs!SM_Prefix
RsMst_GSL!CMPID = TmpRs!CMPID
RsMst_GSL!CommodityID = TmpRs!CommodityID
RsMst_GSL!GodownID = TmpRs!GodownID
RsMst_GSL!StackNo = TmpRs!StackNo
RsMst_GSL!LotNo = TmpRs!LotNo
RsMst_GSL!DepositeWeight = TmpRs!DepositWeight
RsMst_GSL!DepositeBags = TmpRs!DepositBags
RsMst_GSL!WHRBags = 0
RsMst_GSL!CDTFBags = 0
RsMst_GSL!BalanceBags = TmpRs!DepositBags
RsMst_GSL!WHRWeight = 0
RsMst_GSL!CDTFWeight = 0
RsMst_GSL!BalanceWeight = TmpRs!DepositWeight
RsMst_GSL!ExchangeTypeID = TmpRs!ExchangeTypeID
RsMst_GSL!ClientIDRow = TmpRs!ClientIDRow
RsMst_GSL!PledgeNo = ""
RsMst_GSL!Flag = "O"
RsMst_GSL!ISINID = TmpRs!ISINID
RsMst_GSL!CdtfNo = CxTfNo_
RsMst_GSL!CDTFDate = DTPicker1.Value
RsMst_GSL!FirstISINID = TmpRs!ISINID
RsMst_GSL!PledgeBags = 0
RsMst_GSL!PledgeWeight = 0
RsMst_GSL!SpillageFlag = 0
RsMst_GSL!DepositAdjWt = 0
RsMst_GSL!WithdrawalAdjWt = 0
RsMst_GSL!G_UID = GetGUID
RsMst_GSL!Flag1 = "O"
RsMst_GSL!LCDTF_Date = DTPicker1.Value
RsMst_GSL!GrossWt = TmpRs!DepositWeight
RsMst_GSL!KRNo = CxTfNo_
RsMst_GSL.Update
End Sub
Sub AddGSLDetail(CxTfNo_ As Variant)
Dim TxtGCTDGunnyWt, TxtAdjGrossWt, TxtGCTDStdDeduction, TxtGCTDDematWt As Variant

RsMst_CommodityDetail.Filter = "CommodityID = " & TmpRs!CommodityID & " And ExchangeTypeID = " & TmpRs!ExchangeTypeID
If RsMst_CommodityDetail.RecordCount <> 0 Then
    TxtGCTDGunnyWt = Format((TmpRs!DepositBags * RsMst_CommodityDetail!GunnyWt) / 1000, "#######0.000")
    TxtAdjGrossWt = TmpRs!DepositWeight
    tmp = Val(TxtAdjGrossWt) - Val(TxtGCTDGunnyWt)
    TxtGCTDStdDeduction = Format(((tmp) * (RsMst_CommodityDetail!StdDeduction)) / 100, "#######0.000")
    TxtGCTDDematWt = Format(Val(TxtAdjGrossWt) - (Val(TxtGCTDStdDeduction) + Val(TxtGCTDGunnyWt)), "#######0.000")
End If

RsTxn_CxTF_GSL.AddNew
RsTxn_CxTF_GSL!GSLDetailID = getCxtf_GSLMax
RsTxn_CxTF_GSL!SM_Prefix = TmpRs!SM_Prefix
RsTxn_CxTF_GSL!CxTFNo = CxTfNo_
RsTxn_CxTF_GSL!NoofBags = TmpRs!DepositBags
RsTxn_CxTF_GSL!GrossWeight = TmpRs!DepositWeight
RsTxn_CxTF_GSL!DematWeight = TxtGCTDDematWt
RsTxn_CxTF_GSL!ClientIDRow = TmpRs!ClientIDRow
RsTxn_CxTF_GSL!Remarks = "Opening Balance "
RsTxn_CxTF_GSL!Flag = ""
RsTxn_CxTF_GSL!GunnyWeight = TxtGCTDGunnyWt
RsTxn_CxTF_GSL!StdDeduction = TxtGCTDStdDeduction
RsTxn_CxTF_GSL!AdjNCMSLWeight = TxtAdjGrossWt
RsTxn_CxTF_GSL!AdjNClientWeight = 0
RsTxn_CxTF_GSL!WeightGainLoss = 0
RsTxn_CxTF_GSL!GSLID = TmpRs!GSLID
RsTxn_CxTF_GSL!TxnType = "D"
RsTxn_CxTF_GSL!G_UID = GetGUID
RsTxn_CxTF_GSL.Update


RsMst_GSL!DepositeWeight = TxtGCTDDematWt
RsMst_GSL!BalanceWeight = TxtGCTDDematWt
RsMst_GSL.Update


End Sub

Function getCxtf_GSLMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(GSLDetailID) from Txn_CxTF_GSL where SM_Prefix = '" & TmpRs!SM_Prefix & "'"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
End If
Retval = Retval + 1
TmpRs1.Close
getCxtf_GSLMax = Retval

End Function

Sub AddCdftDetail(CxTfNo_ As Variant)

RsTxn_CxTF_Details.AddNew
RsTxn_CxTF_Details!CxTFNo = CxTfNo_
RsTxn_CxTF_Details!TxnType = "D"
RsTxn_CxTF_Details!SM_Prefix = TmpRs!SM_Prefix
RsTxn_CxTF_Details!CxtfDate = DTPicker1.Value
RsTxn_CxTF_Details!CMPID = TmpRs!CMPID
RsTxn_CxTF_Details!CommodityID = TmpRs!CommodityID
RsTxn_CxTF_Details!ExchangeTypeID = TmpRs!ExchangeTypeID
RsTxn_CxTF_Details!BookedSpaceFlag = 0
RsTxn_CxTF_Details!Flag = ""
RsTxn_CxTF_Details!Agent = "Opening Balance"
RsTxn_CxTF_Details!Remark = "Opening Balance"
RsTxn_CxTF_Details!G_UID = GetGUID
RsTxn_CxTF_Details!HologramNo = TmpRs!CMPID & " - " & CxTFNo
RsTxn_CxTF_Details.Update

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   CmdGo.Enabled = False
   Exit Sub
End If
CmdGo.Enabled = True
DTPicker1.Value = Date
End Sub
