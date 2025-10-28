VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_PendingGSLforDC 
   Caption         =   "Pending GSL for Rate Card"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   13845
         TabIndex        =   3
         Top             =   225
         Width           =   1095
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
         Left            =   11535
         TabIndex        =   2
         Top             =   225
         Width           =   1215
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
         Left            =   12750
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record :"
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
         Left            =   8400
         TabIndex        =   5
         Top             =   285
         Width           =   1335
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
         Left            =   9720
         TabIndex        =   4
         Top             =   285
         Width           =   60
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   9240
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8415
      Left            =   120
      TabIndex        =   7
      Top             =   800
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14843
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_PendingGSLforDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mR As Double, mC As Double

Private Sub CmdClear_Click()
Call Grid_Head
LblTotalRecord.Caption = ""
CmdExcel.Enabled = True
End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
MsgBox "Wait till next message."
CmdExcel.Enabled = False
tmp = " Select Distinct StateName,LocationName,SM_Prefix,CMPName,Mst_GSL_ID,GodownNo,Commodity,Client,BalanceBags,BalanceWeight,BagSize,ContractID,NContractID,G_UID,CDTFDate,CxTF_DetailsID,Madeupflag,Flag,HologramNo,TransType From ("
tmp = tmp & " Select Distinct s.StateName,l.LocationName,l.SM_Prefix,Mc.CMPName,GSL.Mst_GSL_ID,MG.GodownNo,mcom.Commodity,Mclient.ClientName+'#'+Mclient.ClientIDRow+'#'+Convert(Varchar,"
tmp = tmp & " Mclient.ClientID) 'Client'  ,GSL.BalanceBags,GSL.BalanceWeight, (GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,GSL.G_UID,"
tmp = tmp & " GSL.CDTFDate , GSL.CxTF_DetailsID, GSL.Madeupflag,TCD.Flag,TCD.HologramNo, 'GSL' TransType"
tmp = tmp & " From  Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Left Join Txn_CxTF_GSL TCG on GSL.SM_Prefix = TCG.SM_Prefix and TCG.TxnType = 'D' And TCG.CxTF_DetailsID = GSL.CxTF_DetailsID"
tmp = tmp & " Left Join Txn_CxTF_Details TCD on TCG.SM_Prefix = TCD.SM_Prefix and TCG.TxnType = TCD.TxnType And TCG.CxTF_DetailsID = TCD.CxTF_DetailsID"
tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0' and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And  GSL.CDTFDate <= '" & Date & "' And GSL.Status not in ('Z') And Convert(Varchar,MG.YearlyBillingFlg)+'-'+COnvert(Varchar,Mcom.YearlyBillingFlg) not in ('1-1')"
'tmp = tmp & " Order By s.StateName,l.LocationName,l.SM_Prefix"
tmp = tmp & " Union All"
tmp = tmp & " Select Distinct s.StateName,l.LocationName,l.SM_Prefix,Mc.CMPName,GSL.Mst_GSL_ID,MG.GodownNo,mcom.Commodity,Mclient.ClientName+'#'+Mclient.ClientIDRow+'#'+Convert(Varchar,Mclient.ClientID) 'Client'  ,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,GSL.G_UID,GSL.CDTFDate,GSL.CxTF_DetailsID, GSL.Madeupflag,b.Flag,b.HologramNo, 'Withdrawal' TransType"
tmp = tmp & " From  Txn_CxTF_GSL  a"
tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
tmp = tmp & " Inner Join Mst_GSL GSL On a.Mst_GSL_ID = GSL.Mst_GSL_ID and a.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on b.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on b.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Where a.TxnType = 'W' And b.CxTFDate > '2017-12-01' and GSL.DepositeBags > 0 And GSL.BalanceBags >= 0 And  GSL.CDTFDate > '2017-01-01'   And b.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And GSL.Status in ('O','E') And Convert(Varchar,MG.YearlyBillingFlg)+'-'+COnvert(Varchar,Mcom.YearlyBillingFlg) not in ('1-1')"
tmp = tmp & " Union All"
tmp = tmp & " Select Distinct s.StateName,l.LocationName,l.SM_Prefix,Mc.CMPName,GSL.Mst_GSL_ID,MG.GodownNo,mcom.Commodity,Mclient.ClientName+'#'+Mclient.ClientIDRow+'#'+Convert(Varchar,Mclient.ClientID) 'Client'  ,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,GSL.G_UID,GSL.CDTFDate,GSL.CxTF_DetailsID, GSL.Madeupflag,b.Flag,b.HologramNo, 'Deposit' TransType"
tmp = tmp & " From  Txn_CxTF_GSL  a"
tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
tmp = tmp & " Inner Join Mst_GSL GSL On a.Mst_GSL_ID = GSL.Mst_GSL_ID and a.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on b.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on b.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Where a.TxnType = 'D' And b.CxTFDate > '2017-12-01' and GSL.DepositeBags > 0 And GSL.BalanceBags >= 0 And  GSL.CDTFDate > '2017-01-01'   And b.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And GSL.Status in ('O','E') And Convert(Varchar,MG.YearlyBillingFlg)+'-'+COnvert(Varchar,Mcom.YearlyBillingFlg) not in ('1-1')"
tmp = tmp & " ) A"
tmp = tmp & " Order By StateName,LocationName,SM_Prefix,CMPName,Mst_GSL_ID"

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount
If TmpRs.RecordCount > 0 Then

For mR = 1 To TmpRs.RecordCount
    With MSHFlexGrid1
        .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
        .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!Mst_GSL_ID), "", TmpRs!Mst_GSL_ID)
        .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
        .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!Client), "", TmpRs!Client)
        .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
        .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
        .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!BagSize), "", TmpRs!BagSize)
        .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!ContractID), "", TmpRs!ContractID)
        .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!NContractID), "", TmpRs!NContractID)
        .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID)
        .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!CDTFDate), "", Format(TmpRs!CDTFDate, "DD-MMM-YYYY"))
        .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID)
        .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!MadeupFlag), "", TmpRs!MadeupFlag)
        
        .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
        .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!HologramNo), "", TmpRs!HologramNo)
        .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!TransType), "", TmpRs!TransType)
        
        
        TmpRs.MoveNext
        ProgressBar1.Value = mR
        .Rows = .Rows + 1
    End With
Next mR
End If
LblTotalRecord.Caption = TmpRs.RecordCount
MsgBox "Done."
CmdExcel.Enabled = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call ClearData
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_PendingGSLforRateCard.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_PendingGSLforRateCard.xls")
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
           If C = 5 Then
            oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           Else
            oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub ClearData()
Call Grid_Head
LblTotalRecord.Caption = ""
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 20 '17
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "State Name":                   .ColWidth(0) = 2000
    .TextMatrix(0, 1) = "Location Name":                .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "SM-Prefix":                    .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "CMP Name":                     .ColWidth(3) = 2000
    .TextMatrix(0, 4) = "GSL ID"::                      .ColWidth(4) = 1500
    .TextMatrix(0, 5) = "Godown No":                    .ColWidth(5) = 1000
    .TextMatrix(0, 6) = "Commodity":                    .ColWidth(6) = 2000
    .TextMatrix(0, 7) = "Client":                       .ColWidth(7) = 2000
    .TextMatrix(0, 8) = "Balance Bags":                 .ColWidth(8) = 1500
    .TextMatrix(0, 9) = "Balance Weight":               .ColWidth(9) = 1500
    .TextMatrix(0, 10) = "Bag Size":                    .ColWidth(10) = 1200
    .TextMatrix(0, 11) = "Contract ID":                 .ColWidth(11) = 1200
    .TextMatrix(0, 12) = "NContract ID":                .ColWidth(12) = 1500
    .TextMatrix(0, 13) = "G_UID":                       .ColWidth(13) = 2500
    .TextMatrix(0, 14) = "CDTF Date":                   .ColWidth(14) = 1500
    .TextMatrix(0, 15) = "CDTF No":                     .ColWidth(15) = 1200
    .TextMatrix(0, 16) = "Made up Flag":                .ColWidth(16) = 1500
    
    .TextMatrix(0, 17) = "Print Flag":                  .ColWidth(16) = 1500
    .TextMatrix(0, 18) = "Hologram":                    .ColWidth(17) = 1500
    .TextMatrix(0, 19) = "Transaction Type":            .ColWidth(18) = 1500
    
    .RowHeight(0) = 600
    .Font.Size = 10
End With
End Sub
