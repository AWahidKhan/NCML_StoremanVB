VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_WithdrawalBOEProcess 
   Caption         =   "Withdrawal BOE Process"
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
      Height          =   1470
      Left            =   240
      TabIndex        =   5
      Top             =   30
      Width           =   14535
      Begin VB.ComboBox CmbStateID 
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
         Left            =   6960
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   3600
      End
      Begin VB.ComboBox CmbExchange 
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
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   3570
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Process BOE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10815
         TabIndex        =   3
         Top             =   960
         Width           =   2370
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   5520
         TabIndex        =   8
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Exchange"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   $"Frm_WithdrawalBOEProcess.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   6
         Top             =   150
         Width           =   12855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   2355
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   225
      TabIndex        =   4
      Top             =   1515
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "Frm_WithdrawalBOEProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityGroupId As Variant
Dim mPostDepoFromDate As Date
Dim mLastDateofCDTF As Date
Dim mExchangeID, mStateID As Variant
    
Private Sub CmbExchange_GotFocus()
tmp = CmbExchange.Text
Call TableMst_ExchangeType(" Where ExchangeTypeID in (1,2) ")

CmbExchange.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Exchange found"
   CmbExchange.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchange.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchange.Text = tmp

End Sub

Private Sub CmbExchange_LostFocus()
If CmbExchange.Text = "" Then
   mExchangeID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchange.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid Exchange selection "
   CmbExchange.SetFocus
   Exit Sub
End If
mExchangeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmbStateID_GotFocus()
If CmbExchange.Text = "" Then
   MsgBox "Select exchange!!!"
   CmbExchange.SetFocus
   Exit Sub
End If


tmp1 = CmbStateID.Text

tmp = "Select distinct MS.StateName,MS.StateID "
tmp = tmp & " from Txn_CxTF_Details TCD"
tmp = tmp & " left Join Txn_CxTF_BOE TCB On TCD.CxTFNo = TCB.CxTFNo And TCD.SM_Prefix = TCB.SM_Prefix"
tmp = tmp & " left join Mst_CMP CMP on CMP.CMPID = TCD.CMPID"
tmp = tmp & " Left Join Mst_Location ML On CMP.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Where TCD.CxtfDate >= '2011-04-01' And TxnType='W' and BOEID is null "
If mExchangeID = 1 Then
   tmp = tmp & " and (TCD.ExchangeTypeID=1)  "
ElseIf mExchangeID = 2 Then
   tmp = tmp & " and (TCD.ExchangeTypeID=2 and TCD.BookedSpaceFlag=1) "
End If
tmp = tmp & " Order By MS.StateName"

Call TmpTable

CmbStateID.Clear

If TmpRs.RecordCount = 0 Then
   MsgBox "No State found"
   CmbStateID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbStateID.AddItem TmpRs!StateName
    TmpRs.MoveNext
Next

CmbStateID.Text = tmp1

End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   mStateID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change

TmpRs.MoveFirst
TmpRs.Find "StateName = '" & CmbStateID.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid State Selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = TmpRs!StateID
End Sub

''''''''''''''''''''Remember To Update Urad Dal Commodity '''''''''''''''''''''''''''''''''''''

'tmp = " Update Mst_Commodity set CommodityGroupID = 21 Where CommodityID = 23 "
'call tmptable
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CmdSave_Click()

On Error GoTo Error

If CmbStateID.Text = "" Then
   MsgBox "Blank State not allowed!!!"
   CmbStateID.SetFocus
   Exit Sub
End If

If CmbExchange.Text = "" Then
   MsgBox "Blank Exchange not allowed!!!"
   CmbExchange.SetFocus
   Exit Sub
End If

CmdSave.Enabled = False

ProgressBar1.Value = 0

tmp = " Select distinct TCD.CxTFNo,TCD.SM_Prefix,TCD.CxTFDate,BOEID , TCD.CommodityID, TCD.CMPID, TCD.ExchangeTypeID, "
tmp = tmp & " TCD.CxTFDate , CMP.LocationID "
tmp = tmp & " from Txn_CxTF_Details TCD "
tmp = tmp & " left Join Txn_CxTF_BOE TCB On TCD.CxTFNo = TCB.CxTFNo And TCD.SM_Prefix = TCB.SM_Prefix "
tmp = tmp & " left join Mst_CMP CMP on CMP.CMPID = TCD.CMPID "
tmp = tmp & " Left Join Mst_Location ML On CMP.LocationID = ML.LocationID "
tmp = tmp & " Left Join Mst_State MS On ML.StateID = MS.StateID "
tmp = tmp & " Where TCD.CxtfDate>='2011-04-01'  And TxnType='W' and BOEID is null " 'and (TCD.ExchangeTypeID=1 or (TCD.ExchangeTypeID=2 and TCD.BookedSpaceFlag=1)) "
tmp = tmp & " And MS.StateID = " & mStateID & " "
If mExchangeID = 1 Then
   tmp = tmp & " And TCD.ExchangeTypeID = 1 "
ElseIf mExchangeID = 2 Then
   tmp = tmp & " And (TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1) "
End If

tmp = tmp & " Order By TCD.SM_Prefix,TCD.CxTFDate,TCD.CxTFNo "

Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
    ProgressBar1.Value = 0
    ProgressBar1.Max = TmpRs3.RecordCount + 2

    For k = 1 To TmpRs3.RecordCount
        Call FillBOE
        Call SaveCxTFBOE
        TmpRs3.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Next
    ProgressBar1.Value = ProgressBar1.Max
End If

Call Grid_Head_Bill
MsgBox "Done"
CmbStateID.Text = ""
CmbExchange.Text = ""
CmdSave.Enabled = True
Exit Sub

Error:

    Call Grid_Head_Bill
    MsgBox "Some error occured at  . Please Check!!!"
    CmbStateID.Text = ""
    CmbExchange.Text = ""
    CmdSave.Enabled = True

End Sub
Private Sub Grid_Head_Bill()

MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 2
MSHFlexGrid6.FixedRows = 1
MSHFlexGrid6.Cols = 22

MSHFlexGrid6.TextMatrix(0, 0) = "CITF Date"
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

MSHFlexGrid6.TextMatrix(0, 17) = "Deposit Billing Cycle"
MSHFlexGrid6.TextMatrix(0, 18) = "Depo No. of Cycle"
MSHFlexGrid6.TextMatrix(0, 19) = "Depo Rate per Cycle"
MSHFlexGrid6.TextMatrix(0, 20) = "Depo Storage Charges Amount"
MSHFlexGrid6.TextMatrix(0, 21) = "Total Storage Charges Amount"


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
MSHFlexGrid6.ColWidth(17) = 1600
MSHFlexGrid6.ColWidth(18) = 1600
MSHFlexGrid6.ColWidth(19) = 1600
MSHFlexGrid6.ColWidth(20) = 1600

MSHFlexGrid6.RowHeight(0) = 600

End Sub
Private Sub DeleteCxTFBOE()
tmp = "Delete from Txn_CxTF_BOE Where CxTFNo=" & TmpRs3!CxTFNo & " And LocationID=" & TmpRs3!LocationID
Call TmpTable
End Sub

Private Sub SaveCxTFBOE()

Call DeleteCxTFBOE

For I = 1 To MSHFlexGrid6.Rows - 2
    If MSHFlexGrid6.TextMatrix(I, 0) = "" Then Exit For
    Call TableTxn_CxTF_BOE(" where CxTFNo=" & TmpRs3!CxTFNo & " And LocationID=" & TmpRs3!LocationID)
    If RsTxn_CxTF_BOE.RecordCount > 0 Then
       RsTxn_CxTF_BOE.MoveFirst
       RsTxn_CxTF_BOE.Find "ClientIDRow= " & Val(MSHFlexGrid6.TextMatrix(I, 3))
    End If
    If RsTxn_CxTF_BOE.EOF = False Then
       'RsTxn_CxTF_BOE!CxTFDate = MSHFlexGrid6.TextMatrix(i, 3)
       RsTxn_CxTF_BOE!WithdrawalBags = IIf(IsNull(RsTxn_CxTF_BOE!WithdrawalBags), 0, RsTxn_CxTF_BOE!WithdrawalBags) + Val(MSHFlexGrid6.TextMatrix(I, 8))
       RsTxn_CxTF_BOE!WithdrawalQty = IIf(IsNull(RsTxn_CxTF_BOE!WithdrawalQty), 0, RsTxn_CxTF_BOE!WithdrawalQty) + Val(MSHFlexGrid6.TextMatrix(I, 9))
       RsTxn_CxTF_BOE!BOEAmount = IIf(IsNull(RsTxn_CxTF_BOE!BOEAmount), 0, RsTxn_CxTF_BOE!BOEAmount) + Val(MSHFlexGrid6.TextMatrix(I, 21))
       RsTxn_CxTF_BOE.Update
    Else
       RsTxn_CxTF_BOE.AddNew
       RsTxn_CxTF_BOE!BOEID = getMaxBOEID
       RsTxn_CxTF_BOE!CxTFNo = TmpRs3!CxTFNo
       RsTxn_CxTF_BOE!LocationID = TmpRs3!LocationID
       RsTxn_CxTF_BOE!CommodityID = TmpRs3!CommodityID
       RsTxn_CxTF_BOE!SM_Prefix = TmpRs3!SM_Prefix
       RsTxn_CxTF_BOE!ClientIDRow = Val(MSHFlexGrid6.TextMatrix(I, 3))
       RsTxn_CxTF_BOE!CxTFDate = MSHFlexGrid6.TextMatrix(I, 0)
       RsTxn_CxTF_BOE!WithdrawalBags = Val(MSHFlexGrid6.TextMatrix(I, 8))
       RsTxn_CxTF_BOE!WithdrawalQty = Val(MSHFlexGrid6.TextMatrix(I, 9))
       RsTxn_CxTF_BOE!BOEAmount = Val(MSHFlexGrid6.TextMatrix(I, 21))
       RsTxn_CxTF_BOE!ExchangeTypeID = TmpRs3!ExchangeTypeID
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
Private Sub FillBOE()

Dim I As Variant

tmp = "Select CommodityGroupID From Mst_Commodity Where CommodityID=" & TmpRs3!CommodityID
Call TmpTable
If TmpRs.RecordCount > 0 Then
   mCommodityGroupId = TmpRs!CommodityGroupID
End If

Call Grid_Head_Bill

If TmpRs3!ExchangeTypeID = "1" Then
   
   tmp = "Select TC.CxTFDate,GSL.CDTFDate,datediff(day,GSL.CDTFDate,TC.CxTFDate)+ 1 NoOfDay,GSL.ClientIDRow,MC.ClientName,TCG.GSLID,MG.GodownNo+'~'+"
   tmp = tmp & " GSL.StackNo +'~'+ GSL.LotNo As GSL,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID"
   tmp = tmp & " From  Txn_CxTF_GSL TCG"
   tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTFNo=TC.CxTFNo And TCG.TxnType=TC.TxnType"
   tmp = tmp & " left Join Mst_GSL GSL on TCG.GSLID=GSL.GSLID And TCG.SM_Prefix=GSL.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
   tmp = tmp & " Inner Join Mst_Client MC on GSL.ClientIDRow=MC.ClientIDRow"
   tmp = tmp & " Where TC.ExchangeTypeID=1 And TC.CxTFNo=" & TmpRs3!CxTFNo & " And TC.TxnType='W' And TC.SM_Prefix='" & TmpRs3!SM_Prefix & "' And GSL.SpillageFlag = 0 "
    
    
   Call TmpTable
    
   If TmpRs.RecordCount > 0 Then
        
        
      For I = 1 To TmpRs.RecordCount
          For c = 0 To TmpRs.Fields.Count - 1
              MSHFlexGrid6.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
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
             MSHFlexGrid6.TextMatrix(I, 20) = Format(Val(MSHFlexGrid6.TextMatrix(I, 18)) * Val(MSHFlexGrid6.TextMatrix(I, 19)) * Val(MSHFlexGrid6.TextMatrix(I, 8)), "#####0.00")
             MSHFlexGrid6.TextMatrix(I, 21) = Val(MSHFlexGrid6.TextMatrix(I, 16)) + Val(MSHFlexGrid6.TextMatrix(I, 20))
          Else
             MSHFlexGrid6.TextMatrix(I, 16) = Format(Val(MSHFlexGrid6.TextMatrix(I, 14)) * Val(MSHFlexGrid6.TextMatrix(I, 15)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
             MSHFlexGrid6.TextMatrix(I, 20) = Format(Val(MSHFlexGrid6.TextMatrix(I, 18)) * Val(MSHFlexGrid6.TextMatrix(I, 19)) * Val(MSHFlexGrid6.TextMatrix(I, 9)), "#####0.00")
             MSHFlexGrid6.TextMatrix(I, 21) = Val(MSHFlexGrid6.TextMatrix(I, 16)) + Val(MSHFlexGrid6.TextMatrix(I, 20))
          End If
          TmpRs.MoveNext
          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
      Next
   End If


ElseIf TmpRs3!ExchangeTypeID = 2 Then

   tmp = " Select TCD.CxTFDate,MG.CDTFDate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfDay,MG.ClientIDRow,MClient.ClientName,"
   tmp = tmp & " TCG.GSLID, MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as GSL,MG.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,0 as 'BagSize',MG.NContractID"
   tmp = tmp & " ,'Daily' Cycle,'Qty' Unit,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) NoOfCycle,MNRC.Rate,((datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) - 7) * TCG.DematWeight * Rate as 'Amount'"
   tmp = tmp & " From Txn_CxTF_Details TCD"
   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
   tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientIDRow = MClient.ClientIDRow"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
   tmp = tmp & " Where TCD.TxnType = 'W' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1 And TCD.CxTFNo=" & TmpRs3!CxTFNo & " And TCD.SM_Prefix='" & TmpRs3!SM_Prefix & "' And MG.SpillageFlag = 0 "
   
   Call TmpTable
    
   If TmpRs.RecordCount > 0 Then
      For I = 1 To TmpRs.RecordCount
          For c = 0 To TmpRs.Fields.Count - 1
              MSHFlexGrid6.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
          Next
          MSHFlexGrid6.TextMatrix(I, 0) = Format(MSHFlexGrid6.TextMatrix(I, 0), "dd-MMM-yyyy")
          MSHFlexGrid6.TextMatrix(I, 1) = Format(MSHFlexGrid6.TextMatrix(I, 1), "dd-MMM-yyyy")
          If MSHFlexGrid6.TextMatrix(I, 7) = "1" Then
             MSHFlexGrid6.TextMatrix(I, 7) = "Spillage"
          Else
             MSHFlexGrid6.TextMatrix(I, 7) = "Normal"
          End If
          
          If Val(MSHFlexGrid6.TextMatrix(I, 2)) < 0 Then
             MSHFlexGrid6.TextMatrix(I, 2) = 0
             MSHFlexGrid6.TextMatrix(I, 14) = 0
             MSHFlexGrid6.TextMatrix(I, 16) = 0
          End If
          
          TmpRs.MoveNext
          MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
      Next
   End If
End If

End Sub

Private Sub FillBillingDetails(Row_ As Variant, mFromDate As Date, mToDate As Date)
Dim mMonth, mYear As Variant

If Val(MSHFlexGrid6.TextMatrix(Row_, 11)) <> 0 Then Exit Sub

tmp = "Select SPRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,BillingUnit,Amount,MB1.BillingCycle as DepoBillingCycle,MS.DepoBillingCycleID,MB1.NoofDays as DepoNoOfDays"
tmp = tmp & " From Mst_SpecialRateCard MS"
tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID"
tmp = tmp & " Inner Join Mst_BillingCycle MB1 on MS.DepoBillingCycleID=MB1.BillingCycleID "

tmp = tmp & " Where ClientIDRow = " & Val(MSHFlexGrid6.TextMatrix(Row_, 3)) & " And CommodityGroupID = " & mCommodityGroupId & " And LocationID = " & TmpRs3!LocationID & "  And ExchangeTypeID = 1 And MS.Flag = 'A'"

Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid6.TextMatrix(Row_, 12) = TmpRs2!BillingCycle
   MSHFlexGrid6.TextMatrix(Row_, 13) = IIf(TmpRs2!BillingUnit = "B", "Bag", "Qty")
   MSHFlexGrid6.TextMatrix(Row_, 17) = TmpRs2!DepoBillingCycle
    'MSHFlexGrid6.TextMatrix(Row_, 14) = Val(MSHFlexGrid6.TextMatrix(Row_, 2)) / TmpRs2!NoofDays
    'MSHFlexGrid6.TextMatrix(Row_, 14) = CDbl(CLng(Val(MSHFlexGrid6.TextMatrix(Row_, 14)) + 0.5))
    
   mLastDateofCDTF = GetLastDateOfMonth(mFromDate)
   mMonth = Month(mFromDate)
   mYear = Year(mFromDate)
            
            
   If mMonth = "12" Then
      mMonth = "01"
      mYear = mYear + 1
   Else
      mMonth = mMonth + 1
   End If
   
     
   mPostDepoFromDate = CDate(mMonth & "-" & "01" & "-" & mYear)
   MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mPostDepoFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoofDays)
   MSHFlexGrid6.TextMatrix(Row_, 18) = CalculateDepoNoOfCycle(mFromDate, mLastDateofCDTF, TmpRs2!DepoBillingCycleID, TmpRs2!DepoNoofDays)
   
'   MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoofDays)

   If TmpRs2!BillingUnit = "Q" Then
      MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs2!Amount
      MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs2!Amount
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
      
         tmp = " Select RatePerBag From Mst_SpecialRateCardDepoDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " >= FromBagSize And  SPRCID = " & TmpRs2!SPRCID & " "
         Call Tmp1Table
         If TmpRs1.RecordCount > 0 Then
            MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs1!RatePerBag
         End If
      
      End If
      
   End If
Else
    
   tmp = "Select SRCID,MS.BillingCycleID,MB.BillingCycle,MB.NoofDays,MS.Flag,RateOnQty,MB1.BillingCycle as DepoBillingCycle,MS.DepoBillingCycleID,MB1.NoofDays as DepoNoOfDays"
   tmp = tmp & " From Mst_StorageRateCard MS "
   tmp = tmp & " Inner Join Mst_BillingCycle MB on MS.BillingCycleID=MB.BillingCycleID "
   tmp = tmp & " Inner Join Mst_BillingCycle MB1 on MS.DepoBillingCycleID=MB1.BillingCycleID "
   tmp = tmp & " Where CommodityGroupID = " & mCommodityGroupId & " And LocationID =" & TmpRs3!LocationID & "  And ExchangeTypeID = 1 And DefaultRate = 1 "
    
   Call Tmp2Table
    
   If TmpRs2.RecordCount > 0 Then
      MSHFlexGrid6.TextMatrix(Row_, 12) = TmpRs2!BillingCycle
      MSHFlexGrid6.TextMatrix(Row_, 17) = TmpRs2!DepoBillingCycle
      MSHFlexGrid6.TextMatrix(Row_, 13) = IIf(TmpRs2!Flag = "B", "Bag", "Qty")
       'MSHFlexGrid6.TextMatrix(Row_, 14) = Val(MSHFlexGrid6.TextMatrix(Row_, 2)) / TmpRs2!NoOfDays
       'MSHFlexGrid6.TextMatrix(Row_, 14) = CDbl(CLng(Val(MSHFlexGrid6.TextMatrix(Row_, 14)) + 0.5))
      
      mLastDateofCDTF = GetLastDateOfMonth(mFromDate)
      
      
      mMonth = Month(mFromDate)
      mYear = Year(mFromDate)
            
            
      If mMonth = "12" Then
         mMonth = "01"
         mYear = mYear + 1
      Else
         mMonth = mMonth + 1
      End If
      
       
      mPostDepoFromDate = CDate(mMonth & "-" & "01" & "-" & mYear)
      MSHFlexGrid6.TextMatrix(Row_, 14) = CalculateNoOfCycle(mPostDepoFromDate, mToDate, TmpRs2!BillingCycleID, TmpRs2!NoofDays)
      MSHFlexGrid6.TextMatrix(Row_, 18) = CalculateDepoNoOfCycle(mFromDate, mLastDateofCDTF, TmpRs2!DepoBillingCycleID, TmpRs2!DepoNoofDays)
      
      If TmpRs2!Flag = "Q" Then
         MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs2!RateOnQty
         MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs2!RateOnQty
      Else
'         tmp = "Select RatePerBag From Mst_StorageRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
'         Call Tmp1Table
'         If TmpRs1.RecordCount > 0 Then
'            MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
'         End If

         If DateDiff("M", mFromDate, mToDate) = 0 Then
            MSHFlexGrid6.TextMatrix(Row_, 15) = 0
         Else
            tmp = "Select RatePerBag From Mst_StorageRateCardDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
               MSHFlexGrid6.TextMatrix(Row_, 15) = TmpRs1!RatePerBag
            End If
         
            tmp = "Select RatePerBag From Mst_StorageRateCardDepoDetail Where " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & " <= ToBagSize And " & Val(MSHFlexGrid6.TextMatrix(Row_, 10)) & ">=FromBagSize And SRCID=" & TmpRs2!SRCID
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
               MSHFlexGrid6.TextMatrix(Row_, 19) = TmpRs1!RatePerBag
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
   mCycleCount = DateDiff("d", mFromDate_, mToDate_) + 1
'w=2 'Weekly
ElseIf mBillingCycleID_ = 2 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) ' - 1
'   DepositMonthCycles = 4 - Left((DatePart("d", mFromDate_) / mNoOfDays_), 2)
   
'   If Format(DatePart("d", mFromDate_) / mNoOfDays_) = 1 Then
'      DepositMonthCycles = 4 '- 0 'Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
'   ElseIf Format(DatePart("d", mFromDate_) / mNoOfDays_) > 4 Then
'      DepositMonthCycles = 4 - (Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) - 1)
'   Else
      
'      If DatePart("d", mFromDate_) Mod mNoOfDays_ = 0 Then
'         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) + 1
'      Else
'         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
'      End If
'      DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
'   End If
   
   WithdrawalMonthCycles = RoundUp(DatePart("d", mToDate_) / mNoOfDays_)
   mCycleCount = (MonthCount * 4) + WithdrawalMonthCycles '+ DepositMonthCycles
'F=3 'Fortnightly
ElseIf mBillingCycleID_ = 3 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) '- 1
'   DepositMonthCycles = IIf(DatePart("d", mFromDate_) - mNoOfDays_ >= 0, 1, 2)
   WithdrawalMonthCycles = IIf(DatePart("d", mToDate_) - mNoOfDays_ >= 0, 2, 1)
   mCycleCount = (MonthCount * 2) + WithdrawalMonthCycles '+ DepositMonthCycles
'M=4 'Monthly
ElseIf TmpRs2!BillingCycleID = 4 Then
   mCycleCount = DateDiff("M", mFromDate_, mToDate_) + 1
End If

CalculateNoOfCycle = mCycleCount

End Function

Private Function CalculateDepoNoOfCycle(mFromDate_ As Date, mToDate_ As Date, mBillingCycleID_ As Variant, mNoOfDays_ As Variant)
Dim mCycleCount As Variant
Dim MonthCount, DepositMonthCycles, WithdrawalMonthCycles As Variant
'D=1 Daily
If mBillingCycleID_ = 1 Then
   mCycleCount = DateDiff("d", mFromDate_, mToDate_) + 1
'w=2 'Weekly
ElseIf mBillingCycleID_ = 2 Then
   MonthCount = DateDiff("M", mFromDate_, mToDate_) - 1
'   DepositMonthCycles = 4 - Left((DatePart("d", mFromDate_) / mNoOfDays_), 2)
   
   If Format(DatePart("d", mFromDate_) / mNoOfDays_) = 1 Then
      DepositMonthCycles = 4 '- 0 'Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
   ElseIf Format(DatePart("d", mFromDate_) / mNoOfDays_) > 4 Then
      DepositMonthCycles = 4 - (Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) - 1)
   Else
      
      If DatePart("d", mFromDate_) Mod mNoOfDays_ = 0 Then
         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1) + 1
      Else
         DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
      End If
'      DepositMonthCycles = 4 - Left(Format(DatePart("d", mFromDate_) / mNoOfDays_), 1)
   End If
   
'   WithdrawalMonthCycles = RoundUp(DatePart("d", mToDate_) / mNoOfDays_)
   mCycleCount = DepositMonthCycles '+ WithdrawalMonthCycles
'F=3 'Fortnightly
ElseIf mBillingCycleID_ = 3 Then
   'MonthCount = DateDiff("M", mFromDate_, mToDate_) - 1
   DepositMonthCycles = IIf(DatePart("d", mFromDate_) - mNoOfDays_ >= 0, 1, 2)
'   WithdrawalMonthCycles = IIf(DatePart("d", mToDate_) - mNoOfDays_ >= 0, 2, 1)
   mCycleCount = DepositMonthCycles ' + WithdrawalMonthCycles
'M=4 'Monthly
ElseIf mBillingCycleID_ = 4 Then
   mCycleCount = DateDiff("M", mFromDate_, mToDate_) + 1
End If

CalculateDepoNoOfCycle = mCycleCount

End Function

Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

Private Sub Form_Load()
If LCase(Gen_UserRole) <> "power" Then
   'CmdSave.Enabled = False
   Frame1.Enabled = False
   MsgBox "Access Denied!!"
   Exit Sub
End If

If Gen_UserLoginID <> "2047" And Gen_UserLoginID <> "1076" Then
   'CmdSave.Enabled = False
   Frame1.Enabled = False
   Exit Sub
Else
   Frame1.Enabled = True
End If

End Sub
