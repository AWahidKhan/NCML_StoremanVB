VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   8010
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   1920
      Picture         =   "Form1.frx":45A3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1695
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   840
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   240
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton CmdMaster 
      Caption         =   "Delete All Master "
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete All Transaction"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3855
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdMaster_Click()
Call Command2_Click

tmp = "delete from Mst_StorageContract"
Call TmpTable
tmp = "delete from Mst_RRN"
Call TmpTable
tmp = "delete from Mst_GSL"
Call TmpTable
tmp = "delete from Mst_Godown"
Call TmpTable
tmp = "delete from Mst_Assayer "
Call TmpTable
tmp = "delete from Mst_ISIN"
Call TmpTable
tmp = "delete from Mst_Branch"
Call TmpTable
tmp = "delete from Mst_Bank"
Call TmpTable
tmp = "delete from Mst_CMP"
Call TmpTable
'tmp = "delete from Mst_Category"
'Call TmpTable
tmp = "delete from Mst_Client"
Call TmpTable
tmp = "delete from Mst_CommodityDetail"
Call TmpTable
tmp = "delete from Mst_Commodity"
Call TmpTable
tmp = "delete from Mst_Depository"
Call TmpTable
tmp = "delete from Mst_InsuranceDetails"
Call TmpTable
tmp = "delete from Mst_InsuranceCoMaster"
Call TmpTable

tmp = "delete from Mst_License"
Call TmpTable
tmp = "delete from Mst_Location"
Call TmpTable
tmp = "delete from Mst_State"
Call TmpTable

tmp = "delete from Mst_StoreManLocations "
Call TmpTable

'tmp = "delete from Mst_UnitType"
'Call TmpTable
'tmp = "delete from Mst_UserRole"
'Call TmpTable
'tmp = "delete from Mst_Users"
'Call TmpTable

Call TableMst_Depository
RsMst_Depository.AddNew
RsMst_Depository!G_UID = GetGUID
RsMst_Depository!DepositoryID = 1
RsMst_Depository!DepositoryName = "NCMSL"
RsMst_Depository!ExchangeTypeID = 1
RsMst_Depository.Update


End Sub

Private Sub Command2_Click()
tmp = "delete from Txn_CDC_Assayer_Detail"
Call TmpTable
tmp = "delete from Txn_CDC_GSL_Detail"
Call TmpTable
tmp = "delete from Txn_CDC_Insurance_Detail"
Call TmpTable
tmp = "delete from Txn_CDC"
Call TmpTable
tmp = "delete from Txn_CxTF_ClientDetail"
Call TmpTable
tmp = "delete from Txn_CxTF_GSL"
Call TmpTable
tmp = "delete from Txn_CxTF_TruckDetail"
Call TmpTable
tmp = "delete from Txn_CxTF_Details"
Call TmpTable
tmp = "delete from Txn_ISINGSLRevalidation"
Call TmpTable
tmp = "delete from Txn_ISINRevalidation"
Call TmpTable
tmp = "delete from Txn_WHR_Detail"
Call TmpTable
tmp = "delete from Txn_CashReceipt"
Call TmpTable


tmp = "delete from Txn_Spillage"
Call TmpTable
tmp = "delete from Txn_PledgeCreation_GSL"
Call TmpTable
tmp = "delete from Txn_PledgeCreation"
Call TmpTable

tmp = "delete from Txn_PledgeRelease_GSL"
Call TmpTable
tmp = "delete from Txn_PledgeRelease"
Call TmpTable

tmp = "delete from Mst_Hologram "
Call TmpTable

tmp = "delete from Txn_AdjTransaction "
Call TmpTable

tmp = "delete from Mst_DR "
Call TmpTable

tmp = "delete from Txn_GSL_Transfer "
Call TmpTable

tmp = "delete from Log_Print_History "
Call TmpTable




tmp = "update Mst_GSL set DepositeWeight = 0, " & _
"DepositeBags = 0, " & _
"WHRBags = 0, " & _
"CDTFBags = 0, " & _
"BalanceBags = 0, " & _
"WHRWeight = 0, " & _
"CDTFWeight = 0, " & _
"BalanceWeight = 0, " & _
"ExchangeTypeID = 0, " & _
"ClientIDRow = 0, " & _
"CommodityID = Null , " & _
"ISINID = 0, " & _
"FirstISINID = 0, " & _
"PledgeBags = 0, " & _
"PledgeWeight = 0, " & _
"PledgeNo = Null, " & _
"CDTFNo = Null, " & _
"DepositAdjWt = 0, " & _
"WithdrawalAdjWt = 0, " & _
"CDTFDate = Null,KRDate = Null, " & _
"BGSLID= Null ,FGSLID = Null ," & _
"LCDTF_Date = Null,GrossWt = Null,KRNo = Null,KRNoBatchID = Null, " & _
"Flag = 'B' "
Call TmpTable


tmp = "update Mst_RRn set DeliveredQty = 0 "
Call TmpTable
MsgBox "Done !!! "

End Sub



Private Sub Form_Load()
'SELECT table_name,ordinal_position,column_name,data_type,  is_nullable,character_maximum_length,* FROM
'information_schema.Columns
'ORDER BY table_name,ordinal_position

If UserType = "Power" Then
   
Else
   Command1.Enabled = False
   Command2.Enabled = False
   CmdMaster.Enabled = False
   MsgBox "Access denied!!!"
   Exit Sub
End If

Dim conn1 As ADODB.Connection
Dim rs1 As ADODB.Recordset
Set conn1 = New ADODB.Connection
conn1.ConnectionString = "Provider=MSDataShape.1;data provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Store-Man.mdb;Jet OLEDB:Database Password=M710211#S600607;" & " Persist Security Info=False"
conn1.Open
Set rs1 = New ADODB.Recordset
'rs1.Open "SHAPE  {SELECT * from Mst_State order by StateID} " & _
'"append (SHAPE  {SELECT * from Mst_Location}   " & _
'"append ({SELECT * from Mst_cmp} relate LocationID to LocationID) relate StateID to StateID)  ", conn1, adOpenDynamic, adLockOptimistic


rs1.Open "   SHAPE  {SELECT * from Mst_State order by StateID } " & _
"   APPEND ((SHAPE  {select * from Mst_Location} " & _
"            APPEND ({select * from [Mst_cmp]}  " & _
"                   RELATE LocationID TO LocationID))  " & _
"          RELATE StateID TO StateID) ", conn1, adOpenDynamic, adLockOptimistic
          
          
Set MSHFlexGrid1.DataSource = rs1
rs1.Close
Set rs1 = Nothing
conn1.Close
Set conn1 = Nothing
With MSHFlexGrid1
.CollapseAll
.BandDisplay = flexBandDisplayVertical
.ColHeader(1) = flexColHeaderOn
.BandIndent(1) = 2
End With
End Sub

