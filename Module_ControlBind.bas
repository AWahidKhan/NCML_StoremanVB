Attribute VB_Name = "Module_ControlBind"
Public Function ConvertToyyyymmdd(mDate_ As Variant) As String
Dim Retval As String
Dim splitstring() As String
'RetVal = " "

splitstring = Split(mDate_, "/")
If UBound(splitstring) = 2 Then
   Retval = splitstring(1) & "/" & splitstring(0) & "/" & splitstring(2)
End If
ConvertToyyyymmdd = Retval
End Function

Public Function GetCommodityName(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_Commodity("Where CommodityID = " & CID_)
If Not RsMst_Commodity.EOF Then
  Retval = RsMst_Commodity!Commodity
End If
GetCommodityName = Retval
End Function

Public Function GetCommodityGroupName(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_CommodityGroup("Where CommodityGroupID = " & CID_)
If Not RsMst_CommodityGroup.EOF Then
  Retval = RsMst_CommodityGroup!COmmodityGroup
End If
GetCommodityGroupName = Retval
End Function


Public Function GetAgentName(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_Agent("Where AgentID = " & CID_)
If Not RsMst_Agent.EOF Then
  Retval = RsMst_Agent!AgentName
End If
GetAgentName = Retval
End Function

Public Function GetStateName(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_State("Where StateID = " & CID_)
If Not RsMst_State.EOF Then
  Retval = RsMst_State!StateName
End If
GetStateName = Retval
End Function
Public Function GetStoreManLocationsName(CID_ As Double, Mode_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_StoreManLocations("Where StoreManLocationsID = " & CID_)
If Not RsMst_StoreManLocations.EOF Then
  If Mode_ = "N" Then
     Retval = RsMst_StoreManLocations!StoreManLocations
  Else
     Retval = RsMst_StoreManLocations!SM_Prefix
  End If
End If
GetStoreManLocationsName = Retval
End Function

Public Function GetLocationName(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_Location("Where LocationID = " & CID_)
If Not RsMst_Location.EOF Then
  Retval = RsMst_Location!Locationname
End If
GetLocationName = Retval
End Function

Public Function GetDepositorName(CID_ As String, Mod_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_Depository("Where DepositoryID = '" & CID_ & "'")
If Not RsMst_Depository.EOF Then
   If Mod_ = "N" Then
      Retval = RsMst_Depository!DepositoryName
   ElseIf Mod_ = "NN" Then
      Retval = IIf(IsNull(RsMst_Depository!NAVDepositoryName), "", NAVDepositoryName)
   Else
      Retval = RsMst_Depository!Address
   End If
End If
GetDepositorName = Retval
End Function

Public Function GetGodownID(CID_ As String, CMPID_ As Double) As Double
Dim Retval As Double
Retval = 0
Call TableMst_Godown("Where GodownNo = '" & CID_ & "' And CMPID = " & CMPID_)
If Not RsMst_Godown.EOF Then
   Retval = RsMst_Godown!GodownID
End If
GetGodownID = Retval
End Function


Public Function GetGodownName(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_Godown("Where GodownID = " & CID_)
If Not RsMst_Godown.EOF Then
   If Mod_ = "N" Then
      Retval = RsMst_Godown!GodownNo
   ElseIf Mod_ = "S" Then
      Retval = RsMst_Godown!Area
   ElseIf Mod_ = "ED" Then
      Retval = IIf(IsNull(RsMst_Godown!AgreementEndDate), "", RsMst_Godown!AgreementEndDate)
   Else
      Retval = RsMst_Godown!Address
   End If
End If
GetGodownName = Retval
End Function

Public Function GetAssayerName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_Assayer(" Where AssayerID = " & CID_)
If Not RsMst_Assayer.EOF Then
  Retval = RsMst_Assayer!AssayerName
End If
GetAssayerName = Retval
End Function

Public Function GetUnitName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_UnitType(" Where UnitTypeID = " & CID_)
If Not RsMst_UnitType.EOF Then
   Retval = RsMst_UnitType!UnitType
End If
GetUnitName = Retval
End Function
Public Function GetLicenseNo(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_License(" Where LicenseID = " & CID_)
If Not RsMst_License.EOF Then
   Retval = RsMst_License!LicenseNo
End If
GetLicenseNo = Retval
End Function


Public Function GetUnitID(CID_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_UnitType(" Where UnitType = '" & CID_ & "'")
If Not RsMst_UnitType.EOF Then
   Retval = RsMst_UnitType!UnitTypeID
End If
GetUnitID = Retval
End Function

Public Function GetCategoryName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_Category(" Where CategoryID = " & CID_)
If Not RsMst_Category.EOF Then
   Retval = RsMst_Category!Category
End If
GetCategoryName = Retval
End Function
Public Function GetCategoryID(CID_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_Category(" Where Category = '" & CID_ & "'")
If Not RsMst_Category.EOF Then
   Retval = RsMst_Category!CategoryID
End If
GetCategoryID = Retval
End Function
Public Function GetCMPName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_CMP(" Where CMPID = " & CID_)
If Not RsMst_CMP.EOF Then
  Retval = RsMst_CMP!CMPName
End If
GetCMPName = Retval
End Function

Public Function GetCMPID(CID_ As String) As Double
Dim Retval As String
Retval = 0
Call TableMst_CMP(" Where CMPName = '" & CID_ & "'")
If Not RsMst_CMP.EOF Then
  Retval = RsMst_CMP!CMPID
End If
GetCMPID = Retval
End Function
'GetDepositorID
Public Function GetDepositoryID(CID_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_Depository(" Where DepositoryName = '" & CID_ & "'")
If Not RsMst_Depository.EOF Then
  Retval = RsMst_Depository!DepositoryID
End If
GetDepositoryID = Retval
End Function
Public Function GetDepositoryName(CID_ As String, ExgType As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_Depository(" Where DepositoryID = '" & CID_ & "' And ExchangeTypeID = " & ExgType)
If Not RsMst_Depository.EOF Then
  Retval = RsMst_Depository!DepositoryName
End If
GetDepositoryName = Retval
End Function

Public Function GetContract(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_StorageContract(" Where ContractID = " & CID_)
If Not RsMst_StorageContract.EOF Then
  Retval = RsMst_StorageContract!ContractName
End If
GetContract = Retval
End Function

Public Function GetExchangeName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_ExchangeType(" Where ExchangeTypeID = " & CID_)
If Not RsMst_ExchangeType.EOF Then
  Retval = RsMst_ExchangeType!ExchangeType
End If
GetExchangeName = Retval
End Function

Public Function GetGSLByGLSID(CID_ As Double, Prefix_ As String) As String()
Dim Retval(18) As String
Retval(0) = "" 'CMP ID
Retval(1) = "" 'CMP Name
Retval(2) = "" 'Godown ID
Retval(3) = "" 'Godown No
Retval(4) = "" 'Stack No
Retval(5) = "" 'Lot No
Retval(6) = "" 'CloseDate
Retval(7) = "" 'CommodityID
Retval(8) = "" 'ExchangeTypeID
Retval(9) = "" 'ClientIDRow
Retval(10) = "" 'ISINID
Retval(11) = "" 'GodownAddress
Retval(12) = "" 'CDTF No
Retval(13) = "" 'PledgeBags
Retval(14) = "" 'PledgeWeight
Retval(15) = "" 'BalanceBags
Retval(16) = "" 'BalanceWt
Retval(17) = "" 'ContractID

      
'tmp = "SELECT Mst_CMP.CMPID, Mst_CMP.CMPName, Mst_Godown.GodownID, Mst_Godown.GodownNo, Mst_GSL.StackNo, " & _
'"Mst_GSL.LotNo, Mst_Godown.CloseDate,Mst_GSL.CommodityID,Mst_GSL.ExchangeTypeID,Mst_GSL.ClientIDRow,Mst_GSL.ISINID " & _
'", Mst_Godown.Address,Mst_GSL.CDTFNo,Mst_GSL.PledgeBags,Mst_GSL.PledgeWeight, Mst_GSL.BalanceBags , Mst_GSL.BalanceWeight, Mst_GSL.NContractID " & _
'"FROM Mst_Godown INNER JOIN " & _
'" (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) " & _
'"ON Mst_Godown.GodownID = Mst_GSL.GodownID WHERE (((Mst_GSL.GSLID)=" & CID_ & ")) And Mst_GSL.Sm_prefix = '" & Prefix_ & "'"
     
     
tmp = "SELECT Mst_CMP.CMPID, Mst_CMP.CMPName, Mst_Godown.GodownID, Mst_Godown.GodownNo, Mst_GSL.StackNo, " & _
"Mst_GSL.LotNo, Mst_Godown.CloseDate,Mst_GSL.CommodityID,Mst_GSL.ExchangeTypeID,Mst_GSL.ClientID,Mst_GSL.ISINID " & _
", Mst_Godown.Address,Mst_GSL.CxTF_DetailsID,Mst_GSL.PledgeBags,Mst_GSL.PledgeWeight, Mst_GSL.BalanceBags , Mst_GSL.BalanceWeight, Mst_GSL.NContractID " & _
"FROM Mst_Godown INNER JOIN " & _
" (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) " & _
"ON Mst_Godown.GodownID = Mst_GSL.GodownID WHERE (((Mst_GSL.Mst_GSL_ID)=" & CID_ & ")) And Mst_GSL.Sm_prefix = '" & Prefix_ & "'"
     
Call Tmp3Table

If Not TmpRs3.EOF Then
  Retval(0) = TmpRs3!CMPID
  Retval(1) = TmpRs3!CMPName
  Retval(2) = TmpRs3!GodownID
  Retval(3) = TmpRs3!GodownNo
  Retval(4) = TmpRs3!StackNo
  Retval(5) = TmpRs3!LotNo
  Retval(6) = IIf(IsNull(TmpRs3!CloseDate), "", TmpRs3!CloseDate)
  Retval(7) = IIf(IsNull(TmpRs3!CommodityID), "", TmpRs3!CommodityID)
  Retval(8) = IIf(IsNull(TmpRs3!ExchangeTypeID), "", TmpRs3!ExchangeTypeID)
  Retval(9) = IIf(IsNull(TmpRs3!ClientID), "", TmpRs3!ClientID)
  Retval(10) = IIf(IsNull(TmpRs3!ISINID), "", TmpRs3!ISINID)
  Retval(11) = IIf(IsNull(TmpRs3!Address), "", TmpRs3!Address)
  Retval(12) = IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
  Retval(13) = IIf(IsNull(TmpRs3!PledgeBags), "", TmpRs3!PledgeBags)
  Retval(14) = IIf(IsNull(TmpRs3!PledgeWeight), "", TmpRs3!PledgeWeight)
  Retval(15) = IIf(IsNull(TmpRs3!BalanceBags), 0, TmpRs3!BalanceBags)
  Retval(16) = IIf(IsNull(TmpRs3!BalanceWeight), 0, TmpRs3!BalanceWeight)
  Retval(17) = IIf(IsNull(TmpRs3!NContractID), 0, TmpRs3!NContractID)
End If
TmpRs3.Close
GetGSLByGLSID = Retval
End Function

'Public Function GetGSLByGLSID(CID_ As Double, Prefix_ As String) As String()
'Dim Retval(17) As String
'Retval(0) = "" 'CMP ID
'Retval(1) = "" 'CMP Name
'Retval(2) = "" 'Godown ID
'Retval(3) = "" 'Godown No
'Retval(4) = "" 'Stack No
'Retval(5) = "" 'Lot No
'Retval(6) = "" 'CloseDate
'Retval(7) = "" 'CommodityID
'Retval(8) = "" 'ExchangeTypeID
'Retval(9) = "" 'ClientIDRow
'Retval(10) = "" 'ISINID
'Retval(11) = "" 'GodownAddress
'Retval(12) = "" 'CDTF No
'Retval(13) = "" 'PledgeBags
'Retval(14) = "" 'PledgeWeight
'Retval(15) = "" 'BalanceBags
'Retval(16) = "" 'BalanceWt
'
'
'
'
'tmp = "SELECT Mst_CMP.CMPID, Mst_CMP.CMPName, Mst_Godown.GodownID, Mst_Godown.GodownNo, Mst_GSL.StackNo, " & _
'"Mst_GSL.LotNo, Mst_Godown.CloseDate,Mst_GSL.CommodityID,Mst_GSL.ExchangeTypeID,Mst_GSL.ClientIDRow,Mst_GSL.ISINID " & _
'", Mst_Godown.Address,Mst_GSL.CDTFNo,Mst_GSL.PledgeBags,Mst_GSL.PledgeWeight, Mst_GSL.BalanceBags , Mst_GSL.BalanceWeight " & _
'"FROM Mst_Godown INNER JOIN " & _
'" (Mst_CMP INNER JOIN Mst_GSL ON Mst_CMP.CMPID = Mst_GSL.CMPID) " & _
'"ON Mst_Godown.GodownID = Mst_GSL.GodownID WHERE (((Mst_GSL.GSLID)=" & CID_ & ")) And Mst_GSL.Sm_prefix = '" & Prefix_ & "'"
'
'Call Tmp3Table
'If Not TmpRs3.EOF Then
'  Retval(0) = TmpRs3!CMPID
'  Retval(1) = TmpRs3!CMPName
'  Retval(2) = TmpRs3!godownid
'  Retval(3) = TmpRs3!GodownNo
'  Retval(4) = TmpRs3!StackNo
'  Retval(5) = TmpRs3!LotNo
'  Retval(6) = IIf(IsNull(TmpRs3!CloseDate), "", TmpRs3!CloseDate)
'  Retval(7) = IIf(IsNull(TmpRs3!CommodityID), "", TmpRs3!CommodityID)
'  Retval(8) = IIf(IsNull(TmpRs3!ExchangeTypeID), "", TmpRs3!ExchangeTypeID)
'  Retval(9) = IIf(IsNull(TmpRs3!ClientIDRow), "", TmpRs3!ClientIDRow)
'  Retval(10) = IIf(IsNull(TmpRs3!ISINID), "", TmpRs3!ISINID)
'  Retval(11) = IIf(IsNull(TmpRs3!Address), "", TmpRs3!Address)
'  Retval(12) = IIf(IsNull(TmpRs3!CdtfNo), "", TmpRs3!CdtfNo)
'  Retval(13) = IIf(IsNull(TmpRs3!PledgeBags), "", TmpRs3!PledgeBags)
'  Retval(14) = IIf(IsNull(TmpRs3!PledgeWeight), "", TmpRs3!PledgeWeight)
'  Retval(15) = IIf(IsNull(TmpRs3!BalanceBags), 0, TmpRs3!BalanceBags)
'  Retval(16) = IIf(IsNull(TmpRs3!BalanceWeight), 0, TmpRs3!BalanceWeight)
'
'End If
'TmpRs3.Close
'GetGSLByGLSID = Retval
'End Function

'Public Function GetClientNameByRow(CID_ As Double) As String()
'
'Dim Retval(6) As String 'Dim RetVal(5) As String
'Retval(0) = "" 'Client ID
'Retval(1) = "" 'Client Name
'Retval(2) = "" 'DP ID
'Retval(3) = "" 'DP Name
'Retval(4) = ""
'Retval(5) = ""
'tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.Address,Mst_Client.NAVClientName " & _
'      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
'      "WHERE  Mst_Client.ClientIDRow = " & CID_ & ""
'Call Tmp3Table
'If Not TmpRs3.EOF Then
'  Retval(0) = TmpRs3!ClientID
'  Retval(1) = TmpRs3!ClientName
'  Retval(2) = TmpRs3!DepositoryID
'  Retval(3) = TmpRs3!DepositoryName
'  Retval(4) = IIf(IsNull(TmpRs3!Address), "", TmpRs3!Address)
'  Retval(5) = IIf(IsNull(TmpRs3!NAVClientName), "", TmpRs3!NAVClientName)
'End If
'TmpRs3.Close
'GetClientNameByRow = Retval
'End Function
'Public Function GetClientNameByRow(CID_ As Double) As String()
'
'Dim Retval(8) As String 'Dim RetVal(5) As String
'Retval(0) = "" 'Client ID
'Retval(1) = "" 'Client Name
'Retval(2) = "" 'DP ID
'Retval(3) = "" 'DP Name
'Retval(4) = "" 'Address
'Retval(5) = "" 'NAVName
'Retval(6) = "" 'NACV Code
'Retval(7) = "" 'RNAV Code
'tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.Address,Mst_Client.NAVClientName ,Mst_Client.NAVCode ,Mst_Client.RNAVCode " & _
'      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
'      "WHERE  Mst_Client.ClientIDRow = " & CID_ & ""
'Call Tmp3Table
'If Not TmpRs3.EOF Then
'  Retval(0) = TmpRs3!CLIENTID
'  Retval(1) = TmpRs3!ClientName
'  Retval(2) = TmpRs3!DepositoryID
'  Retval(3) = TmpRs3!DepositoryName
'  Retval(4) = IIf(IsNull(TmpRs3!Address), "", TmpRs3!Address)
'  Retval(5) = IIf(IsNull(TmpRs3!NAVClientName), "", TmpRs3!NAVClientName)
'  Retval(6) = IIf(IsNull(TmpRs3!NAVCode), "", TmpRs3!NAVCode)
'  Retval(7) = IIf(IsNull(TmpRs3!RNAVCode), "", TmpRs3!RNAVCode)
'End If
'TmpRs3.Close
'GetClientNameByRow = Retval
'End Function
Public Function GetClientNameByRow(CID_ As Double) As String()

Dim Retval(9) As String 'Dim RetVal(5) As String
Retval(0) = "" 'Client ID
Retval(1) = "" 'Client Name
Retval(2) = "" 'DP ID
Retval(3) = "" 'DP Name
Retval(4) = "" 'Address
Retval(5) = "" 'NAVName
Retval(6) = "" 'NACV Code
Retval(7) = "" 'RNAV Code
Retval(8) = "" 'KYCFlag

'Parag
'tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.Address,Mst_Client.NAVClientName ,Mst_Client.NAVCode ,Mst_Client.RNAVCode,Mst_Client.KYCFlag " & _
'      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
'      "WHERE  Mst_Client.ClientIDRow = " & CID_ & ""

tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.Address,Mst_Client.NAVClientName ,Mst_Client.NAVCode ,Mst_Client.RNAVCode,Mst_Client.KYCFlag " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE  Mst_Client.ClientID = " & CID_ & ""


Call Tmp3Table

If Not TmpRs3.EOF Then
   'Parag
   'Retval(0) = TmpRs3!ClientID
   Retval(0) = TmpRs3!ClientIDRow
   Retval(1) = TmpRs3!ClientName
   Retval(2) = TmpRs3!DepositoryID
   Retval(3) = TmpRs3!DepositoryName
   Retval(4) = IIf(IsNull(TmpRs3!Address), "", TmpRs3!Address)
   Retval(5) = IIf(IsNull(TmpRs3!NAVClientName), "", TmpRs3!NAVClientName)
   Retval(6) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
   Retval(7) = IIf(IsNull(TmpRs3!RNAVCode), "", TmpRs3!RNAVCode)
   Retval(8) = IIf(IsNull(TmpRs3!KYCFlag), 0, TmpRs3!KYCFlag)
End If

TmpRs3.Close
GetClientNameByRow = Retval

End Function


Public Function GetClientName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_Client(" Where ClientID = " & CID_)
If Not RsMst_Client.EOF Then
  Retval = RsMst_Client!ClientName
End If
GetClientName = Retval
End Function
Public Function GetClientIDRow(CID_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_Client(" Where ClientID = '" & CID_ & "'")
If Not RsMst_Client.EOF Then
  Retval = RsMst_Client!ClientIDRow
End If
GetClientIDRow = Retval
End Function

Public Function GetClientNameDtl(CID_ As String, ExgTyp As Double) As String()
Dim Retval(4, 2) As String
Retval(0, 0) = False ' Muliple Client ID
Retval(1, 0) = ""    ' Client Name
Retval(2, 0) = ""    ' Depository ID
Retval(3, 0) = ""    ' Record Count

Dim dpl As Variant
Call TableMst_Client(" Where ExchangeTypeID = " & ExgTyp & " And ClientIDRow = '" & CID_ & "'")
If Not RsMst_Client.EOF Then
  Retval(0, 0) = True
  Retval(1, 0) = RsMst_Client!ClientName
  Retval(2, 0) = RsMst_Client!DepositoryID
  Retval(3, 0) = RsMst_Client.RecordCount
    
  For i0 = 1 To RsMst_Client.RecordCount
      dpl = dpl & "~" & RsMst_Client!DepositoryID
      RsMst_Client.MoveNext
  Next
  Retval(3, 1) = dpl
  
End If
GetClientNameDtl = Retval
End Function
Public Function GetClientNameDtlFrmCRowID(CID_ As Variant) As String()
Dim Retval(4, 2) As String
Retval(0, 0) = False ' Muliple Client ID
Retval(1, 0) = ""    ' Client Name
Retval(2, 0) = ""    ' Depository ID
Retval(3, 0) = ""    ' Record Count
Dim dpl As Variant
Call TableMst_Client(" Where ClientIDRow = " & CID_ & "")
If Not RsMst_Client.EOF Then
  Retval(0, 0) = True
  Retval(1, 0) = RsMst_Client!ClientName
  Retval(2, 0) = RsMst_Client!DepositoryID
  Retval(3, 0) = RsMst_Client.RecordCount
  For i0 = 1 To RsMst_Client.RecordCount
      dpl = dpl & "~" & RsMst_Client!DepositoryID
      RsMst_Client.MoveNext
  Next
  Retval(3, 1) = dpl
End If
GetClientNameDtlFrmCRowID = Retval
End Function

Public Function GetISINName(CID_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_ISIN(" Where ISINID = '" & CID_ & "'")
If Not RsMst_ISIN.EOF Then
  Retval = RsMst_ISIN!ISINDescription
End If
GetISINName = Retval
End Function
Public Function GetExchangeID(CID_ As String) As Double
Dim Retval As Double
Retval = 0
Call TableMst_ExchangeType(" Where ExchangeType = '" & CID_ & "'")
If Not RsMst_ExchangeType.EOF Then
  Retval = RsMst_ExchangeType!ExchangeTypeID
End If
GetExchangeID = Retval
End Function

Public Function GetDPFromClient(CID_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_Client(" Where ClientId = '" & CID_ & "'")
If Not RsMst_Client.EOF Then
   Retval = RsMst_Client!DepositoryID
End If
GetDPFromClient = Retval
End Function
Public Function GetInsuranceCoName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_InsuranceCo(" Where InsuranceCoId = " & CID_)
If Not RsMst_InsuranceCo.EOF Then
   Retval = RsMst_InsuranceCo!InsuranceCoName
End If
GetInsuranceCoName = Retval
End Function
Public Function GetLocationFromCMP(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_CMP(" Where CMPID = " & CID_)
If Not RsMst_CMP.EOF Then
  Retval = GetLocationName(RsMst_CMP!LocationID)
End If
GetLocationFromCMP = Retval
End Function
Public Function GetLocationIDFromCMP(CID_ As Double) As String
Dim Retval As Double
Retval = 0
Call TableMst_CMP(" Where CMPID = " & CID_)
If Not RsMst_CMP.EOF Then
  Retval = (RsMst_CMP!LocationID)
End If
GetLocationIDFromCMP = Retval
End Function

Public Function GetCMPAddress(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_CMP(" Where CMPID = " & CID_)
If Not RsMst_CMP.EOF Then
  If Mod_ = "A" Then
     Retval = RsMst_CMP!Address
  Else
     Retval = RsMst_CMP!CMPAlias
  End If
End If
GetCMPAddress = Retval
End Function
Public Function GetLicenseID(CID_ As String) As Double
Dim Retval As String
Retval = 0
Call TableMst_License(" Where LicenseNo = '" & CID_ & "'")
If Not RsMst_License.EOF Then
   Retval = RsMst_License!LicenseID
End If
GetLicenseID = Retval
End Function

Public Function GetBankName(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_Bank("Where BankID = " & CID_)
If Not RsMst_Bank.EOF Then
   If Mod_ = "N" Then
      Retval = RsMst_Bank!BankName
   Else
      Retval = RsMst_Bank!Address
   End If
End If
GetBankName = Retval
End Function
Public Function GetBranchName(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_Branch("Where BranchID = " & CID_)
If Not RsMst_Branch.EOF Then
   If Mod_ = "N" Then
      Retval = RsMst_Branch!BranchName
   Else
      Retval = RsMst_Branch!Address
   End If
End If
GetBranchName = Retval
End Function

Public Function Validate_ComBags(CommodityID_ As Variant, bags_ As Double, Weight As Double) As String
Dim Retval As String
Dim WtPerBag As Double
Retval = ""
If bags_ = 0 Then
   WtPerBag = 0
Else
   WtPerBag = Format((Weight * 1000) / bags_, "########0.000")
End If
Call TableMst_Commodity("Where CommodityID = " & CommodityID_)
If Not RsMst_Commodity.EOF Then
   If WtPerBag >= RsMst_Commodity!Minimum And WtPerBag <= RsMst_Commodity!Maximum Then
      Retval = ""
   Else
      Retval = "Weight is " & Str(WtPerBag) & " Kg per bag which is not in the range of minimum (" & RsMst_Commodity!Minimum & ") to maximum (" & Str(RsMst_Commodity!Maximum) & ") set in commodity master. Still you want to continue ? Click 'Yes' to proceed further or 'No' to change weight or no of bags ."
   End If
End If
Validate_ComBags = Retval
End Function

Public Function GetPreFix(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = ""
'Call Mst_StoreManLocations Where BranchID = " & CID_)

'tmp = "SELECT Mst_StoreManLocations.SM_Prefix " & _
'"FROM (Mst_Location INNER JOIN Mst_CMP ON Mst_Location.LocationID = " & _
'"Mst_CMP.LocationID) INNER JOIN Mst_StoreManLocations ON " & _
'"Mst_Location.StoreManLocationsID = Mst_StoreManLocations.StoreManLocationsID " & _
'" WHERE (((Mst_CMP.CMPID)= " & CID_ & "))"

'tmp = "SELECT Mst_StoreManLocations.SM_Prefix,Mst_Location.SiloLocation " & _
'"FROM (Mst_Location INNER JOIN Mst_CMP ON Mst_Location.LocationID = " & _
'"Mst_CMP.LocationID) INNER JOIN Mst_StoreManLocations ON " & _
'"Mst_Location.StoreManLocationsID = Mst_StoreManLocations.StoreManLocationsID " & _
'" WHERE (((Mst_CMP.CMPID)= " & CID_ & "))"

tmp = "Select ML.SM_Prefix,ML.SiloLocation,ML.LocationName,ML.LocationID,ML.StoreManLocationsID From Mst_CMP "
tmp = tmp & " Inner Join Mst_Location ML On Mst_CMP.LocationID = ML.LocationID "
tmp = tmp & " Where Mst_CMP.CMPID = " & Val(CID_) & ""

Call Tmp3Table

If Not TmpRs3.EOF Then
   If Mod_ = "P" Then
      Retval = TmpRs3!SM_Prefix
   ElseIf Mod_ = "S" Then
      Retval = IIf(IsNull(TmpRs3!SiloLocation), False, TmpRs3!SiloLocation)
   Else
      Retval = TmpRs3!StoreManLocationsID
   End If
End If
TmpRs3.Close
GetPreFix = Retval
End Function


Public Function GetPreFixByLocation(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = ""
'Call Mst_StoreManLocations Where BranchID = " & CID_)

tmp = "SELECT Mst_Location.SM_Prefix " & _
"FROM Mst_Location WHERE ((LocationID)= " & CID_ & ")"
Call TmpTable

If Not TmpRs.EOF Then
   If Mod_ = "P" Then
      Retval = TmpRs!SM_Prefix
   Else
      Retval = TmpRs!LocationID
   End If
End If
GetPreFixByLocation = Retval
End Function
Public Function GetStateFromLocation(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_Location(" Where LocationID = " & CID_)
If Not RsMst_Location.EOF Then
  Retval = GetStateName(RsMst_Location!StateID)
End If
GetStateFromLocation = Retval
End Function
Public Function GetLicenseFromGSL(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_GSL(" Where GSLID = " & CID_)
If Not RsMst_Location.EOF Then
  Retval = GetLicenseNo(RsMst_GSL!GodownID)
End If
GetLicenseFromGSL = Retval
End Function

Public Function GetEHOPCMPID(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_CMP(" Where CMPID = " & CID_)
If Not RsMst_CMP.EOF Then
  Retval = RsMst_CMP!EHOPCMPID
End If
GetEHOPCMPID = Retval
End Function

Public Function GetTimeStamp() As String
Dim Retval As String
Retval = 0
Retval = Format(Now, "yyyymmddhhmmss")

'If InStr(RetVal, "-") > 0 Then
'   RetVal = Replace(RetVal, "-", "")
'ElseIf InStr(RetVal, "/") > 0 Then
'   RetVal = Replace(RetVal, "/", "")
'End If
'RetVal = Replace(RetVal, ":", "")
GetTimeStamp = Retval
End Function
Public Sub Log_Print_History(Prefix_, DocType_, DocNo_, FLName_)
Call TableLog_PrintHistory
RsLog_PrintHistory.AddNew
RsLog_PrintHistory!SM_Prefix = Prefix_
RsLog_PrintHistory!DocType = DocType_
RsLog_PrintHistory!DocNo = DocNo_
RsLog_PrintHistory!FileName = FLName_
RsLog_PrintHistory!UserId = Gen_UserLoginID
RsLog_PrintHistory!Concurrency = Now
RsLog_PrintHistory.Update

End Sub
Public Function GetDepositorName_ID(CID_ As String, Mod_ As String, Exg_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_Depository("Where DepositoryID = '" & CID_ & "' and ExchangeTypeID = " & Exg_ & " ")
If Not RsMst_Depository.EOF Then
   If Mod_ = "N" Then
      Retval = RsMst_Depository!DepositoryID & "~" & RsMst_Depository!DepositoryName
   Else
     Retval = RsMst_Depository!Address
   End If
End If
GetDepositorName_ID = Retval
End Function

Public Function GetCMPIDEHOPID(CID_ As Variant) As Double
Dim Retval As String
Retval = 0
Call TableMst_CMP(" Where EHOPCMPID = '" & CID_ & "' ")
If Not RsMst_CMP.EOF Then
  Retval = RsMst_CMP!CMPID
End If
GetCMPIDEHOPID = Retval

End Function

Public Function GetCommodityID(CID_ As String) As Double
Dim Retval As Double
Retval = 0
Call TableMst_Commodity(" Where Commodity = '" & CID_ & "'")
If Not RsMst_Commodity.EOF Then
  Retval = RsMst_Commodity!CommodityID
End If
GetCommodityID = Retval
End Function

'Public Function Check_HologramNo(CID_ As String) As Boolean
'Dim retval As Boolean
'Dim LocName As String
'
'retval = True
'
'Call TableMst_Hologram(" Where HologramNo = '" & CID_ & "' ")
'
'If RsMst_Hologram.RecordCount > 0 Then
'
'   Call TableMst_Location(" Where SM_Prefix = '" & RsMst_Hologram!SM_Prefix & "' ")
'   If RsMst_Location.RecordCount > 0 Then
'      LocName = RsMst_Location!LocationName
'   End If
'   If RsMst_Hologram!DocType = "WHR" Then
'      tmp = "Select CMPID from Txn_WHR_Detail Where WHRNo = " & RsMst_Hologram!DocNo & ""
'      Call Tmp4Table
'      If TmpRs4.RecordCount > 0 Then
'         MsgBox "Hologram Already used in Location : " & LocName & " - " & RsMst_Hologram!DocType & " " & TmpRs4!CMPID & " number :-" & (Val(RsMst_Hologram!DocNo) - TmpRs4!CMPID * 1000000) & " !!!"
'         TmpRs4.Close
'      End If
'   Else
'      MsgBox "Hologram Already used in Location : " & LocName & " - " & RsMst_Hologram!DocType & " number :-" & RsMst_Hologram!DocNo & " !!!"
'   End If
'   retval = False
'End If
'RsMst_Hologram.Close
'Check_HologramNo = retval
'End Function

Public Function Check_HologramNo(CID_ As Variant) As Boolean

Dim Retval As Boolean
Dim LocName As String

Retval = True

If IsNumeric(CID_) = False Or Val(CID_) < 0 Then
   MsgBox "Invalid numeric format!!!"
   Retval = False
   Check_HologramNo = Retval
   Exit Function
End If

Call TableMst_Hologram(" Where NewHologramNo = " & Val(CID_) & " And HologramNo is Null")

If RsMst_Hologram.RecordCount > 0 Then

   Call TableMst_Location(" Where SM_Prefix = '" & RsMst_Hologram!SM_Prefix & "' ")
   If RsMst_Location.RecordCount > 0 Then
      LocName = RsMst_Location!Locationname
   End If
   If RsMst_Hologram!DocType = "WHR" Then
      tmp = "Select CMPID from Txn_WHR_Detail Where WHRNo = " & RsMst_Hologram!DocNo & ""
      Call Tmp4Table
      If TmpRs4.RecordCount > 0 Then
         MsgBox "Hologram Already used in Location : " & LocName & " - " & RsMst_Hologram!DocType & " " & TmpRs4!CMPID & " number :- " & (Val(RsMst_Hologram!DocNo) - TmpRs4!CMPID * 1000000) & " !!!"
         TmpRs4.Close
      End If
   Else
      MsgBox "Hologram Already used in Location : " & LocName & " - " & RsMst_Hologram!DocType & " number :- " & RsMst_Hologram!DocNo & " !!!"
   End If
   Retval = False
   RsMst_Hologram.Close
Else
   Call TableMst_Hologram(" Where NewHologramNo = " & Val(CID_) & " And HologramNo is Not Null")

   If RsMst_Hologram.RecordCount > 0 Then
      Call TableMst_Location(" Where SM_Prefix = '" & RsMst_Hologram!SM_Prefix & "' ")
      If RsMst_Location.RecordCount > 0 Then
         LocName = RsMst_Location!Locationname
      End If
      If RsMst_Hologram!DocType = "WHR" Then
         tmp = "Select CMPID from Txn_WHR_Detail Where WHRNo = " & RsMst_Hologram!DocNo & ""
         Call Tmp4Table
         If TmpRs4.RecordCount > 0 Then
            MsgBox "Hologram used in Location : " & LocName & " - " & RsMst_Hologram!DocType & " " & TmpRs4!CMPID & " number :- " & (Val(RsMst_Hologram!DocNo) - TmpRs4!CMPID * 1000000) & " !!!"
            TmpRs4.Close
            Retval = False
         End If
      Else
         MsgBox "Hologram used in Location : " & LocName & " - " & RsMst_Hologram!DocType & " number :- " & RsMst_Hologram!DocNo & " !!!"
         Retval = False
      End If
   End If
End If

Check_HologramNo = Retval

End Function


Public Function GetAssetType(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_AssetType("Where AssetsID = " & CID_)
If Not RsMst_AssetType.EOF Then
      Retval = RsMst_AssetType!AssetType
End If
GetAssetType = Retval
End Function


Public Function GetAssetStatus(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_AssetStatus("Where AssetsStatusID = " & CID_)
If Not RsMst_AssetStatus.EOF Then
      Retval = RsMst_AssetStatus!AssetStatus
End If
GetAssetStatus = Retval
End Function


Public Function GetAssetTypeID(CID_ As String) As Variant
Dim Retval As Variant
Retval = 0
Call TableMst_AssetType(" Where AssetType = '" & CID_ & "'")
If Not RsMst_AssetType.EOF Then
  Retval = RsMst_AssetType!AssetsID
End If
GetAssetTypeID = Retval
End Function


Public Function GetAssetStatusID(CID_ As String) As Variant
Dim Retval As Variant
Retval = 0
Call TableMst_AssetStatus(" Where AssetStatus = '" & CID_ & "'")
If Not RsMst_AssetStatus.EOF Then
  Retval = RsMst_AssetStatus!AssetsStatusID
End If
GetAssetStatusID = Retval
End Function


Public Function GetDesignaion(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_Designation("Where DesignationID = " & CID_)
If Not RsMst_Designation.EOF Then
      Retval = RsMst_Designation!Designation
End If
GetDesignaion = Retval
End Function

Public Function GetEmployeeName(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0
tmp = "Select * from Mst_Employee Where EmployeeID = " & CID_ & ""
'Call Mst_EmployeeTable("Where EmployeeID = " & CID_ & "")
Call Tmp3Table
If Not TmpRs3.EOF Then
      Retval = TmpRs3!EmployeeName
End If
GetEmployeeName = Retval
End Function


Public Function GetEmployeeNonName(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0
tmp = "Select * from Mst_Employee Where EmployeeID = " & CID_ & ""
'Call Mst_EmployeeTable("Where EmployeeID = " & CID_ & "")
Call Tmp6Table
If Not TmpRs6.EOF Then
      Retval = TmpRs6!EmployeeName & "~" & TmpRs6!EmployeeNo
End If
GetEmployeeNonName = Retval
End Function
Public Function GetSuperwiserName(CID_ As Variant, Typ_ As String) As Variant
Dim Retval As Variant
Retval = 0
tmp = "Select * from Mst_SecurityDetails Where SecurityDetailID = " & CID_ & ""
'Call Mst_EmployeeTable("Where EmployeeID = " & CID_ & "")
Call Tmp3Table
If Not TmpRs3.EOF Then
   If Typ_ = "N" Then
      Retval = TmpRs3!SecurityName
   Else
      Retval = TmpRs3!SecurityAgencyID
   End If
End If
GetSuperwiserName = Retval
End Function


Public Function GetSecurityAgency(CID_ As Variant, Typ_ As String) As String
Dim Retval As Variant
Retval = 0
Call TableMst_SecurityAgency("Where SecurityAgencyID = " & CID_)
If Not RsMst_SecurityAgency.EOF Then
   If Typ_ = "N" Then
       Retval = RsMst_SecurityAgency!AgencyName
   Else
       Retval = RsMst_SecurityAgency!Flag
   End If
      
End If
GetSecurityAgency = Retval
End Function


Public Function GetSecurityType(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_SecurityShifts("Where SecurityShiftsID = " & CID_)
If Not RsMst_SecurityShifts.EOF Then
      Retval = RsMst_SecurityShifts!SecurityShifts
End If
GetSecurityType = Retval
End Function


Public Function GetSecurityAssetID(CID_ As String) As Variant
Dim Retval As Variant
Retval = 0
Call TableMst_SecurityAssets(" Where SecurityAsset = '" & CID_ & "'")
If Not RsMst_SecurityAssets.EOF Then
  Retval = RsMst_SecurityAssets!SecurityAssetID
End If
GetSecurityAssetID = Retval
End Function


Public Function GetSecurityAsset(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_SecurityAssets("Where SecurityAssetID = " & CID_)
If Not RsMst_SecurityAssets.EOF Then
      Retval = RsMst_SecurityAssets!SecurityAsset
End If
GetSecurityAsset = Retval
End Function
Public Function GetExpenseType(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_ExpenseType("Where ExpenseTypeID = " & CID_)
If Not RsMst_ExpenseType.EOF Then
      Retval = RsMst_ExpenseType!ExpenseType
End If
GetExpenseType = Retval
End Function
Public Function GetDesignationFromEmp(CID_ As Variant) As String
Dim Retval As String
Retval = 0
Call TableMst_Employee(" Where EmployeeID = " & CID_)
If Not RsMst_Employee.EOF Then
  Retval = GetDesignaion(RsMst_Employee!DesignationID)
End If
GetDesignationFromEmp = Retval
End Function

Public Function GetAuditType(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_AuditType("Where AuditTypeID = " & CID_)
If Not RsMst_AuditType.EOF Then
  Retval = RsMst_AuditType!AuditType
End If
GetAuditType = Retval
End Function

Public Function GetAuditSegmentType(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_AuditSegment("Where SegmentTypeID = " & CID_)
If Not RsMst_AuditSegment.EOF Then
  Retval = RsMst_AuditSegment!SegmentType
End If
GetAuditSegmentType = Retval
End Function

Public Function GetAuditPoint(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_AuditPoint("Where AuditPointID = " & CID_)
If Not RsMst_AuditPoint.EOF Then
  Retval = RsMst_AuditPoint!AuditPoint
End If
GetAuditPoint = Retval
End Function

Public Function GetCMPFromGodown(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_Godown(" Where GodownID = " & CID_)
If Not RsMst_Godown.EOF Then
  Retval = GetCMPName(RsMst_Godown!CMPID)
End If
GetCMPFromGodown = Retval
End Function

Public Function GetEmployeeEmailID(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0
Call TableMst_Employee("Where EmployeeID = " & CID_ & "")
If Not RsMst_Employee.EOF Then
      Retval = RsMst_Employee!EmployeeEmailID
End If
GetEmployeeEmailID = Retval
End Function


Public Function GetCRIDBYPINO(CID_ As Variant) As String
Dim Retval As Variant
Retval = ""
Call TableTxn_CashReceiptInvoiceDetails("Where PI_NO = '" & CID_ & "'")
If Not RsTxn_CashReceiptInvoiceDetails.EOF Then
  Retval = RsTxn_CashReceiptInvoiceDetails!NewCRID
End If
GetCRIDBYPINO = Retval
End Function

Public Function GetRRNIDFromRRNNo(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_RRN("Where RRNNO = " & CID_)
If Not RsMst_RRN.EOF Then
  Retval = RsMst_RRN!RRNID
End If
End Function


Public Function GetBankID(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
Call TableMst_Bank(" Where BankName = '" & CID_ & "'")
If Not RsMst_Bank.EOF Then
   Retval = RsMst_Bank!BankID
End If
GetBankID = Retval
End Function
Public Function GetLocationID(CID_ As Variant) As Double
Dim Retval As Variant
Retval = 0
Call TableMst_Location(" Where LocationName = '" & CID_ & "'")
If Not RsMst_Location.EOF Then
   Retval = RsMst_Location!LocationID
End If
GetLocationID = Retval
End Function



Public Function GetExpenseGroupName(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_ExpenseGroup(" Where ExpenseGroupID = " & CID_)
If Not RsMst_ExpenseGroup.EOF Then
   Retval = RsMst_ExpenseGroup!ExpenseGroup
End If
GetExpenseGroupName = Retval
End Function
Public Function GetTDSDescription(CID_ As Variant) As String
Dim Retval As String
Retval = 0
Call TableMst_TDS(" Where TDSID = " & CID_)
If Not RsMst_TDS.EOF Then
   Retval = RsMst_TDS!Description
End If
GetTDSDescription = Retval
End Function


Public Function GetServiceTaxDescription(CID_ As Variant) As String
Dim Retval As String
Retval = 0
Call TableMst_ServiceTax(" Where ServiceTaxID = " & CID_)
If Not RsMst_ServiceTax.EOF Then
   Retval = RsMst_ServiceTax!Description
End If
GetServiceTaxDescription = Retval
End Function
Public Function GetVendorName(CID_ As Variant, Typ_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_Vendor(" Where VendorID = " & CID_)

If Not RsMst_Vendor.EOF Then
   If Typ_ = "N" Then
      Retval = RsMst_Vendor!VendorName
   ElseIf Typ_ = "A" Then
      Retval = RsMst_Vendor!VendorAddress
   Else
      Retval = RsMst_Vendor!EmailID
   End If
End If
GetVendorName = Retval
End Function


Public Function GetClientIDRowByNameAndExchange(CID_ As Variant, Exg_ As Double) As Variant
Dim Retval As String
Retval = 0
Call TableMst_Client(" Where ClientName = '" & CID_ & "' And ExchangeTypeID = " & Exg_ & " ")
If Not RsMst_Client.EOF Then
   Retval = RsMst_Client!ClientIDRow
End If
GetClientIDRowByNameAndExchange = Retval
End Function

Public Function GetRRNNo(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_RRN("Where RRNID = " & CID_)
If Not RsMst_RRN.EOF Then
   Retval = RsMst_RRN!RRNNo
End If
GetRRNNo = Retval
End Function


Public Function GetEmployeeIDFromEmailid(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0
tmp = "Select * from Mst_Employee Where EmployeeEmailID = '" & CID_ & "' And Flag = 'Y'"
'Call Mst_EmployeeTable("Where EmployeeID = " & CID_ & "")
Call Tmp3Table
If Not TmpRs3.EOF Then
      Retval = TmpRs3!EmployeeID
End If
GetEmployeeIDFromEmailid = Retval

End Function
Public Function GetDesignationFromEmailid(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0
tmp = "Select * from Mst_Employee Where EmployeeEmailID = '" & CID_ & "' And Flag = 'Y'"
'Call Mst_EmployeeTable("Where EmployeeID = " & CID_ & "")
Call Tmp3Table
If Not TmpRs3.EOF Then
      Retval = TmpRs3!DesignationID
End If
GetDesignationFromEmailid = Retval

End Function

Public Function GetLocationFromCMPName(CID_ As Variant) As String
Dim Retval As String
Retval = 0
Call TableMst_CMP(" Where CMPName = '" & CID_ & "' ")
If Not RsMst_CMP.EOF Then
   Retval = GetLocationName(RsMst_CMP!LocationID)
End If
GetLocationFromCMPName = Retval
End Function

Public Function GetCreatedByName(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0
Call TableMst_Users(" Where EmployeeNo = '" & CID_ & "' ")
If Not RsMst_Users.EOF Then
      Retval = RsMst_Users!EmployeeName
End If
GetCreatedByName = Retval

End Function

Public Function GetLessorName(CID_ As Variant) As String
Dim Retval As String
Retval = 0
Call TableMst_Lessor(" Where LessorID = " & CID_ & " ")
If Not RsMst_Lessor.EOF Then
   Retval = RsMst_Lessor!LessorName
End If
GetLessorName = Retval
End Function

Public Function GetExpenseTypeID(CID_ As Variant) As Double
Dim Retval As Variant
Retval = 0
Call TableMst_ExpenseType(" Where ExpenseType = '" & CID_ & "'")
If Not RsMst_ExpenseType.EOF Then
   Retval = RsMst_ExpenseType!ExpenseTypeID
End If
GetExpenseTypeID = Retval
End Function


Public Function GetGLAccNAVCode(GLType_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_GLAccType("Where GLAccType = '" & GLType_ & "'")
If Not RsMst_GLAccType.EOF Then
  Retval = RsMst_GLAccType!NavCode
End If
GetGLAccNAVCode = Retval
End Function


Public Function GetDepositoryNameFromNavCode(NavCode_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_Depository(" Where NavCode='" & NavCode_ & "'")
If Not RsMst_Depository.EOF Then
   Retval = RsMst_Depository!DepositoryName
End If
GetDepositoryNameFromNavCode = Retval
End Function

Public Function GetBillingCycle(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_BillingCycle("Where BillingCycleID = " & CID_)
If Not RsMst_BillingCycle.EOF Then
   If Mod_ = "D" Then
      Retval = RsMst_BillingCycle!NoOfDays
   Else
      Retval = RsMst_BillingCycle!BillingCycle
   End If
End If
GetBillingCycle = Retval
End Function
Public Function GetServiceTaxNAVCode(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_ServiceTax(" Where ServiceTaxID = " & CID_)
If Not RsMst_ServiceTax.EOF Then
   Retval = IIf(IsNull(RsMst_ServiceTax!NavCode), "", RsMst_ServiceTax!NavCode)
End If
GetServiceTaxNAVCode = Retval
End Function
Public Function GetFAGBankName(CID_ As Double, Mod_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_FAGBank("Where FAGBankID = " & CID_)
If Not RsMst_FAGBank.EOF Then
   If Mod_ = "N" Then
      Retval = RsMst_FAGBank!BankName
   ElseIf Mod_ = "AC" Then
      Retval = RsMst_FAGBank!AccountNo
   Else
      Retval = RsMst_FAGBank!Address
   End If
End If
GetFAGBankName = Retval
End Function

Public Function GetITVendorName(CID_ As Variant, Typ_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_ITVendor(" Where VendorID = " & CID_)

If Not RsMst_ITVendor.EOF Then
   If Typ_ = "N" Then
      Retval = RsMst_ITVendor!VendorName
   Else
      Retval = RsMst_ITVendor!EmailID
   End If
End If
GetITVendorName = Retval
End Function

Public Function GetClientGroup(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_ClientGroup(" Where ClientGroupID = " & CID_)
If Not RsMst_ClientGroup.EOF Then
  Retval = RsMst_ClientGroup!ClientGroupName
End If
GetClientGroup = Retval
End Function
Public Function GetLessorAddress(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_Lessor(" Where LessorID = " & CID_ & " ")
If Not RsMst_Lessor.EOF Then
   Retval = RsMst_Lessor!Address
End If
GetLessorAddress = Retval
End Function

Public Function GetBillingCycleDays(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0
Call TableMst_BillingCycle(" Where BillingCycleID = " & CID_ & " ")
If Not RsMst_BillingCycle.EOF Then
   Retval = RsMst_BillingCycle!NoOfDays
End If
GetBillingCycleDays = Retval
End Function

Public Function GetReportingDesignation(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_Designation("Where DesignationID = " & CID_)
If Not RsMst_Designation.EOF Then
      Retval = IIf(IsNull(RsMst_Designation!ReportingDesignationID), 1, RsMst_Designation!ReportingDesignationID)
End If
GetReportingDesignation = Retval
End Function

Public Function GetCMPIDFromGodown(CID_ As Double) As String
Dim Retval As String
Retval = 0
Call TableMst_Godown(" Where GodownID = " & CID_)
If Not RsMst_Godown.EOF Then
   Retval = RsMst_Godown!CMPID
End If
GetCMPIDFromGodown = Retval
End Function

Public Function GetQualityParameter(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_QualityParameter(" Where QualityParameterID = " & CID_ & "")
If Not RsMst_QualityParameter.EOF Then
   Retval = RsMst_QualityParameter!QualityParameter
End If
GetQualityParameter = Retval
End Function

Public Function GetTestingAgencyName(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_TestingAgency(" Where TestingAgencyID = " & CID_ & "")
If Not RsMst_TestingAgency.EOF Then
   Retval = RsMst_TestingAgency!AgencyName
End If
GetTestingAgencyName = Retval
End Function

Public Function CheckHologramInRange(mHologramNo_ As Variant, mLocationID_ As Variant) As Boolean
Dim mflg As Boolean
mflg = False
Call TableTxn_Hologram("Where LocationID=" & mLocationID_)

If RsTxn_Hologram.RecordCount > 0 Then
   For I = 0 To RsTxn_Hologram.RecordCount - 1
       If Val(mHologramNo_) >= RsTxn_Hologram!FromHologramNo And Val(mHologramNo_) <= RsTxn_Hologram!ToHologramNo Then
          mflg = True
          Exit For
       End If
       RsTxn_Hologram.MoveNext
   Next
End If
CheckHologramInRange = mflg

End Function

Public Function GetSecurityDesignation(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_SecurityDesignation(" Where SecurityDesignationID = " & CID_ & "")
If Not RsMst_SecurityDesignation.EOF Then
   Retval = RsMst_SecurityDesignation!SecurityDesignation
End If
GetSecurityDesignation = Retval
End Function


Public Function GetSecReportingDesignation(CID_ As Variant) As String
Dim Retval As Variant
Retval = 0
Call TableMst_Designation("Where DesignationID = " & CID_)
If Not RsMst_Designation.EOF Then
   Retval = IIf(IsNull(RsMst_Designation!SecReportingDesignationID), 0, RsMst_Designation!SecReportingDesignationID)
End If
GetSecReportingDesignation = Retval
End Function

Public Function GetServiceTaxDescFromNAVCode(CID_ As Variant) As String
Dim Retval As String
Retval = 0
Call TableMst_ServiceTax(" Where NAVCode = '" & CID_ & "'")
If Not RsMst_ServiceTax.EOF Then
   Retval = RsMst_ServiceTax!Description
End If
GetServiceTaxDescFromNAVCode = Retval
End Function

Public Function GetDepositoryIDFromNavCode(NavCode_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_Depository(" Where NavCode='" & NavCode_ & "'")
If Not RsMst_Depository.EOF Then
   Retval = RsMst_Depository!DepositoryID
End If
GetDepositoryIDFromNavCode = Retval
End Function

Public Function GetStateIndiaName(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0

Call TableMst_StateIndia(" Where StateIndiaID = " & CID_ & "")
If Not RsMst_StateIndia.EOF Then
   Retval = RsMst_StateIndia!StateIndiaName
End If
GetStateIndiaName = Retval

End Function

Public Function GetRRNID(CID_ As String) As String
Dim Retval As String
Retval = ""
Call TableMst_RRN("Where RRNNo = '" & CID_ & "'")
If Not RsMst_RRN.EOF Then
   Retval = RsMst_RRN!RRNID
End If
GetRRNID = Retval
End Function

'Public Function GetClientsRemInvoices(Optional ByVal wc As String = "", Optional ByVal CRG_UID As String = "", Optional ByVal AdjG_UID As String = "", Optional ByVal AdvAllocID_ As Double = 0, Optional ByVal Flag_ As String = "")
'
''---Txn_NewCashReceiptTIDetail
''---Txn_CRTDSAdjustmentDetail
'
'tmp = "Select TINo,TIDate,TIAmount,TIID,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode,TISTAmount ,TI.AdjAmountFAG,"
'tmp = tmp & " (Select Sum(TCRT.RecAmount) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID "
'If CRG_UID <> "" Then
'   tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
'End If
'If AdvAllocID_ > 0 Then
'   tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
'End If
'
'tmp = tmp & " )RecAmount,"
'tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID"
'If CRG_UID <> "" Then
'   tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
'End If
'If AdvAllocID_ > 0 Then
'   tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
'End If
'tmp = tmp & " ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID"
'If AdjG_UID <> "" Then
'   tmp = tmp & " And TCR.G_UID<>'" & AdjG_UID & "'"
'End If
'
'tmp = tmp & " )RecTDS,"
'tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID "
'If CRG_UID <> "" Then
'   tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
'End If
'If AdvAllocID_ > 0 Then
'   tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
'End If
'tmp = tmp & " )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID "
'If AdjG_UID <> "" Then
'   tmp = tmp & " And TCR.G_UID<>'" & AdjG_UID & "'"
'End If
'
'tmp = tmp & " )RecOtherDed"
'tmp = tmp & " From  Txn_TIDetails TI"
'tmp = tmp & " Where  TINo<>''"
'If LCase(Flag_) <> "show" Then
'   tmp = tmp & " And TIAmount-(AdjAmountFAG) - ("
'   tmp = tmp & " (select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y'"
'   If CRG_UID <> "" Then
'      tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
'   End If
'   If AdvAllocID_ > 0 Then
'      tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
'   End If
'   tmp = tmp & " )+"
'   tmp = tmp & " (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R'"
'
'   If AdjG_UID <> "" Then
'      tmp = tmp & " And TCR.G_UID<>'" & AdjG_UID & "'"
'   End If
'   tmp = tmp & " ))>0"
'End If
'If wc <> "" Then
'   tmp = tmp & " " & wc
'End If
'
'tmp = tmp & " Order By TI.TIDate,TI.NAVCode"
'
'Call TmpTable
'
'End Function

Public Function GetClientsRemInvoices(Optional ByVal wc As String = "", Optional ByVal CRG_UID As String = "", Optional ByVal AdjG_UID As String = "", Optional ByVal AdvAllocID_ As Double = 0, Optional ByVal Flag_ As String = "")

'---Txn_NewCashReceiptTIDetail
'---Txn_CRTDSAdjustmentDetail

tmp = "Select TINo,TIDate,TIAmount,TIID,TI.PostInvoiceNo,TI.EntryNo,ML.LocationName,TI.NAVCode,TISTAmount ,TI.AdjAmountFAG,"
tmp = tmp & " (Select Sum(TCRT.RecAmount) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID "
If CRG_UID <> "" Then
   tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
End If
If AdvAllocID_ > 0 Then
   tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
End If

tmp = tmp & " )RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID"
If CRG_UID <> "" Then
   tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
End If
If AdvAllocID_ > 0 Then
   tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
End If
tmp = tmp & " ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID"
If AdjG_UID <> "" Then
   tmp = tmp & " And TCR.G_UID<>'" & AdjG_UID & "'"
End If

tmp = tmp & " )RecTDS,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID "
If CRG_UID <> "" Then
   tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
End If
If AdvAllocID_ > 0 Then
   tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
End If
tmp = tmp & " )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID "
If AdjG_UID <> "" Then
   tmp = tmp & " And TCR.G_UID<>'" & AdjG_UID & "'"
End If

tmp = tmp & " )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join  MSt_Location ML on ML.LocationID=TI.LocationID"
tmp = tmp & " Where  TINo<>''"
If LCase(Flag_) <> "show" Then
   tmp = tmp & " And TIAmount-(AdjAmountFAG) - ("
   tmp = tmp & " (select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y'"
   If CRG_UID <> "" Then
      tmp = tmp & " And TCR.G_UID<>'" & CRG_UID & "'"
   End If
   If AdvAllocID_ > 0 Then
      tmp = tmp & " And isnull(TCRT.CRAAID,0) <> " & AdvAllocID_ & ""
   End If
   tmp = tmp & " )+"
   tmp = tmp & " (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R'"

   If AdjG_UID <> "" Then
      tmp = tmp & " And TCR.G_UID<>'" & AdjG_UID & "'"
   End If
   tmp = tmp & " ))>0"
End If
If wc <> "" Then
   tmp = tmp & " " & wc
End If

tmp = tmp & " Order By TI.TIDate,TI.NAVCode"

Call TmpTable

End Function

Public Function GetStateID(CID_ As Variant) As Double
Dim Retval As Variant
Retval = 0
Call TableMst_State(" Where StateName = '" & CID_ & "'")
If Not RsMst_State.EOF Then
   Retval = RsMst_State!StateID
End If
GetStateID = Retval
End Function

Public Function CheckTIInCR(CID_ As Variant) As Boolean
Dim Retval As Variant
Retval = False
tmp = " Select TIID from Txn_NewCashReceiptTIDetail  TCRT"
tmp = tmp & " inner join Txn_NewCashReceipt TCR on TCR.NewCRID=TCRT.NewCRID And TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And TIID=" & CID_
tmp = tmp & " Union"
tmp = tmp & " Select TIID from Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " inner join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID=TCRT.CRTDSAdjustmentID And TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R' And TIID=" & CID_

Call Tmp6Table
If TmpRs6.RecordCount > 0 Then
   Retval = True
End If
CheckTIInCR = Retval
End Function

Public Function GetTIDetails(PINo_ As Double, Flag_ As String) As String()
Dim Retval(11) As String
Retval(0) = "" 'TIID
Retval(1) = "" 'TI No
Retval(2) = "" 'TIDate
Retval(3) = "" 'TI Amount
Retval(4) = "" 'Post Inv No
Retval(5) = "" 'Entry No
Retval(6) = "" 'NAV Code
Retval(7) = "" 'TI ST Amt
Retval(8) = "" 'Rec Amt
Retval(9) = "" 'Rec TDS
Retval(10) = "" 'Rec Other Deduction

tmp = "Select TI.TIID,TI.TINo,TI.TIDate,TI.TIAmount,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode,TI.TISTAmount ,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+isnull(TI.AdjAmountFAG,0) RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID ) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDS,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID  )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Where TI.PIID=" & PINo_ & " and TI.PIFLag='" & Flag_ & "'  "

Call Tmp6Table

If Not TmpRs6.EOF Then
  Retval(0) = TmpRs6!TIID
  Retval(1) = TmpRs6!TINo
  Retval(2) = TmpRs6!TIDate
  Retval(3) = TmpRs6!TIAmount
  Retval(4) = TmpRs6!PostInvoiceNo
  Retval(5) = IIf(IsNull(TmpRs6!EntryNo), 0, TmpRs6!EntryNo)
  Retval(6) = IIf(IsNull(TmpRs6!NavCode), "", TmpRs6!NavCode)
  Retval(7) = IIf(IsNull(TmpRs6!TISTAmount), 0, TmpRs6!TISTAmount)
  Retval(8) = IIf(IsNull(TmpRs6!RecAmount), 0, TmpRs6!RecAmount)
  Retval(9) = IIf(IsNull(TmpRs6!RecTDS), 0, TmpRs6!RecTDS)
  Retval(10) = IIf(IsNull(TmpRs6!RecOtherDed), 0, TmpRs6!RecOtherDed)
End If
TmpRs6.Close
GetTIDetails = Retval
End Function


Public Function GetNAVDepositoryName(CID_ As String) As String
Dim Retval As String
Retval = 0
Call TableMst_Depository(" Where DepositoryID = '" & CID_ & "'")
If Not RsMst_Depository.EOF Then
   Retval = IIf(IsNull(RsMst_Depository!NAVDepositoryName), "", RsMst_Depository!NAVDepositoryName)
End If
GetNAVDepositoryName = Retval
End Function

Public Function GetEarthequakeZone(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_EarthquakeZone("Where EQZoneID = " & CID_)
If Not RsMst_EarthquakeZone.EOF Then
  Retval = RsMst_EarthquakeZone!EQZone
End If
GetEarthequakeZone = Retval
End Function

Public Function GetHRDesignationFromHRPro(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""

tmp = "Select MD.Designation from Mst_Employee ME Inner Join Mst_Designation MD on ME.DesignationID = MD.DesignationID Where ME.EmployeeNo = '" & CID_ & "'"
Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = TmpRsHRPro!Designation
End If

GetHRDesignationFromHRPro = Retval

End Function

Public Function GetHRLocation(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
tmp = "Select * from Mst_Location Where LocationID= " & CID_ & ""

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = TmpRsHRPro!Locationname
End If
GetHRLocation = Retval

End Function

Public Function GetHRDirectReporteeFromHRPro(CID_ As Variant, mFlag_ As String) As Variant
Dim Retval As String
Retval = ""

tmp = "Select ME1.FullName + '~' + ME1.EmployeeNo as  EmployeeName, ME.LocationID from Mst_Employee ME Inner Join Mst_Employee ME1 on ME.ReportingToID = ME1.EmployeeID Where ME.EmployeeNo = '" & CID_ & "'"

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   If mFlag_ = "D" Then
      Retval = TmpRsHRPro!EmployeeName
   ElseIf mFlag_ = "L" Then
      Retval = TmpRsHRPro!LocationID
   End If
End If

GetHRDirectReporteeFromHRPro = Retval

End Function


Public Function GetCommodityCommtrackFlag(CID_ As Double) As Variant
Dim Retval As Variant
Retval = 0
Call TableMst_Commodity("Where CommodityID = " & CID_)
If Not RsMst_Commodity.EOF Then
   Retval = IIf(IsNull(RsMst_Commodity!CommTrackFlag), 0, RsMst_Commodity!CommTrackFlag)
End If
GetCommodityCommtrackFlag = Retval
End Function

Public Function GetStateIndiaFromStateName(CID_ As Variant) As Variant

Dim Retval As Variant
Retval = 0

Call TableMst_State("Where StateName = '" & CID_ & "'")

If Not RsMst_State.EOF Then
   Retval = GetStateIndiaName(RsMst_State!StateIndiaID)
End If

GetStateIndiaFromStateName = Retval

End Function

Public Function GetBranchID(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
Call TableMst_Branch(" Where BranchName = '" & CID_ & "'")
If Not RsMst_Branch.EOF Then
   Retval = RsMst_Branch!BranchID
End If
GetBranchID = Retval
End Function

Public Function GetVehicleType(CID_ As Variant, Mod_ As String) As String
Dim Retval As String
Retval = ""
If IsNull(CID_) = False Then
   Call TableMst_VehicleType("Where VehicleTypeID = " & CID_)
   If Not RsMst_VehicleType.EOF Then
      If LCase(Mod_) = "n" Then
         Retval = RsMst_VehicleType!VehicleType
      ElseIf LCase(Mod_) = "min" Then
         Retval = RsMst_VehicleType!MinTareWt
      ElseIf LCase(Mod_) = "max" Then
         Retval = RsMst_VehicleType!MaxTareWt
      End If
   End If
End If
GetVehicleType = Retval
End Function

Public Function GetCRFFirmName(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
tmp = "Select * from Txn_CRF Where CRFID =  " & CID_ & ""

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = TmpRsHRPro!FirmName
End If
GetCRFFirmName = Retval

End Function

Public Function GetClientIDRowFromClientID(CID_ As Variant) As Double
Dim Retval As Double
Retval = 0
Call TableMst_Client(" Where ClientIDRow = '" & CID_ & "' And Status='A'")
If Not RsMst_Client.EOF Then
  Retval = RsMst_Client!ClientID
End If
GetClientIDRowFromClientID = Retval
End Function

Public Function GetHRCommodity(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
tmp = "Select * from Mst_Commodity Where CommodityID= " & CID_ & ""

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = TmpRsHRPro!CommodityName
End If
GetHRCommodity = Retval

End Function

Public Function GetHRVariety(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
tmp = "Select * from Mst_Variety Where VarietyID= " & CID_ & ""

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = TmpRsHRPro!VarietyName
End If
GetHRVariety = Retval

End Function

Public Function GetHRBank(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
tmp = "Select * from Mst_Bank Where BankID= " & CID_ & ""

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = TmpRsHRPro!BankName
End If
GetHRBank = Retval

End Function


Public Function GetHRBranch(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
tmp = "Select * from Mst_Branch Where BranchID= " & CID_ & ""

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = TmpRsHRPro!BranchName
End If
GetHRBranch = Retval

End Function

Public Function GetJournalDetailsFromHRPro(EmpNo_ As Variant, JournalType_ As Variant)
Dim Retval As String
Retval = ""
tmp = " SELECT MBCM.MappingID,MBCM.Flag,MBCM.JournalCode,ME.FullName+'~'+ME.EmployeeNo 'Employee',MJ.JournalTypeName  FROM Mst_EmpJournalCodeMapping MBCM Inner Join Mst_Employee ME On MBCM.EmployeeID = ME.EmployeeID  Inner Join Mst_JournalType MJ On MBCM.JournalTypeID = MJ.JournalTypeID  Where MBCM.Flag = 'A' And ME.EmployeeNo = '" & Trim(EmpNo_) & "' And MJ.JournalTypeName = '" & Trim(JournalType_) & "'"
Call TmpTableHRPro
If Not TmpRsHRPro.EOF Then
   Retval = IIf(IsNull(TmpRsHRPro!JournalCode), "", TmpRsHRPro!JournalCode)
End If
GetJournalDetailsFromHRPro = Retval
End Function

Public Function GetFAGBankAcNo(mFAGBankID_ As Variant) As Variant
Call TableMst_FAGBank(" Where FAGBankId = " & Val(mFAGBankID_) & " ")
If RsMst_FAGBank.RecordCount > 0 Then
   'GetFAGBankAcNo = IIf(IsNull(RsMst_FAGBank!NAVCode), 0, RsMst_FAGBank!NAVCode)
   GetFAGBankAcNo = IIf(IsNull(RsMst_FAGBank!BankName), "", RsMst_FAGBank!BankName) & "~" & IIf(IsNull(RsMst_FAGBank!AccountNo), "", RsMst_FAGBank!AccountNo)
Else
    GetFAGBankAcNo = ""
End If
End Function

Public Function GetHRBranchIFSC(CID_ As Variant) As Variant
Dim Retval As String
Retval = ""
tmp = "Select * from Mst_Branch Where BranchID= " & CID_ & ""

Call TmpTableHRPro

If Not TmpRsHRPro.EOF Then
   Retval = IIf(IsNull(TmpRsHRPro!IFSCCode), "", TmpRsHRPro!IFSCCode)
End If
GetHRBranchIFSC = Retval

End Function

Public Function GetGLAccNAVCodeFromInvoiceType(InvID As Variant, AccType As Variant, Optional Flag_ As Variant) As Variant
Dim Retval As String
Retval = ""

Call TableMst_GLAccType("Where AccountType = '" & AccType & "' And GLAccTypeID in (Select Distinct GLAccTypeID From Mst_GLAccountIncomeTypeMapping Where IETypeID = " & InvID & ") ")

If Not RsMst_GLAccType.EOF Then
   If LCase(Flag_) = "h" Then
      Retval = RsMst_GLAccType!HSNNo
   Else
      Retval = RsMst_GLAccType!NavCode
   End If
End If

GetGLAccNAVCodeFromInvoiceType = Retval

End Function

Public Function GetGLAccNAVCodeFromExpenseType(ExpenseTypeID As Variant, AccType As Variant, Optional Flag_ As Variant) As Variant
Dim Retval As String
Retval = ""
Call TableMst_GLAccType("Where AccountType = '" & AccType & "' And GroupType = 'E' And GLAccTypeID in (Select Distinct GLAccTypeID From Mst_GLAccountIncomeTypeMapping Where ExpenseTypeID = " & ExpenseTypeID & ") ")
If Not RsMst_GLAccType.EOF Then
   If LCase(Flag_) = "h" Then
      Retval = RsMst_GLAccType!HSNNo
   Else
      Retval = RsMst_GLAccType!NavCode
   End If
End If
GetGLAccNAVCodeFromExpenseType = Retval
End Function
Public Function GetServiceTaxID(CID_ As Variant) As String
Dim Retval As String
Retval = ""
Call TableMst_ServiceTax(" Where Description = '" & CID_ & "'")
If Not RsMst_ServiceTax.EOF Then
   Retval = IIf(IsNull(RsMst_ServiceTax!ServiceTaxID), "", RsMst_ServiceTax!ServiceTaxID)
End If
GetServiceTaxID = Retval
End Function

Public Function GetRegion(CID_ As Variant) As Variant
Dim Retval As Variant
Retval = 0

Call TableMst_Region(" Where RegionID = " & CID_ & "")
If Not RsMst_Region.EOF Then
   Retval = RsMst_Region!Region
End If
GetRegion = Retval

End Function

Public Function GetInsuranceRiskCovered(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_InsuranceRiskCovered(" Where InsuranceRiskCoveredID = " & CID_)
If Not RsMst_InsuranceRiskCovered.EOF Then
   Retval = IIf(IsNull(RsMst_InsuranceRiskCovered!InsuranceRiskCovered), "", RsMst_InsuranceRiskCovered!InsuranceRiskCovered)
End If
GetInsuranceRiskCovered = Retval
End Function

Public Function GetInsuranceBy(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_InsuranceBy(" Where InsuranceByID = " & CID_)
If Not RsMst_InsuranceBy.EOF Then
   Retval = IIf(IsNull(RsMst_InsuranceBy!InsuranceByName), "", RsMst_InsuranceBy!InsuranceByName)
End If
GetInsuranceBy = Retval
End Function



Public Function GetCLientGSTDetails(CID_ As Variant, CMPID_ As Variant, Flg_ As Variant) As Variant

Dim Retval As String
Retval = ""
tmp = "Select Distinct a.ClientStateOfficeAddress,a.GSTNavCode,a.GSTNo,a.RGSTNavCode from Mst_ClientStateWiseGST a"
tmp = tmp & " Inner Join Mst_StateIndia si On a.StateIndiaID = si.StateIndiaID"
tmp = tmp & " Inner Join Mst_State s On si.StateIndiaID = s.StateIndiaID"
tmp = tmp & " Inner Join Mst_Location l On s.StateID = l.StateID"
tmp = tmp & " Inner Join Mst_CMP C On l.LocationID = c.LocationID"
tmp = tmp & " Where c.CMPID = " & Val(CMPID_) & "  and a.ClientID = " & Val(CID_) & ""

Call Tmp7Table

If Not TmpRs7.EOF Then
   If LCase(Flg_) = "a" Then
      Retval = IIf(IsNull(TmpRs7!ClientStateOfficeAddress), "", TmpRs7!ClientStateOfficeAddress)
   ElseIf LCase(Flg_) = "g" Then
      Retval = IIf(IsNull(TmpRs7!GSTNo), "", TmpRs7!GSTNo)
   ElseIf LCase(Flg_) = "gnav" Then
      Retval = IIf(IsNull(TmpRs7!GSTNavCode), "", TmpRs7!GSTNavCode)
   ElseIf LCase(Flg_) = "rgnav" Then
      Retval = IIf(IsNull(TmpRs7!RGSTNavCode), "", TmpRs7!RGSTNavCode)
   Else
      Retval = TmpRs7!ClientStateOfficeAddress
   End If
End If
GetCLientGSTDetails = Retval


End Function
Public Function GetStateWiseNAVCode(ClID_ As Variant, mLID_ As Variant, Typ_ As Variant) As String
Dim Retval As String

Retval = ""

tmp = "Select Distinct a.ClientStateWiseGSTID,a.GSTNAVCode,a.RGSTNAVcode,a.* from Mst_ClientStateWiseGST a"
tmp = tmp & " Inner Join Mst_StateIndia MSI on a.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Inner Join Mst_State MS on MSI.StateIndiaID = MS.StateIndiaID"
tmp = tmp & " Inner Join Mst_Location ML on MS.StateID = ML.STateID"
tmp = tmp & " Where a.ClientID = " & ClID_ & " And ML.LocationID = " & mLID_ & ""

Call Tmp7Table

If Not TmpRs7.EOF Then
   If Typ_ = "N" Then
      Retval = IIf(IsNull(TmpRs7!GSTNavCode), "", TmpRs7!GSTNavCode)
   ElseIf Typ_ = "R" Then
      Retval = IIf(IsNull(TmpRs7!RGSTNavCode), "", TmpRs7!RGSTNavCode)
   End If
End If
GetStateWiseNAVCode = Retval

End Function

Public Function GetClientfromContract(CID_ As Double) As String
Dim Retval As String
Retval = ""
Call TableMst_StorageContract(" Where ContractID = " & CID_)
If Not RsMst_StorageContract.EOF Then
  Retval = RsMst_StorageContract!ClientID
End If
GetClientfromContract = Retval
End Function





