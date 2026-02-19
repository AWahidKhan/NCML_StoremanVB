Attribute VB_Name = "Module1"
Option Explicit
Public cn As ADODB.Connection
Public MGroup As String
Public cns As String
Public NToS As String
Public ATpmProd(10)
Public MTcode As String
Public MBillN As Variant 'Double
Public MbillD As Variant
Public tmp As Variant
Public tmp1 As Variant
Public NToSFLG As String
Public FromDat As Variant
Public ToDat As Variant
Public AuditedDate As Variant
Public From_GenDate As Variant
Public To_GenDate As Variant
Public Gen_RDC_RiceMilled As String
Public Gen_OryLoc_ID As String
Public Gen_Location As String
Public Gen_UserID As String
Public Gen_UserName As String
Public Gen_ImpExpFlag As String
Public Gen_FreezOpnDate As Variant
Public Gen_VersionEXE As Variant
Public Gen_DbPath As Variant
Public Gen_UserTypeID As Variant
Public Gen_UserLoginID As Variant
Public Gen_LastLogin As Variant
Public Gen_GSLID(7) As Variant
Public Gen_DBID As Variant
Public Gen_DatFormat As Variant
Public Gen_DBType As Variant
Public Gen_ExcelPassword As Variant
Public Gen_ExpireGSLFlag As Variant
Public Gen_UserRole As Variant
Public Gen_PassExpire As Variant
Public Gen_ClaimApplicant As Variant


Public Gen_UserEmailID As String
Public servertxt As String
Public oSMTP As SMTPSession

Public Gen_FLName As Variant
Public Gen_mTimeStamp As Variant
Public Gen_FormName As String
Public Gen_FieldName As String
Public Gen_Percentage As Variant

' Globle Veriable for User-level location mapping  27-jan-2010
Public Gen_WcCMP As String
Public Gen_WcLocation As String
Public Gen_WcSMLocation As String
Public Gen_WcState As String
Public Gen_WcEmployeeID As String
Public Gen_DesignationID As Variant
Public Gen_EmployeeID As Variant

Public Gen_DefaultPrinter As Boolean

Public cnHRPro As ADODB.Connection

Public cnCMPro As ADODB.Connection

Public Gen_ShowGSLMapping As Boolean

Public Gen_WHRDetail(5) As Variant

' ----------- Guid Code
'lincoln pharma
'plantinum corp.
'Ankur  : 95252-250101
'9890536330
'www.rmpinfotec.biz
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public Const S_OK = 0 ' return value from CoCreateGuid

Public Gen_FormAccess As Boolean

Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Public Gen_EmailConfigFlag As Variant

Public Gen_SecLocation, Gen_SecLocationID As Variant

Public cnSM_Attachment As ADODB.Connection

Public Gen_SysTime As Variant
Public Gen_BlockingDate As Variant

Public Gen_GodownInspection(3) As Variant

Public Gen_EmailGodownViability As Variant
Public Gen_BlockingSecurityDay As Variant
Public Gen_BlockGodRentProcessDay As Variant
Public Gen_BlockingSecurityBilling As Variant
Public Gen_BlockingFumigationBilling As Variant

Public Gen_BlockValueNNTransfer As Variant
Public Gen_BlockValueNCDEXTransfer As Variant

Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Public Gen_FAGBankIDForPayment As Variant

Public Gen_AzureDbPath As Variant

Public cnAzureStoreMan As ADODB.Connection

Public Gen_DbPathFromAzure As Variant

' ----- End Guid code
Public Function CompactDBC()
cn.Close
Set cn = Nothing
'cn.Properties.Refresh
'AuditFrezRs.Close
'ompactDatabase(

'UsersRs.Close
TmpRs.Close
Dim Pth As Variant
'strTempDB
Pth = App.Path & "\"
'strTempDB = "Procurment.mdb"
 
Dim JRO As JRO.JetEngine
Set JRO = New JRO.JetEngine

JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;" _
& "Data Source=" & Pth & "\CMProDb.mdb", _
"Provider=Microsoft.Jet.OLEDB.4.0;" _
& "Data Source=" & Pth & "\CMProDbNew.mdb" _
& ";Jet OLEDB:Engine Type=4"

JRO.CompactDatabase _
"Provider=Microsoft.Jet.OLEDB.4.0;" _
& "Data Source=C:\myData1.mdb", _
"Provider=Microsoft.Jet.OLEDB.4.0;" _
& "Data Source=C:\myDataNewName.mdb" _
& ";Jet OLEDB:Engine Type=4"
End Function

Public Function DBC()
Dim Pth As Variant
'28423229 Sankar
'9821146495

Pth = App.Path
Set cn = New ADODB.Connection
'cns = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=sa;Initial Catalog=CM;Data Source=(local)"

'cns = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=CM;Data Source=(local)"

If Gen_DbPath = "" Then
   'Gen_DbPath = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=Abcd123$;Initial Catalog=StoreManVB;Data Source=10.20.20.15" 'GetPath()
   'Gen_DbPath = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=$QL@#$%67890;Initial Catalog=Storemandb;Data Source=10.23.12.49,15333" 'GetPath()
   Gen_DbPath = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=Cy37e$84t#ji6r;Initial Catalog=SM_20251210;Data Source=10.23.12.51,1433"
   'Gen_DbPath = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=storemanuatuser;pwd=StoreManUATDB@123$%;Initial Catalog=StoreMan_UAT_GSLInvoice;Data Source=ncmluatdbserver.database.windows.net"
   
End If

If LCase(Gen_DbPathFromAzure) = "y" Then
   'Gen_DbPath = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=$QL@#$%67890;Initial Catalog=Storemandb;Data Source=10.23.12.49,15333"
   Gen_DbPath = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=Cy37e$84t#ji6r;Initial Catalog=SM_20251210;Data Source=10.23.12.51,1433"
   'GetPath()
   'Gen_DbPath = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=storemanuatuser;pwd=StoreManUATDB@123$%;Initial Catalog=StoreMan_UAT_GSLInvoice;Data Source=ncmluatdbserver.database.windows.net"
Else
   'Gen_DbPath = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=$QL@#$%67890;Initial Catalog=Storemandb;Data Source=10.23.12.49,15333" 'GetPath()
   Gen_DbPath = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=Cy37e$84t#ji6r;Initial Catalog=SM_20251210;Data Source=10.23.12.51,1433"
   'Gen_DbPath = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=storemanuatuser;pwd=StoreManUATDB@123$%;Initial Catalog=StoreMan_UAT_GSLInvoice;Data Source=ncmluatdbserver.database.windows.net"
   'Gen_DbPath = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=s@ral52#91dB;Initial Catalog=Store-Man;Data Source=10.20.20.6" 'GetPath()
End If

cns = Gen_DbPath
cn.CommandTimeout = 0 'Gen_SysTime '100000


cn.Open cns

End Function

Public Function GetPath() As String
Dim Retval, Stmp As String
Retval = ""

Dim Pth As Variant
Pth = App.Path
Set cn = New ADODB.Connection

cns = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth & "\Store-ManDBCon.DAT; " & _
"Jet OLEDB:Database Password=M710211#S600607;"
cn.Open cns

Stmp = "select * from Mst_Config where ID = " & Gen_DBID
Set TmpRs0 = New ADODB.Recordset
TmpRs0.Open Stmp, cn, adOpenKeyset, adLockOptimistic

If TmpRs0.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs0!DbPath), "", TmpRs0!DbPath)
   Gen_DatFormat = IIf(IsNull(TmpRs0!DatFormat), "", TmpRs0!DatFormat)
End If
TmpRs0.Close
cn.Close
cns = ""
GetPath = Retval



End Function

Public Function StatusBarData(FName As Object)

FName.StatusBar1.Panels.Add
FName.StatusBar1.Panels.Add
FName.StatusBar1.Panels.Add
FName.StatusBar1.Panels.Add
FName.StatusBar1.Panels.Add
FName.StatusBar1.Panels(1).Text = "Login: " & Gen_UserName & "/" & UserType
FName.StatusBar1.Panels(2).Text = "Location : " & Gen_Location
FName.StatusBar1.Panels(3).Text = GetIpAddrTable '"Audited Upto  : " & AuditedDate
FName.StatusBar1.Panels(4).Text = "Store-Man Version " & Gen_VersionEXE ' 1.0.0.10.11"
FName.StatusBar1.Panels(5).Text = "Copyright : NCML "
FName.StatusBar1.Panels(6).Text = Gen_RDC_RiceMilled

FName.StatusBar1.Panels(1).Width = 2500
FName.StatusBar1.Panels(2).Width = 4000
FName.StatusBar1.Panels(3).Width = 2800
FName.StatusBar1.Panels(4).Width = 2500
FName.StatusBar1.Panels(5).Width = 2000
FName.StatusBar1.Panels(6).Width = 2500
End Function
Public Function GetFrmLoadAccess(FName As String) As Boolean()
Call Gen_ReadSysConfigINI
Dim Retval(3) As Boolean
Retval(0) = False
Retval(1) = False
Retval(2) = False
Retval(3) = False

Gen_DbPathFromAzure = "N"
tmp = "select b.HMenuFlag, b.FormName, b.MenuName, a.FormAccess, a.ADD_, a.EDIT_, a.Delete_, b.MenuID From " & _
"Mst_MenuHasUserRoleVBDB a inner join Mst_MenuDetailVBDB b on a.MenuID = b.MenuID Where a.UserGroupID =  " & Gen_UserTypeID & _
" And  b.FormName = '" & FName & "'"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Retval(0) = IIf(TmpRs!formaccess = "Y", True, False)
   Gen_FormAccess = Retval(0)
   Retval(1) = IIf(TmpRs!ADD_ = "Y", True, False)
   Retval(2) = IIf(TmpRs!EDIT_ = "Y", True, False)
   Retval(3) = IIf(TmpRs!Delete_ = "Y", True, False)
End If
GetFrmLoadAccess = Retval
Gen_DbPathFromAzure = "Y"
End Function



Public Function GButtonLock(FName As Object, yesNo As Boolean)
Call Gen_ReadSysConfigINI
If yesNo = True Then
   FName.SaveCmd.Enabled = False
   FName.SearchCmd.Caption = "Search"
Else
   FName.SaveCmd.Enabled = True
   FName.SearchCmd.Caption = "Cancel"
End If
FName.AddCmd.Enabled = yesNo
FName.EditCmd.Enabled = yesNo
FName.DeleteCmd.Enabled = yesNo
FName.SearchCmd.Enabled = True
FName.NextCmd.Enabled = yesNo
FName.PreviousCmd.Enabled = yesNo
FName.FirstCmd.Enabled = yesNo
FName.LastCmd.Enabled = yesNo
FName.ExitCmd.Enabled = True

Dim FrmBtnAccess() As Boolean
FrmBtnAccess = GetFrmLoadAccess(FName.Name)
FName.AddCmd.Enabled = IIf(FrmBtnAccess(1) = False, False, yesNo)
FName.EditCmd.Enabled = IIf(FrmBtnAccess(2) = False, False, yesNo)
FName.DeleteCmd.Enabled = IIf(FrmBtnAccess(3) = False, False, yesNo)

End Function

Public Function GButtonLockAD(FName As Object)


FName.SaveCmd.Enabled = False
FName.AddCmd.Enabled = False
FName.EditCmd.Enabled = False
FName.DeleteCmd.Enabled = False
FName.SearchCmd.Enabled = False
FName.NextCmd.Enabled = False
FName.PreviousCmd.Enabled = False
FName.FirstCmd.Enabled = False
FName.LastCmd.Enabled = False
FName.ExitCmd.Enabled = True

End Function


Public Function GButtonLock_1(FName As Object, yesNo As Boolean)
If yesNo = True Then
   FName.SaveCmd.Enabled = False
   FName.SearchCmd.Caption = "Search"
Else
   FName.SaveCmd.Enabled = True
   FName.SearchCmd.Caption = "Cancel"
End If
FName.AddCmd.Enabled = yesNo
FName.EditCmd.Enabled = yesNo
FName.DeleteCmd.Enabled = yesNo

FName.SearchCmd.Enabled = yesNo
FName.NextCmd.Enabled = yesNo
FName.PreviousCmd.Enabled = yesNo
FName.FirstCmd.Enabled = yesNo
FName.LastCmd.Enabled = yesNo
End Function


Public Function GButtonLock_Pwd(FName As Object, yesNo As Boolean)
If yesNo = True Then
   FName.SaveCmd.Enabled = False
   FName.SearchCmd.Caption = "Search"
Else
   FName.SaveCmd.Enabled = True
   FName.SearchCmd.Caption = "Cancel"
End If
FName.AddCmd.Enabled = yesNo
FName.EditCmd.Enabled = yesNo
FName.DeleteCmd.Enabled = yesNo

FName.SearchCmd.Enabled = yesNo
FName.NextCmd.Enabled = yesNo
FName.PreviousCmd.Enabled = yesNo
FName.FirstCmd.Enabled = yesNo
FName.LastCmd.Enabled = yesNo
End Function



Public Function DateValidation(FDate As Date) As Boolean
Dim Yn As Boolean

'Call ConfigMstTable
'
'If ConfigRs!AuditedUpto > FDate Then
'   MsgBox "Transaction Audited up to " & tmp & ""
'   'Return yn
'
'End If


End Function

'Public Function CustomerRs()
'SqlC = "select * from customer order by 'CustomerName'"
'Set CMrs = New ADODB.Recordset
'CMrs.Open SqlC, cn, adOpenKeyset, adLockOptimistic


'End Function


Public Function NumStrConv(Numb_ As Double) As String
Dim oNumb, Snum As Variant
oNumb = Numb_
NToS = ""

If InStr(Numb_, ".") > 0 Then
   Snum = mID(Numb_, 1, InStr(Numb_, ".") - 1)
   NumStrConvStr (Snum)
   Snum = mID(oNumb, InStr(oNumb, ".") + 1)
   If Len(Snum) = 1 Then
      Snum = Snum & "0"
   End If
   NToS = NToS & " and Paise "
   NumStrConvStr (Snum)
   NToS = NToS
Else
   Snum = oNumb
   NumStrConvStr (Snum)
End If
If NToS <> "" Then
   NToS = NToS & " Only"
End If
NumStrConv = NToS
End Function
Public Function NumStrConvStr(Numb As Double)
Dim tmp As Variant
Dim tmp1 As Variant
Dim aStr(20)
Dim ASt1(9)
aStr(1) = "One "
aStr(2) = "Two "
aStr(3) = "Three "
aStr(4) = "Four "
aStr(5) = "Five "
aStr(6) = "Six "
aStr(7) = "Seven "
aStr(8) = "Eight "
aStr(9) = "Nine "
aStr(10) = "Ten "
aStr(11) = "Eleven "
aStr(12) = "Twelve "
aStr(13) = "Thirteen "
aStr(14) = "Fourteen "
aStr(15) = "Fifteen "
aStr(16) = "Sixteen "
aStr(17) = "Seventeen "
aStr(18) = "Eighteen "
aStr(19) = "Nineteen "
aStr(20) = "Twenty "

ASt1(1) = " "
ASt1(2) = "Twenty "
ASt1(3) = "Thirty "
ASt1(4) = "Fourty "
ASt1(5) = "Fifty "
ASt1(6) = "Sixty "
ASt1(7) = "Seventy "
ASt1(8) = "Eighty "
ASt1(9) = "Ninty "



tmp = Trim(Str(Numb))
tmp1 = "000000000"

tmp = mID(tmp1, 1, 9 - Len(tmp)) & tmp


If Len(tmp) >= 9 Then
   'MsgBox Len(tmp)
   tmp1 = mID(tmp, 1, 2)
  
   If Val(tmp1) > 20 Then
      tmp1 = mID(tmp, 1, 1)
      NToS = NToS & ASt1(Val(tmp1))
      tmp1 = mID(tmp, 2, 1)
      NToS = NToS & aStr(Val(tmp1)) & "Crore "
   Else
      If tmp1 > 0 Then
         NToS = NToS & aStr(Val(tmp1)) & "Crore "
      End If
   End If
   tmp = mID(tmp, 3)
End If

If Len(tmp) >= 7 Then
   'MsgBox Len(tmp)
   tmp1 = mID(tmp, 1, 2)
   If Val(tmp1) > 20 Then
      tmp1 = mID(tmp, 1, 1)
      NToS = NToS & ASt1(Val(tmp1))
      tmp1 = mID(tmp, 2, 1)
      NToS = NToS & aStr(Val(tmp1)) & "Lac "
   Else
      If tmp1 > 0 Then
         NToS = NToS & aStr(Val(tmp1)) & "Lac "
      End If
   End If
   tmp = mID(tmp, 3)
   
End If

If Len(tmp) >= 5 Then
   'MsgBox Len(tmp)
   tmp1 = mID(tmp, 1, 2)
   
   If Val(tmp1) > 20 Then
      tmp1 = mID(tmp, 1, 1)
      NToS = NToS & ASt1(Val(tmp1))
      tmp1 = mID(tmp, 2, 1)
      NToS = NToS & aStr(Val(tmp1)) & "Thousand "
   Else
      If tmp1 > 0 Then
         NToS = NToS & aStr(Val(tmp1)) & "Thousand "
      End If
   End If
   tmp = mID(tmp, 3)
End If

If Len(tmp) = 3 Then
   'MsgBox Len(tmp)
   tmp1 = mID(tmp, 1, 1)
   If tmp1 > 0 Then
      NToS = NToS & aStr(Val(tmp1)) & "Hundred "
   End If
   tmp = mID(tmp, 2)
End If

If Len(tmp) >= 2 Then
   'MsgBox Len(tmp)
   tmp1 = mID(tmp, 1, 2)
   If Val(tmp1) > 20 Then
      tmp1 = mID(tmp, 1, 1)
      NToS = NToS & ASt1(Val(tmp1))
      tmp1 = mID(tmp, 2, 1)
      NToS = NToS & aStr(Val(tmp1))
   Else
      If tmp1 > 0 Then
         NToS = NToS & aStr(Val(tmp1))
      End If
   End If
   tmp = mID(tmp, 2)
End If



'MsgBox NToS
End Function







Public Function IncFtsConv(Numb As Double)
Dim tmp As Variant
Dim tmp1 As Variant
Dim AFts(20)
Dim ASt1(9)
AFts(1) = 0.083333
AFts(2) = 0.166666
AFts(3) = 0.25
AFts(4) = 0.333333
AFts(5) = 0.416666
AFts(6) = 0.5
AFts(7) = 0.583333
AFts(8) = 0.666666
AFts(9) = 0.75
AFts(10) = 0.83333
AFts(11) = 0.91666
Dim l As Integer
If Numb <> 0 Then NToS = AFts(Val(Numb))

Exit Function
NToS = ""
tmp = InStr(Str(Numb), ".")
If tmp <> 0 Then
   tmp1 = mID(Numb, tmp)
   NToS = mID(Numb, 1, tmp - 1) + AFts(Val(10))
Else
   NToS = Numb
End If

'MsgBox NToS
End Function
Public Function IsOprater(strToCheck) As Boolean
Dim lng As Long
Dim bln As Boolean
bln = False
If Len(strToCheck) > 0 Then
    For lng = 1 To Len(strToCheck)
        Select Case Asc(mID$(strToCheck, lng, 1))
        Case 60 To 62
        ' > < = <>
        bln = True 'False
        Exit For
        
        Case Else
            'bln = False
            'Exit For
        End Select
    Next lng
End If

IsOprater = bln
End Function

Public Function IsClean(strToCheck) As Boolean
Dim lng As Long
Dim bln As Boolean
bln = True
If Len(strToCheck) > 0 Then
    For lng = 1 To Len(strToCheck)
        Select Case Asc(mID$(strToCheck, lng, 1))
        Case 32
        ' Space
        Case 48 To 57
        ' Numbers 0 to 9
        Case 65 To 90
        ' Letters A to Z
        Case 97 To 122
        ' Letters a to z
        Case Else
            bln = False
            Exit For
        
        End Select
    Next lng
End If

IsClean = bln
End Function



Public Function RemoveSpChar(strToCheck) As String
Dim Rstr As Variant
Dim lng As Long
For lng = 1 To Len(strToCheck)
        Select Case Asc(mID$(strToCheck, lng, 1))
        Case 32
        Rstr = Rstr + mID$(strToCheck, lng, 1)
        Case 48 To 57
        ' Numbers 0 to 9
          Rstr = Rstr + mID$(strToCheck, lng, 1)
        Case 65 To 90
        ' Letters A to Z
          Rstr = Rstr + mID$(strToCheck, lng, 1)
        Case 97 To 122
        ' Letters a to z
          Rstr = Rstr + mID$(strToCheck, lng, 1)

        Case Else
             Rstr = Rstr + " "
        End Select
Next lng
RemoveSpChar = Rstr
End Function

Function LeadingZeros(ExpectedLen As Integer, ActualLen As Integer) As String
    LeadingZeros = String$(ExpectedLen - ActualLen, "0")
End Function

Function GetGUID() As String
    
    Dim lResult As Long
    Dim lguid As GUID
    Dim MyguidString As String
    Dim MyGuidString1 As String
    Dim MyGuidString2 As String
    Dim MyGuidString3 As String
    Dim DataLen As Integer
    Dim StringLen As Integer
    Dim I%
    
    On Error GoTo error_olemsg
    
    lResult = CoCreateGuid(lguid)
    
    If lResult = S_OK Then
        MyGuidString1 = Hex$(lguid.Data1)
        StringLen = Len(MyGuidString1)
        DataLen = Len(lguid.Data1)
        MyGuidString1 = LeadingZeros(2 * DataLen, StringLen) & MyGuidString1        'First 4 bytes (8 hex digits)
        
        MyGuidString2 = Hex$(lguid.Data2)
        StringLen = Len(MyGuidString2)
        DataLen = Len(lguid.Data2)
        MyGuidString2 = LeadingZeros(2 * DataLen, StringLen) & Trim$(MyGuidString2)        'Next 2 bytes (4 hex digits)
        
        MyGuidString3 = Hex$(lguid.Data3)
        StringLen = Len(MyGuidString3)
        DataLen = Len(lguid.Data3)
        MyGuidString3 = LeadingZeros(2 * DataLen, StringLen) & Trim$(MyGuidString3)        'Next 2 bytes (4 hex digits)
        
        GetGUID = MyGuidString1 & MyGuidString2 & MyGuidString3
        
        For I% = 0 To 7
           MyguidString = MyguidString & Format$(Hex$(lguid.Data4(I%)), "00")
        Next I%
        
        'MyGuidString contains last 8 bytes of Guid (16 hex digits)
        GetGUID = GetGUID & MyguidString
    Else
        GetGUID = "00000000" ' return zeros if function unsuccessful
    End If
    
    Exit Function

error_olemsg:
         MsgBox "Error " & Str(Err) & ": " & Error$(Err)
         GetGUID = "00000000"
         Exit Function

End Function


Public Function Gen_DateCheck(strToCheck, Typ_) As Boolean
Dim Dat_Chk As Boolean
Dat_Chk = True
If IsDate(strToCheck) = False Then
   MsgBox "Invalid Date Format !!!!!"
   Dat_Chk = False
End If
If Dat_Chk = True Then
    If CDate(strToCheck) <= AuditedDate Then
       MsgBox Typ_ & " Date Must be greater than Audited Transaction " & AuditedDate & " !!!!! "
       'TxtPRADate.SetFocus
       Dat_Chk = False
    End If
    If CDate(strToCheck) > Date Then
       MsgBox "Future Transaction Date Not Allowed  !!!!! "
       'TxtPRADate.SetFocus
       Dat_Chk = False
    End If
End If
Gen_DateCheck = Dat_Chk
End Function
Public Function GenWc(wc_ As String, ID As Variant, Field_ As String, Type_ As String, Mod_ As String) As String
Dim Retval
Retval = wc_
Dim Eql As String
If Type_ = "S" And Mod_ = "N" Then
   Eql = "= '" & ID & "'"
ElseIf Type_ = "S" And Mod_ = "L" Then
    Eql = " like '%" & ID & "%'"
ElseIf Type_ = "S" And Mod_ = "I" Then
    Eql = ID
ElseIf Type_ = "N" And Mod_ = "I" Then
    Eql = ID
Else
    Eql = " = " & ID
End If

If Retval = "" Then
   If Type_ = "S" Then
      Retval = "Where " & Field_ & Eql
   Else
      Retval = "Where " & Field_ & Eql
   End If
Else
   If Type_ = "S" Then
      Retval = Retval & " And " & Field_ & Eql
   Else
      Retval = Retval & " And " & Field_ & Eql
   End If
End If
GenWc = Retval
End Function

Public Sub Gen_Create_Xls(FileName_ As String)

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Set oSheet = oBook.Worksheets(1)
oBook.SaveAs (Pat & "\excel\" & FileName_)

oExcel.Quit

End Sub

Public Function AssignHelpData(FName As Object, FieldName_ As String)

Gen_FormName = FName.Name
Gen_FieldName = FieldName_
Unload FrmHelp
FrmHelp.Show
End Function
Public Function Gen_SendMail(FName As Object, ToADD_ As String, semail_ As String, MsgSubject_ As String, Message1_ As String)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue>"
htmlStringtail = "</Font></body></html>"
'servertxt = "172.16.20.5"
'ret = SimplePing("mail.ncml.com")
    
'With oSMTP
'  'connection
'
'  .Server = servertxt
'  'message
'  .MailFrom = "parag.r@ncml.com"
'  .SendTo = ToADD_
'  .MessageSubject = MsgSubject_ '"Confidential Mail"
''  .MessageText = MsgSubject_ & " :- " & Message1_ '"This is a test message"
'  .MessageHTML = htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail
'  '.MessageHTML = htmlstring
'  'send email
'  .TimeStamp = Now()
'  .SendEmail
'End With

Dim objSMTP
' Create mailer component
Set objSMTP = CreateObject("MailBee.SMTP")
htmlStringhead = "<html><body><font size= 3 face= ""Candara"" color = Blue>"
htmlStringtail = "</Font></body></html>"
' Unlock SMTP component
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"

' Set SMTP server name
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "10.20.20.13" '"mail.ncml.com" '"mail.ncml.com"
objSMTP.EnableLogging = True
objSMTP.LogFilePath = "C:\temp\StoreMan\pwdlog.txt"
' Set message properties
objSMTP.FromAddr = "paassuport@ncml.com"
objSMTP.ToAddr = ToADD_
objSMTP.Subject = MsgSubject_ '"Message in HTML format"

' Set HTML format for the body
objSMTP.BodyFormat = 1

' Place HTML contents into the body
objSMTP.BodyText = htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"

' Send it!
objSMTP.Send


End Function

Public Function GetEndDate(mMonth_ As Variant, mStartDate_ As Variant)
Dim Retval As Variant

If mMonth_ = 1 Or mMonth_ = 3 Or mMonth_ = 5 Or mMonth_ = 7 Or mMonth_ = 8 Or mMonth_ = 10 Or mMonth_ = 12 Then
   Retval = DateAdd("D", 31, mStartDate_)
ElseIf mMonth_ = 4 Or mMonth_ = 6 Or mMonth_ = 9 Or mMonth_ = 11 Then
   Retval = DateAdd("D", 30, mStartDate_)
ElseIf mMonth_ = 2 Then
   If (Year(mStartDate_) Mod 4) = 0 Then
      Retval = DateAdd("D", 29, mStartDate_)
   Else
      Retval = DateAdd("D", 28, mStartDate_)
   End If
End If

GetEndDate = Retval - 1

End Function


Public Sub Gen_ReadSysConfigINI()

Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "SysConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)

sFileLine = objTS.ReadLine

If UCase(mID(Trim(sFileLine), 1, 1)) = "Y" Then
   Beep
   MsgBox mID(Trim(sFileLine), 3, 100)
   End
End If
objTS.Close
End Sub

Public Function CheckActiveUsers(mUserID_ As Variant) As Boolean
Dim Retval As Boolean
Retval = True
Call TableMst_Employee(" Where EmployeeNo = '" & mUserID_ & "'")
If RsMst_Employee.RecordCount > 0 Then
   If LCase(RsMst_Employee!Flag) = "n" Then
      Retval = False
   End If
Else
   Retval = False
End If
CheckActiveUsers = Retval
End Function

Public Function GetDefaultPrinter() As Printer
    Dim strBuffer As String * 254
    Dim iRetValue As Long
    Dim strDefaultPrinterInfo As String
    Dim tblDefaultPrinterInfo() As String
    Dim objPrinter As Printer
   
    ' Retreive current default printer information
    iRetValue = GetProfileString("windows", "device", ",,,", strBuffer, 254)
    strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
    For Each objPrinter In Printers
        If objPrinter.DeviceName = tblDefaultPrinterInfo(0) Then
            ' Default printer found !
            Gen_DefaultPrinter = True
            Exit For
        End If
    Next
    ' If not found, return nothing
    If tblDefaultPrinterInfo(0) = "" Then
        Set objPrinter = Nothing
        Gen_DefaultPrinter = False
    End If
    Set GetDefaultPrinter = objPrinter
End Function
Public Function RoundUp(dblValue As Double) As Double
On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
   RoundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
Else
   RoundUp = dblValue
End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox Err.Description, vbInformation, "Round Up"
End Function

Public Function DBCHRPro()
Dim Pth As Variant
'28423229 Sankar
'9821146495

Pth = App.Path
Set cnHRPro = New ADODB.Connection
'cns = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;pwd=sa;Initial Catalog=CM;Data Source=(local)"
cns = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=CM;Data Source=(local)"

Gen_DBType = "SQL"
'If Gen_DbPath = "" Then
'  Gen_DbPath = GetPath()
'End If
cns = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=$QL@#$%67890;Initial Catalog=HRPRo;Data Source=10.23.12.49,15333"
'cns = GetPathHRPro
cnHRPro.Open cns
cnHRPro.CommandTimeout = 0
End Function

Public Function GetPathHRPro() As String
Dim Retval, Stmp As String
Retval = ""

Dim Pth As Variant
Pth = App.Path
Set cn = New ADODB.Connection

cns = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth & "\HRD_VBDBCon.DAT; " & _
"Jet OLEDB:Database Password=M710211#S600607;"
cn.Open cns

Stmp = "select * from Mst_Config where ID =1 " ' & Gen_DBID
Set TmpRs0 = New ADODB.Recordset
TmpRs0.Open Stmp, cn, adOpenKeyset, adLockOptimistic

If TmpRs0.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs0!DbPath), "", TmpRs0!DbPath)
   Gen_DatFormat = IIf(IsNull(TmpRs0!DatFormat), "", TmpRs0!DatFormat)
End If
TmpRs0.Close
cn.Close
cns = ""
'GetPathHRPro = Retval
GetPathHRPro = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=$QL@#$%67890;Initial Catalog=HRPro;Data Source=10.23.12.49,15333"


End Function

Public Function DBCCMPro()
Dim Pth As Variant

Pth = App.Path
Set cnCMPro = New ADODB.Connection

cns = GetPathCMPro
cnCMPro.Open cns
End Function
Public Function GetPathCMPro() As String
Dim Retval, Stmp As String

Retval = ""

Dim Pth As Variant
Pth = App.Path
Set cn = New ADODB.Connection

cns = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth & "\CMProDb.mdb; " & _
"Jet OLEDB:Database Password=M710211#S600607;"
cn.Open cns

Stmp = "Select * from Mst_Config where ID = 1"
Set TmpRs0 = New ADODB.Recordset
TmpRs0.Open Stmp, cn, adOpenKeyset, adLockOptimistic

If TmpRs0.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs0!DbPath), "", TmpRs0!DbPath)
   Gen_DatFormat = IIf(IsNull(TmpRs0!DatFormat), "", TmpRs0!DatFormat)
End If
TmpRs0.Close
cn.Close
cns = ""
'GetPathCMPro = Retval
GetPathCMPro = "Provider=MSOLEDBSQL;Persist Security Info=False;User ID=appuser;pwd=$QL@#$%67890;Initial Catalog=CMPro;Data Source=10.23.12.49,15333"

End Function

Public Sub Gen_ReadEmailConfigINI()

Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)

sFileLine = objTS.ReadLine
Gen_EmailConfigFlag = UCase(mID(Trim(sFileLine), 1, 1))
objTS.Close


End Sub
Public Sub Gen_SysTimeINI()

Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "SysTime.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)

sFileLine = objTS.ReadLine
Gen_SysTime = Trim(sFileLine)
objTS.Close


End Sub

Public Function GetPathSMAttachment() As String
Dim Retval, Stmp As String
Retval = ""

Dim Pth As Variant
Pth = App.Path
Set cn = New ADODB.Connection

cns = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth & "\Store-ManAttachDBCon.DAT; " & _
"Jet OLEDB:Database Password=M710211#S600607;"
cn.Open cns

Stmp = "select * from Mst_Config where ID = 1"

Set TmpRs0 = New ADODB.Recordset

TmpRs0.Open Stmp, cn, adOpenKeyset, adLockOptimistic

If TmpRs0.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs0!DbPath), "", TmpRs0!DbPath)
   Gen_DatFormat = IIf(IsNull(TmpRs0!DatFormat), "", TmpRs0!DatFormat)
End If

TmpRs0.Close
cn.Close
cns = ""

GetPathSMAttachment = Retval

End Function

Public Function DBCSM_Attachment()
Dim Pth As Variant
Pth = App.Path
Set cnSM_Attachment = New ADODB.Connection

cns = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=CM;Data Source=(local)"

cns = GetPathSMAttachment
cnSM_Attachment.CommandTimeout = 0 '100000
cnSM_Attachment.Open cns

End Function

Public Function LoadFileFromDB(FileName As String, _
  Rs As Object, FieldName As String) As Boolean
Dim iFileNum As Integer
Dim lFileLength As Long
Dim abBytes() As Byte
Dim iCtr As Integer

On Error GoTo ErrorHandler
If Not TypeOf Rs Is ADODB.Recordset Then Exit Function

iFileNum = FreeFile
Open FileName For Binary As #iFileNum
lFileLength = LenB(Rs(FieldName))

abBytes = Rs(FieldName).GetChunk(lFileLength)
Put #iFileNum, , abBytes()
Close #iFileNum
LoadFileFromDB = True

ErrorHandler:
End Function

Public Function SaveFileToDB(ByVal FileName As String, _
   Rs As Object, FieldName As String) As Boolean

Dim iFileNum As Integer
Dim lFileLength As Long

Dim abBytes() As Byte
Dim iCtr As Integer

'On Error GoTo ErrorHandler
'If Dir(FileName) = "" Then Exit Function
If Not TypeOf Rs Is ADODB.Recordset Then Exit Function

'read file contents to byte array
iFileNum = FreeFile
Open FileName For Binary Access Read As #iFileNum
lFileLength = LOF(iFileNum)
ReDim abBytes(lFileLength)
Get #iFileNum, , abBytes()

'put byte array contents into db field
Rs.Fields(FieldName).AppendChunk abBytes()
Close #iFileNum

SaveFileToDB = True
ErrorHandler:
End Function

Function GetFileFromPath(vPath As Variant)

Dim Items() As String
Items = Split(vPath, "\")
If UBound(Items) = -1 Then Exit Function
GetFileFromPath = Items(UBound(Items))

End Function


Public Sub Gen_ReadEmailGodownViability()

Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailIDsforGodownViability.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)

sFileLine = objTS.ReadLine
Gen_EmailGodownViability = Trim(sFileLine)
objTS.Close
End Sub

Public Function GetMaxID(mColumn_ As Variant, mTable_ As Variant)
Dim Retval As Double
tmp = "Select Max(" & mColumn_ & ") from " & mTable_ & ""
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
GetMaxID = Retval
End Function

Public Function Gen_ReadPathConfigINI(mType As Variant) As String
Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "PathConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""
Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 7)) = LCase(mType) Then
      mstr = mID(Trim(sFileLine), 9, Len(sFileLine))
      Exit Do
   End If
Loop

objTS.Close
Gen_ReadPathConfigINI = mstr
End Function

Public Function Gen_ReadMailServerIPINI(mType As Variant) As String

Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 6)) = LCase(mType) Then
      mstr = mID(Trim(sFileLine), 8, Len(sFileLine))
      Exit Do
   End If
Loop

objTS.Close
Gen_ReadMailServerIPINI = mstr

End Function

Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim I As Integer
   For I = 0 To NrOfEntries - 1
      Dim j As Integer, S As String: S = ""
      For j = 0 To 3: S = S & IIf(j > 0, ".", "") & Buf(4 + I * 24 + j): Next
      IpAddrs(I) = S
   Next
   GetIpAddrTable = IpAddrs(0) & "-" & IpAddrs(1)
   
End Function

Public Function Gen_ReadNFinMailIDINI(mType As Variant) As String

Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, Len(mType))) = LCase(mType) Then
      mstr = mID(Trim(sFileLine), Len(mType) + 2, Len(sFileLine))
      Exit Do
   End If
Loop

objTS.Close
Gen_ReadNFinMailIDINI = mstr

End Function

''-- To Read Navision Edit User Role From EmailConfig.ini file
Public Function Gen_ReadNavisionEditUserRole() As String
Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 17)) = LCase("NavisionEditRole-") Then
      mstr = sFileLine
      Exit Do
   End If
Loop
objTS.Close
Gen_ReadNavisionEditUserRole = mstr
''NavisionEditRole-2047,1076,1569,2756
End Function

''-- To Read New Letter Format FAG BankIDs From EmailConfig.ini file
Public Function Gen_ReadLetterFormatFAGBankID() As String
Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 22)) = LCase("LetterFormatFAGBankID-") Then
      mstr = Right(sFileLine, Len(sFileLine) - 22)
      Exit Do
   End If
Loop
objTS.Close
Gen_ReadLetterFormatFAGBankID = mstr
''--LetterFormatFAGBankID-15
End Function

Public Function Gen_ReadPaymentBank() As String
Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 14)) = LCase("PaymentByBank-") Then
      mstr = sFileLine
      Exit Do
   End If
Loop
objTS.Close
Gen_ReadPaymentBank = mstr
End Function

Public Function ValidateGSTIN(ByVal mTINNo_ As Variant, ByVal mGSTINNo_ As String) As String
Dim mGSTIN01, mGSTIN02, mGSTIN03, mGSTIN04, mGSTIN05, ErrorMsg As Variant
Dim Retval As Boolean
Retval = True

If Len(mGSTINNo_) <> 15 Then
    ErrorMsg = "Invalid GST Number due to Length of GST No."
    Retval = False
    GoTo ExitFunction
End If

If IsClean(mGSTINNo_) = False Then
    ErrorMsg = "Special character not allowed in GST No."
    Retval = False
    GoTo ExitFunction
End If


'Commented on 23 Aug 2017 for allowing upload of FARmer cases by Parag

'mGSTIN01 = mID(mGSTINNo_, 1, 2)
'mGSTIN02 = mID(mGSTINNo_, 3, 10)
'mGSTIN03 = mID(mGSTINNo_, 13, 1)
'mGSTIN04 = mID(mGSTINNo_, 14, 1)
'mGSTIN05 = mID(mGSTINNo_, 15, 1)
'If Trim(mGSTIN01) <> Trim(mTINNo_) Then
'    ErrorMsg = "State TIN Code is mismatch in given GST No."
'    Retval = False
'    GoTo ExitFunction
'End If
'If ValidatePANCard(mGSTIN02) = False Then
'   ErrorMsg = "Invalid PAN No is enter in given GST No."
'   Retval = False
'   GoTo ExitFunction
'End If

ExitFunction:
ValidateGSTIN = Retval & "~" & ErrorMsg
End Function

Public Function ValidatePANCard(ByVal mPanEntry As String) As Boolean
Dim I As Integer
ValidatePANCard = True
If Len(mPanEntry) <> 10 Then
   ValidatePANCard = False
   Exit Function
End If
If Not IsNumeric(mID(mPanEntry, 6, 4)) Then
   ValidatePANCard = False
   Exit Function
End If
If Not CheckAtoZ(mID(mPanEntry, 10, 1)) Then
   ValidatePANCard = False
   Exit Function
End If
For I = 1 To 5
    If Not CheckAtoZ(mID(mPanEntry, I, 1)) Then
       ValidatePANCard = False
       Exit Function
    End If
Next
End Function

Public Function CheckAtoZ(ByVal mChar As String) As Boolean
CheckAtoZ = True
If ((Asc(mChar) < 65) Or (Asc(mChar) > 90)) Then
   CheckAtoZ = False
End If
End Function

Public Function ValidateTANNo(ByVal mTAN_ As String) As Boolean
Dim I As Integer
ValidateTANNo = True
If Len(mTAN_) > 0 Then
    For I = 1 To Len(mTAN_)
        Select Case Asc(mID$(mTAN_, I, 1))
        ''Case 32       '' Space
        Case 48 To 57   ' Numbers 0 to 9
        Case 65 To 90   ' Letters A to Z
        Case 97 To 122  ' Letters a to z
        Case Else
            ValidateTANNo = False
            Exit Function
        End Select
    Next I
End If

If Len(mTAN_) <> 10 Then
   ValidateTANNo = False
   Exit Function
End If
If Not IsNumeric(mID(mTAN_, 5, 5)) Then
   ValidateTANNo = False
   Exit Function
End If
If Not CheckAtoZ(mID(mTAN_, 10, 1)) Then
   ValidateTANNo = False
   Exit Function
End If
For I = 1 To 4
    If Not CheckAtoZ(mID(mTAN_, I, 1)) Then
       ValidateTANNo = False
    End If
Next
End Function

Public Function Gen_GSTDataUpdateUserRole() As Boolean
Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant

Dim msplitUserRole As Variant
Dim mResult As Boolean
Dim mI As Double

Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 22)) = LCase("GSTDataUpdateUserRole-") Then   ''"GSTDataUpdateUserRole-"
      mstr = Right(sFileLine, Len(sFileLine) - 22)
        If mstr <> "" Then
            msplitUserRole = Split(mstr, ",")
            For mI = 0 To UBound(msplitUserRole)
                If LCase(Gen_UserRole) = LCase(msplitUserRole(mI)) Then
                    mResult = True
                End If
            Next mI
        End If
   End If
Loop

objTS.Close
Gen_GSTDataUpdateUserRole = mResult
End Function

Public Function Gen_DepositContractCalculationFactor(mFactorString As String) As String
Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 39)) = LCase(mFactorString) Then   ''"DepositContractCurrentConversionFactor-"
      mstr = Right(sFileLine, Len(sFileLine) - 39)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 30)) = LCase(mFactorString) Then   ''"DepositContractDivisionFactor-"
      mstr = Right(sFileLine, Len(sFileLine) - 30)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 33)) = LCase(mFactorString) Then   ''"DepositContractFumigationFactorY-"
      mstr = Right(sFileLine, Len(sFileLine) - 33)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 33)) = LCase(mFactorString) Then   ''"DepositContractFumigationFactorN-"
      mstr = Right(sFileLine, Len(sFileLine) - 33)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 39)) = LCase(mFactorString) Then   ''"DepositContractCMPMtoAMRMApprovalLimit-"
      mstr = Right(sFileLine, Len(sFileLine) - 39)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 40)) = LCase(mFactorString) Then   ''"DepositContractNonKnownRentMinimumValue-"
      mstr = Right(sFileLine, Len(sFileLine) - 40)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 36)) = LCase(mFactorString) Then   ''"DepositContractCMGGroupToBeApproved-"
      mstr = Right(sFileLine, Len(sFileLine) - 36)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 32)) = LCase(mFactorString) Then   ''"DepositContractDefaultDaysRange-"
      mstr = Right(sFileLine, Len(sFileLine) - 32)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 32)) = LCase(mFactorString) Then   ''"DepositContractInFutureDateFlag-"
      mstr = Right(sFileLine, Len(sFileLine) - 32)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 37)) = LCase(mFactorString) Then   ''"DepositContractDefaultAttachmentFlag-"
      mstr = Right(sFileLine, Len(sFileLine) - 37)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 38)) = LCase(mFactorString) Then   ''"DepositContractAllowBackDateEntryDays-"
      mstr = Right(sFileLine, Len(sFileLine) - 38)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 43)) = LCase(mFactorString) Then   ''"DepositContractAdditionalGroupToBeApproved-"
      mstr = Right(sFileLine, Len(sFileLine) - 43)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 26)) = LCase(mFactorString) Then   ''"DepositContractChangeFlag-"
      mstr = Right(sFileLine, Len(sFileLine) - 26)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 26)) = LCase(mFactorString) Then   ''"DepositContractChangeDate-"
      mstr = Right(sFileLine, Len(sFileLine) - 26)
      Exit Do
   End If
Loop
objTS.Close
Gen_DepositContractCalculationFactor = mstr
End Function

Public Function Gen_NFinPledgeCalculationFactor(mFactorString As String) As String
Dim mstr As String
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim lngFreefile As Long
Dim sFileLine As Variant
Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(App.Path & "\" & "EmailConfig.ini")
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
mstr = ""

Do While Not objTS.AtEndOfStream
   sFileLine = objTS.ReadLine
   If LCase(mID(sFileLine, 1, 28)) = LCase(mFactorString) Then   ''"NFinFundingAmountPercentage-"
      mstr = Right(sFileLine, Len(sFileLine) - 28)
      Exit Do
   ElseIf LCase(mID(sFileLine, 1, 11)) = LCase(mFactorString) Then   ''"NFinFeesPA-"
      mstr = Right(sFileLine, Len(sFileLine) - 11)
      Exit Do
   End If
Loop

objTS.Close
Gen_NFinPledgeCalculationFactor = mstr
End Function

Public Function DBCAzureStoreMan()
Dim Pth As Variant

Pth = App.Path
Set cnAzureStoreMan = New ADODB.Connection

If Gen_AzureDbPath = "" Then
   Gen_AzureDbPath = GetPathAzureStoreMan()
End If
cnAzureStoreMan.CommandTimeout = 0 'Gen_SysTime '100000
cns = Gen_AzureDbPath
cnAzureStoreMan.Open cns

End Function
Public Function GetPathAzureStoreMan() As String
Dim Retval, Stmp As String

Retval = ""

Dim Pth As Variant
Pth = App.Path
Set cn = New ADODB.Connection

cns = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth & "\AzureStoreManDb.mdb; " & _
"Jet OLEDB:Database Password=M710211#S600607;"
cn.Open cns

Stmp = "Select * from Mst_Config where ID = 1"
Set TmpRs0 = New ADODB.Recordset
TmpRs0.Open Stmp, cn, adOpenKeyset, adLockOptimistic

If TmpRs0.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs0!DbPath), "", TmpRs0!DbPath)
   Gen_DatFormat = IIf(IsNull(TmpRs0!DatFormat), "", TmpRs0!DatFormat)
End If
TmpRs0.Close
cn.Close
cns = ""
GetPathAzureStoreMan = Retval
End Function


Public Function GetQueyForWebservice(TabName As Variant, Flag As Variant, RsQuery As ADODB.Recordset, Optional WhereClause As Variant, Optional mQuery As Variant, Optional mExcludeCol As Variant) As Variant

Dim iCtr, jCtr, mColumnList, mColumnValue, mUpdateColumnList, ID, PrimaryColName As Variant

mColumnList = "": mColumnValue = "": mUpdateColumnList = ""

Select Case Flag

Case "A"
     For iCtr = 0 To RsQuery.Fields.Count - 1
         
            If GetExcludedCol(RsQuery.Fields.Item(iCtr).Name, mExcludeCol) = False Then
         
                If LCase(RsQuery.Fields.Item(iCtr).Name) = "foreign" Then
                   mColumnList = IIf(mColumnList = "", "[" & RsQuery.Fields.Item(iCtr).Name & "]", mColumnList & "," & "[" & RsQuery.Fields.Item(iCtr).Name & "]")
                Else
                   mColumnList = IIf(mColumnList = "", RsQuery.Fields.Item(iCtr).Name, mColumnList & "," & RsQuery.Fields.Item(iCtr).Name)
                End If
                
                If mColumnValue = "" Then
                   mColumnValue = GetDataTypeOfRecordSet(RsQuery, iCtr)
                Else
                   mColumnValue = mColumnValue & "," & GetDataTypeOfRecordSet(RsQuery, iCtr)
                End If
           End If
     Next
     tmp = """Insert into " & TabName & ""
     tmp = tmp & " (" & mColumnList & ")"
     tmp = tmp & " Values"
     tmp = tmp & " (" & mColumnValue & ")"""
Case "E"
     If mQuery = "" Then
        For iCtr = 0 To RsQuery.Fields.Count - 1
            If GetExcludedCol(RsQuery.Fields.Item(iCtr).Name, mExcludeCol) = False Then
                    
                    If LCase(RsQuery.Fields.Item(iCtr).Name) <> LCase(PrimaryColName) And LCase(RsQuery.Fields.Item(iCtr).Name) <> "mkey1" Then
                       If LCase(RsQuery.Fields.Item(iCtr).Name) = "foreign" Then
                          mColumnList = "[" & RsQuery.Fields.Item(iCtr).Name & "]"
                       Else
                          mColumnList = RsQuery.Fields.Item(iCtr).Name
                       End If
                       mColumnValue = GetDataTypeOfRecordSet(RsQuery, iCtr)
                       If mUpdateColumnList = "" Then
                          mUpdateColumnList = mColumnList & "=" & mColumnValue
                       Else
                          mUpdateColumnList = mUpdateColumnList & "," & mColumnList & "=" & mColumnValue
                       End If
                    End If
            End If
        Next
        If LCase(WhereClause) = "all" Then
           mUpdateColumnList = mUpdateColumnList
        Else
           mUpdateColumnList = mUpdateColumnList & " " & WhereClause
        End If
        
        tmp = """Update " & TabName & " Set "
        tmp = tmp & mUpdateColumnList & """"
     Else
        tmp = """ " & mQuery & " """
     End If
Case "D"
     If LCase(WhereClause) <> "all" Then
        tmp = """Delete From " & TabName & " " & WhereClause & """"
     Else
        tmp = """Delete From " & TabName & " """
     End If
Case Else
     tmp = ""
End Select
GetQueyForWebservice = Replace(tmp, """", "")
End Function
Public Function GetExcludedCol(mColName As Variant, mExcludeCol_ As Variant) As Boolean
Dim Retval As Boolean
Dim strSplitString, xyz As Variant
Retval = False

strSplitString = Split(mExcludeCol_, ",")

For xyz = 0 To UBound(strSplitString)
    If LCase(mColName) = "" & LCase(strSplitString(xyz)) & "" Then
       Retval = True
       Exit For
    End If
Next

GetExcludedCol = Retval
End Function


Public Function GetDataTypeOfRecordSet(Rst As ADODB.Recordset, I As Variant) As Variant

Dim mValue As Variant
mValue = Rst.Fields(I).Value
If InStr(1, mValue, "'", vbTextCompare) > 0 Then
   mValue = Replace(mValue, "'", "`", 1)
End If

If InStr(1, mValue, Chr(34), vbTextCompare) > 0 Then
   mValue = Replace(mValue, Chr(34), "`", 1)
End If

Select Case Rst.Fields.Item(I).Type

Case adChar
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "'" & mValue & "'")
Case adVarChar
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "'" & mValue & "'")
Case adSmallInt
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "" & Rst.Fields(I).Value & "")
Case adUnsignedTinyInt
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "" & Rst.Fields(I).Value & "")
Case adDBTimeStamp
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "'" & mValue & "'")
Case adNumeric
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "" & Rst.Fields(I).Value & "")
Case adLongVarBinary
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "" & Rst.Fields(I).Value & "")
Case adVarWChar
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "'" & mValue & "'")
Case adVarNumeric
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "'" & mValue & "'")
Case adInteger
     GetDataTypeOfRecordSet = IIf(IsNull(Rst.Fields(I).Value), "NULL", "" & Rst.Fields(I).Value & "")
End Select

End Function


