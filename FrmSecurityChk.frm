VERSION 5.00
Begin VB.Form FrmSecurityChk 
   Caption         =   "System Error :"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Lbladdress 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSystemname 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "FrmSecurityChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*
' Lbladdress Label For Display System IP Address
' lblSystemname  For Display System Name
'*/
Dim Flag As Variant
 Option Explicit
   Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
   Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

   Private Const WS_VERSION_REQD = &H101
   Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
   Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
   Private Const MIN_SOCKETS_REQD = 1
   Private Const SOCKET_ERROR = -1
   Private Const WSADescription_Len = 256
   Private Const WSASYS_Status_Len = 128
   Private Type HOSTENT
       hName As Long
       hAliases As Long
       hAddrType As Integer
       hLength As Integer
       hAddrList As Long
   End Type

   Private Type WSADATA
       wversion As Integer
       wHighVersion As Integer
       szDescription(0 To WSADescription_Len) As Byte
       szSystemStatus(0 To WSASYS_Status_Len) As Byte
       iMaxSockets As Integer
       iMaxUdpDg As Integer
       lpszVendorInfo As Long
   End Type

 

   Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
   Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
   Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
   Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen&) As Long
   Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
   Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Sub SocketsInitialize()
   Dim WSAD As WSADATA
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, sMsg As String
       iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
       If iReturn <> 0 Then
           MsgBox "Winsock.dll is not responding."
           End
       End If
       If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
           sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
           sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
           sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
           sMsg = sMsg & " is not supported by winsock.dll "
           MsgBox sMsg
           End
       End If
       If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
           sMsg = "This application requires a minimum of "
           sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
           MsgBox sMsg
           End
       End If

       End Sub

  Function hibyte(ByVal wParam As Integer)
       hibyte = wParam \ &H100& And &HFF&
   End Function

    Function lobyte(ByVal wParam As Integer)
       lobyte = wParam And &HFF&
   End Function

   Sub SocketsCleanup()
   Dim lReturn As Long
       lReturn = WSACleanup()
       If lReturn <> 0 Then
           MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup """
           End
       End If
   End Sub

Private Sub CmdOk_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Gen_FreezOpnDate = CDate("31/10/2007")
Shell App.Path & "\LCIDsrv.exe"
tmp = "SELECT Mst_Rice_Miller.* From Mst_Rice_Miller WHERE (((Mst_Rice_Miller.RM_ID)=0));"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MsgBox "RM-ID 0 is not allowed; Please get in touch with Sankaranarayanan (02230419145) or Mayur Dhamani (022-30419172) for rectifying this error immediately"
   End
End If

lblSystemname.Caption = "System Name  :- "
SocketsInitialize
 Dim hostname As String * 256
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim I As Integer
   Dim ip_address As String
       If gethostname(hostname, 256) = SOCKET_ERROR Then

           MsgBox "Windows Sockets error " & Str(WSAGetLastError())
       Else
           hostname = Trim$(hostname)
       End If
       hostent_addr = gethostbyname(hostname)
       If hostent_addr = 0 Then
           MsgBox "Winsock.dll is not responding."
       End If
       RtlMoveMemory host, hostent_addr, LenB(host)
       RtlMoveMemory hostip_addr, host.hAddrList, 4
       ReDim temp_ip_address(1 To host.hLength)
       RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
       For I = 1 To host.hLength
           ip_address = ip_address & temp_ip_address(I) & "."
       Next
       ip_address = mID$(ip_address, 1, Len(ip_address) - 1)
       lblSystemname.Caption = lblSystemname.Caption + hostname
       Lbladdress.Caption = "System Address   :- " + ip_address
       tmp = mID(hostname, 1, 5)
       If UCase(tmp) = "NCMSL" Or tmp = "abc-e" Or ip_address = "192.168.1.29" Then
          Label1.Caption = "Valid User"
          '----------
          'MsgBox ip_address ip_address = "mail.ncml.com6" Or
          If ip_address = "10.0.0.20" Or ip_address = "10.0.0.21" Or ip_address = "220.227.198.49" Or ip_address = "220.227.198.52" Then
             ' By pass check for Citrix Login user
             'MsgBox "By Pass " & ip_address
          Else
                'MsgBox "Go for check"
                Dim Sd, Td  As Variant
                Dim strSplitString() As String
                strSplitString = Split(Date, "/")
                'On Error GoTo M_lbl:
                Td = strSplitString(0)
                Sd = Month(Date)
                Td = Day(Date)
                If Month(Date) <> Day(Date) Then
                   If Month(Date) = strSplitString(0) Then
                      MsgBox "Check your Date Setting.It Must be UK. "
                      End
                   End If
                End If
'                'Call Check_RegionalSetting
'                If Flag = "E" Or Flag = "" Then
'                   If Flag = "" Then
'                      MsgBox "Cannot detect your regional settings!! Please contact Mayur (022-30419172) immediately!! "
'                   Else
'                    MsgBox "Check your Date Setting.It Must be UK. "
'                   End If
'                   End
'                End If
          End If
          '---------
         'MsgBox "Flag " & Flag
         'FrmMain.Show
         'frmCheckVersion.Show
         If Val(Gen_OryLoc_ID) = 0 Then
            'FrmLocationMap.Show
            Unload Me
            'Unload frmCheckVersion
            Exit Sub
         End If

         'FrmStart.Show
         Unload Me
         'Unload frmCheckVersion
         Exit Sub
       Else
          Label1.Caption = "System Error : In-Valid User contact Technology Department "
       End If
       'End
       
End Sub

 
Public Sub Check_RegionalSetting()
'  MsgBox "Inside Check"
  Dim lcid As Long
  Dim ShortDateFormat As String
  Dim aStr2           As String
  Dim aStr1           As String
  Dim oLCID As Object
  Dim aStr            As String
  'MsgBox "Line 1 "
  aStr1 = String(255, Chr(0))
  aStr2 = String(255, Chr(0))
  ' Use out of proc server to get LCID in separate thread
  'MsgBox "Line 2 " & aStr1 & " " & aStr2
  Set oLCID = CreateObject("lcidsrvr.class1")
  'MsgBox "Line 3 "
  lcid = oLCID.GetLCID
  'MsgBox "Line 4 "
  'Text1.Text = lcid & " (" & Hex$(lcid) & ")"
  GetLocaleInfo lcid, 31, aStr1, 255
  aStr1 = Left(aStr1, InStr(1, aStr1, Chr(0)) - 1)
  Label1.Caption = aStr1
  GetLocaleInfo lcid, 2, aStr2, 255
  aStr2 = Left(aStr2, InStr(1, aStr2, Chr(0)) - 1)
  'Label1.Caption = aStr2 & "-" & aStr1
  'MsgBox lcid
  If lcid <> 2057 Then
     Flag = "E"
  End If
End Sub
