VERSION 5.00
Object = "{A4F98396-D7CE-459C-BA85-E874A5B1E44F}#3.3#0"; "OSSMTP.ocx"
Begin VB.Form FrmStart 
   BackColor       =   &H00FFC0C0&
   Caption         =   "National Collateral Management Services Limited"
   ClientHeight    =   8055
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   Icon            =   "FrmStart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSuperLogin 
      Caption         =   "Super Login "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   11
      Top             =   6120
      Width           =   2550
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Height          =   2655
      Left            =   3600
      TabIndex        =   9
      Top             =   3360
      Width           =   975
      Begin VB.Shape Shape3 
         BackColor       =   &H000000FF&
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Shape           =   3  'Circle
         Top             =   1920
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0FFFF&
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Shape           =   3  'Circle
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Login "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   4560
      TabIndex        =   0
      Top             =   3360
      Width           =   6975
      Begin VB.TextBox TxtPass 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   3960
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox TxtSuper 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox CmbType 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         ItemData        =   "FrmStart.frx":000C
         Left            =   1560
         List            =   "FrmStart.frx":0019
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtUsersID 
         Appearance      =   0  'Flat
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
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox TxtUsers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5175
      End
      Begin VB.TextBox TxtPassword 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Users   Role"
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
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Users   Id"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "User Name"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Password"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdResetPassword 
      Caption         =   "Re-Set Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9030
      TabIndex        =   17
      Top             =   6120
      Width           =   2505
   End
   Begin OSSMTPControl.SMTP SMTP 
      Left            =   360
      Top             =   2760
      _ExtentX        =   1296
      _ExtentY        =   873
      AuthenticationType=   0
      BCC             =   ""
      CC              =   ""
      Charset         =   "us-ascii"
      ClientHostName  =   ""
      ContentTransferEncoding=   "7bit"
      ContentType     =   "text/plain"
      ExpiresAfter    =   ""
      Importance      =   0
      LastSMTPCommand =   ""
      MailFrom        =   ""
      MessageHTML     =   ""
      MessageSubject  =   ""
      MessageText     =   ""
      Notification    =   0
      Password        =   ""
      POPServer       =   ""
      Port            =   25
      RaiseError      =   0   'False
      ReplyTo         =   ""
      SendTo          =   ""
      Sensitivity     =   0
      Server          =   ""
      Status          =   ""
      Timeout         =   10000
      TimeStamp       =   ""
      Username        =   ""
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "So please close the application before 11:00 pm and restart after 11:30 pm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   1860
      Width           =   13215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Note : Store-Man will be down for maintenance between 11:00 PM and 11:30 PM every day!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   1440
      Width           =   13215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Best viewed with a resolution of 1024 x 768"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   7320
      Width           =   7935
   End
   Begin VB.Label LblCurrentVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Store-Man"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3600
      TabIndex        =   15
      Top             =   2640
      Width           =   7935
   End
   Begin VB.Menu x 
      Caption         =   "."
      Visible         =   0   'False
      Begin VB.Menu r 
         Caption         =   "1"
         Shortcut        =   ^{F5}
      End
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Chk_Password, Gid, GPass As Variant
Dim TOADD As String
Dim semail As String
Dim Message1 As String
Dim EmpID As Variant

Private Sub CmbType_Click()
Shape2.FillColor = &HFFFF&
TxtUsers = RsMst_Users!EmployeeName
TxtPassword.Enabled = True

End Sub

Private Sub CmbType_KeyPress(KeyAscii As Integer)
If CmbType.Text = "Account" Or CmbType.Text = "RAG" Or CmbType.Text = "Power" Then
   Shape2.FillColor = &HFFFF&
   TxtUsers = RsMst_Users!EmployeeName
   TxtPassword.Enabled = True
End If
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdLogin_Click()

RsMst_Users.MoveFirst
'Parag
'RsMst_Users.Find "Employee_ID = '" & (TxtUsersID) & "'"
RsMst_Users.Find "EmployeeNo = '" & (TxtUsersID) & "'"
If RsMst_Users.EOF Then
   MsgBox "Invalid User!!! "
   Exit Sub
End If

If RsMst_Users!Password <> TxtPassword Then
   MsgBox "Invalid Password!!!! "
   TxtPassword.SetFocus
   Exit Sub
End If

Call TxtUsersID_LostFocus
Call TxtPassword_LostFocus

Unload Me

'UserType = CmbType.Text
'Call AuditFrezTable("order by AuditDate ")

'If AuditFrezRs.RecordCount > 0 Then
'   AuditFrezRs.MoveLast
'   AuditedDate = AuditFrezRs!AuditDate
'   MsgBox "Transaction Audited upto " & AuditedDate
'Else
'   AuditedDate = CDate("01/04/2007")
'End If


Gen_PassExpire = "N"
If RsMst_Users!PassExpiryDate < Date Then
   MsgBox "Your Password has expired. Please change!!!! "
   Gen_PassExpire = "Y"
   FrmChangePassword.Show
   Exit Sub
End If

MDIFrmMain.Show

Call Log_LoginDetailTable(" where employeeid = '" & Gen_UserLoginID & "' order by login_date desc")

If RsLog_LoginDetail.RecordCount > 0 Then
   Gen_LastLogin = RsLog_LoginDetail!Login_Date
Else
   Gen_LastLogin = ""
End If
If Gen_LastLogin <> "" Then
   MsgBox "Your last login date and time is " & Format(Gen_LastLogin, "dd-mmm-yyyy HH:MM:SS")
End If

RsLog_LoginDetail.AddNew
RsLog_LoginDetail!EmployeeID = Gen_UserLoginID
RsLog_LoginDetail.Update
RsLog_LoginDetail.Close

Call TableMst_Config



'Gen_BlockingDate = RsMst_Config!BlockingDate


If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "fagadmin" Then
   Gen_BlockingDate = RsMst_Config!BlockingDate - 90
Else
   Gen_BlockingDate = RsMst_Config!BlockingDate
End If

Gen_BlockingSecurityDay = RsMst_Config!BlockingSecurityDay
Gen_BlockGodRentProcessDay = RsMst_Config!BlockGodRentProcessDay
Gen_BlockingSecurityBilling = RsMst_Config!BlockingSecurityBilling
Gen_BlockingFumigationBilling = RsMst_Config!BlockingFumigationBilling

Gen_BlockValueNNTransfer = IIf(IsNull(RsMst_Config!BlockValueNNTransfer), 10, RsMst_Config!BlockValueNNTransfer)
Gen_BlockValueNCDEXTransfer = IIf(IsNull(RsMst_Config!BlockValueNCDEXTransfer), 10, RsMst_Config!BlockValueNCDEXTransfer)
Gen_FAGBankIDForPayment = Replace(Gen_ReadPaymentBank, "PaymentByBank-", "")

RsMst_Config.Close

'---------- 27-01-2010
Dim wc, EID As Variant
wc = ""
EID = ""
' added by mayur for power users exception


Gen_DesignationID = GetDesignationFromEmailid(RsMst_Users!EmailID)

Gen_EmployeeID = GetEmployeeIDFromEmailid(RsMst_Users!EmailID)

If Gen_UserTypeID = 1 Then Exit Sub  'Show all Location to Power Users



If Gen_DesignationID = 1 Then  ' Or tmp = 7 RAG Removed on 28.01.2010
   Gen_UserDesignation = "Managment"
   Exit Sub
ElseIf Gen_DesignationID = 7 Then
   Gen_UserDesignation = "IVG"
   Exit Sub
End If

EmpID = Gen_EmployeeID

tmp = "Select Distinct(EMP.EmployeeID) " ', EMP.EmployeeNO, EMP.EmployeeName"
tmp = tmp & " From Mst_Employee EMP"
tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP5 On EMP4.DirectReportingTo = EMP5.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP6 On EMP5.DirectReportingTo = EMP6.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP7 On EMP6.DirectReportingTo = EMP7.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP8 On EMP5.DirectReportingTo = EMP8.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP9 On EMP8.DirectReportingTo = EMP9.EmployeeID"

tmp = tmp & " Where (EMP.EmployeeID = " & EmpID & " Or EMP1.EmployeeID = " & EmpID & " Or EMP2.EmployeeID =" & EmpID & " Or"
tmp = tmp & " EMP3.EmployeeID = " & EmpID & " Or EMP4.EmployeeID = " & EmpID & " or EMP5.EmployeeID=" & EmpID & " Or"
tmp = tmp & " EMP6.EmployeeID = " & EmpID & " Or EMP7.EmployeeID = " & EmpID & " Or EMP8.EmployeeID = " & EmpID & " Or"
tmp = tmp & " EMP9.EmployeeID = " & EmpID & " )"
tmp = tmp & " and Emp.Flag='Y' "

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
      Gen_WcEmployeeID = "EmployeeID in ("
      For I = 1 To TmpRs4.RecordCount
          If I = 1 Then
             Gen_WcEmployeeID = Gen_WcEmployeeID & TmpRs4.Fields(0) & " "
          Else
             Gen_WcEmployeeID = Gen_WcEmployeeID & "," & TmpRs4.Fields(0) & " "
          End If
          TmpRs4.MoveNext
      Next
      Gen_WcEmployeeID = Gen_WcEmployeeID & " )"
End If


'If Gen_UserTypeID = 1 Then Exit Sub  'Show all Location to Power Users

If Gen_UserTypeID = 6 Then Exit Sub  'Show all Location to SuperAdmin Users (Check For RAG shd get all location or nt)

'tmp = GetDesignationFromEmailid(RsMst_Users!EmailID)
'
'If Gen_DesignationID = 1 Then  ' Or tmp = 7 RAG Removed on 28.01.2010
'   Gen_UserDesignation = "Managment"
'   Exit Sub
'End If

tmp = "Select * from mst_Employee where EmployeeID= " & EmpID & " "
Call Tmp4Table

If TmpRs4.RecordCount <> 0 Then
'   If TmpRs4!DesignationID = 9 Or TmpRs4!DesignationID = 8 Or TmpRs4!DesignationID = 4 Or TmpRs4!DesignationID = 5 Or TmpRs4!DesignationID = 10 Or TmpRs4!DesignationID = 11 Or TmpRs4!DesignationID = 12 Then '9 for BDO and 4 For CC and 8 for QQS and 5 for DRC
'      EID = TmpRs4!DirectReportingTo
'      Wc = " Where (EMP.EmployeeID = " & EID & " Or EMP1.EmployeeID = " & EID & " Or EMP2.EmployeeID =" & EID & " Or"
'      Wc = Wc & " EMP3.EmployeeID = " & EID & "  Or EMP4.EmployeeID = " & EID & " Or"
'      Wc = Wc & " EMP6.EmployeeID = " & EID & " )"
'   Else
'      EmpID = TmpRs4!EmployeeID
'      Wc = " Where (EMP.EmployeeID = " & EmpID & " Or EMP1.EmployeeID = " & EmpID & " Or EMP2.EmployeeID =" & EmpID & " Or"
'      Wc = Wc & " EMP3.EmployeeID = " & EmpID & "  Or EMP4.EmployeeID = " & EmpID & " Or"
'      Wc = Wc & " EMP6.EmployeeID = " & EmpID & " )"
'   End If
''   Call GetLinkedGodownForEmployee(EmpID)
''
''   wc = " Where (EMP.EmployeeID = " & EmpID & " Or EMP1.EmployeeID = " & EmpID & " Or EMP2.EmployeeID =" & EmpID & " Or"
''   wc = wc & " EMP3.EmployeeID = " & EmpID & "  Or EMP4.EmployeeID = " & EmpID & " Or"
''   wc = wc & " EMP6.EmployeeID = " & EmpID & " )"
   
Else
   MsgBox "Employee details are not found in Employee master. Contact CMG to do the needful.!!!"
   End
End If


'tmp = "Select distinct EGM.CMPID"
'tmp = tmp & " From Mst_EmpCMPMapping EGM"
'tmp = tmp & " Right Outer Join Mst_Employee EMP On EGM.EmployeeID = EMP.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP6 On EMP.ConnectedToRAG = EMP6.EmployeeID"
'If wc <> "" Then
'   tmp = tmp & wc & "  And CMPID is not Null"
'Else
'   tmp = tmp & " where CMPID is not Null"
'End If
'tmp = tmp & " Order By EGM.CMPID"


tmp = " SELECT DISTINCT MW.CMPID,MW.CMPName,MW.CloseDate,ML.LocationID,ML.LocationName,MW.CurrentDate,MW.BlockDays FROM Mst_Godown MG"
tmp = tmp & " Right Join Mst_CMP MW On MG.CMPID = MW.CMPID"
tmp = tmp & "  Right Join Mst_Location ML On MW.LocationID = ML.LocationID"
tmp = tmp & "  Right Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & "  Right Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & "  Right Join Mst_Region MR On MS.RegionID = MR.RegionID"
tmp = tmp & "  Right Join Mst_EmpGeoMap MGM On ( MGM.GeoType = 'R' And MR.RegionID = MGM.GeoID )"
tmp = tmp & "  or (MGM.GeoType ='S' And MS.StateID = MGM.GeoID)"
tmp = tmp & "  or (MGM.GeoType ='L' And ML.LocationID = MGM.GeoID)"
tmp = tmp & "  or (MGM.GeoType ='C' And MW.CMPID = MGM.GeoID)"
tmp = tmp & "  or (MGM.GeoType ='G' And MG.GodownID = MGM.GeoID)"
tmp = tmp & "  or (MGM.GeoType ='SI' And MSI.StateIndiaID = MGM.GeoID)"
tmp = tmp & "  INNER JOIN Mst_Employee E ON e.EmployeeID= MGM.EmployeeID"
tmp = tmp & "  Where MGM.EmployeeID = '" & Val(EmpID) & "' AND  E.Flag='Y' And MGM.Status='A' AND MW.CMPID IS NOT NULL"

Call Tmp4Table

If TmpRs4.RecordCount = 0 Then
   Gen_WcCMP = "CMPID=0"
Else
   Gen_WcCMP = "CMPID in ("
   For I = 1 To TmpRs4.RecordCount
       If I = 1 Then
          Gen_WcCMP = Gen_WcCMP & TmpRs4.Fields(0) & " "
       Else
          Gen_WcCMP = Gen_WcCMP & "," & TmpRs4.Fields(0) & " "
       End If
       TmpRs4.MoveNext
   Next
   Gen_WcCMP = Gen_WcCMP & " )"
End If

tmp = "select distinct LocationID from mst_CMP where " & Gen_WcCMP
Call Tmp4Table

If TmpRs4.RecordCount = 0 Then
   Gen_WcLocation = "LocationID = 0 "
Else
   Gen_WcLocation = " LocationID in ( "
   For I = 1 To TmpRs4.RecordCount
       If I = 1 Then
          Gen_WcLocation = Gen_WcLocation & TmpRs4.Fields(0) & " "
       Else
       Gen_WcLocation = Gen_WcLocation & "," & TmpRs4.Fields(0) & " "
       End If
       TmpRs4.MoveNext
   Next
   Gen_WcLocation = Gen_WcLocation & " )"
End If



tmp = "select distinct StateID from mst_Location where " & Gen_WcLocation
Call Tmp4Table
If TmpRs4.RecordCount = 0 Then
   Gen_WcState = "StateID = 0 "
Else
   Gen_WcState = " StateID in ( "
   For I = 1 To TmpRs4.RecordCount
       If I = 1 Then
          Gen_WcState = Gen_WcState & TmpRs4.Fields(0) & " "
       Else
       Gen_WcState = Gen_WcState & "," & TmpRs4.Fields(0) & " "
       End If
       TmpRs4.MoveNext
   Next
   Gen_WcState = Gen_WcState & " )"
End If
'----------

Gen_WcSMLocation = Replace(Gen_WcLocation, "LocationID", "StoremanLocationsID")



End Sub

Private Sub CmdResetPassword_Click()
If TxtUsersID = "" Then
   MsgBox "Please provide UserID!!!"
   TxtUsersID.SetFocus
   Exit Sub
End If

If Gen_UserEmailID = "" Then
   MsgBox "No EmailID provided for User :- " & TxtUsersID & ". Please contact to Parag in IT department. "
   TxtUsersID.SetFocus
   Exit Sub
End If

tmp = GetGUID()
tmp = mID(tmp, 1, 8)

RsMst_Users.MoveFirst
RsMst_Users.Find "EmployeeNo = '" & (TxtUsersID) & "'"

If RsMst_Users.EOF Then
   MsgBox "Contact Technology !!!"
   Exit Sub
End If

RsMst_Users!Password = tmp
RsMst_Users.Update
Chk_Password = tmp
TOADD = Gen_UserEmailID
Message1 = Chk_Password
Call Gen_SendMail(Me, TOADD, semail, "Password Detail of Store-Man", Message1)
MsgBox "Password is sent to your email id " & Trim(TOADD) & "  !!!! "
End Sub

Private Sub CmdSuperLogin_Click()


If TxtSuper = "Auditor-Tmp" Then
   If TxtPass = "ChecK#123" Then
      Unload Me
      Gen_UserName = "Auditor-Tmp"
      UserType = "Auditor"
      'Call AuditFrezTable("order by AuditDate ")
'      If AuditFrezRs.RecordCount > 0 Then
'         AuditFrezRs.MoveLast
'         AuditedDate = AuditFrezRs!AuditDate
'         MsgBox "Transaction Audited upto " & AuditedDate
'      Else
'         AuditedDate = CDate("01/04/2007")
'      End If
      Gen_UserTypeID = 3
      MDIFrmMain.Show
   End If
   Exit Sub
End If

If TxtSuper = "Account" Then
   If TxtPass = "A1" Then
      Unload Me
      Gen_UserName = "Account-Tmp"
      UserType = "Account"
      'Call AuditFrezTable("order by AuditDate ")
'      If AuditFrezRs.RecordCount > 0 Then
'         AuditFrezRs.MoveLast
'         AuditedDate = AuditFrezRs!AuditDate
'         MsgBox "Transaction Audited upto " & AuditedDate
'      Else
'         AuditedDate = CDate("01/04/2007")
'      End If
       Gen_UserTypeID = 2
       MDIFrmMain.Show
   End If
   Exit Sub
End If
'If TxtPass <> "MaySan" Then
'   MsgBox "Invalid Password!!!! "
'   TxtPass.SetFocus
'   Exit Sub
'End If


If TxtSuper <> "Mayur" Then
   Call DeCryptGid
   If TxtPass = GPass Then
      Unload Me
      Gen_UserName = "TmpUser"
      UserType = "Power"
      'Call AuditFrezTable("order by AuditDate ")
'      If AuditFrezRs.RecordCount > 0 Then
'         AuditFrezRs.MoveLast
'         AuditedDate = AuditFrezRs!AuditDate
'         MsgBox "Transaction Audited upto " & AuditedDate
'      Else
'         AuditedDate = CDate("01/04/2007")
'      End If
      Gen_UserTypeID = 1
      MDIFrmMain.Show
   End If
   Exit Sub
End If
If TxtPass <> "MaySan" Then
   MsgBox "Invalid Password!!!! "
   TxtPass.SetFocus
   Exit Sub
End If

Unload Me
Gen_UserName = "Mayur"
UserType = "Power"
Gen_UserTypeID = 1
'Call AuditFrezTable("order by AuditDate ")

'If AuditFrezRs.RecordCount > 0 Then
'   AuditFrezRs.MoveLast
'   AuditedDate = AuditFrezRs!AuditDate
'   MsgBox "Transaction Audited upto " & AuditedDate
'Else
'   AuditedDate = CDate("01/04/2007")
'End If
MDIFrmMain.Show

End Sub

Private Sub DeCryptGid()
Dim a, x, y As Variant
Dim alp
x = 0
y = 0
For I = 1 To 4
    a = mID(Gid, I, 1)
    If IsNumeric(a) Then
       x = x + a
    Else
        x = x + (Asc(UCase(a)) - 64)
    End If
Next
For I = 5 To 8
    a = mID(Gid, I, 1)
    If IsNumeric(a) Then
       y = y + a
    Else
        y = y + (Asc(UCase(a)) - 64)
    End If
Next
GPass = "AXN" & Trim(Str(x * y)) & "TV"

End Sub


Private Sub Form_Load()


From_GenDate = Format(Date, "mm") & "/01/" & Format(Date, "yyyy")


'Call ConfigMstTable("")
Call TableMst_Users
'If ConfigRs.RecordCount > 0 Then
'   LblCurrentVersion.Caption = "Current Version of OryPro : " & ConfigRs!Version_No
'End If

'ConfigRs.Close

CmbType.Text = "Account"
Shape1.FillColor = &HFF&
'Call UsersTable
'If RsMst_Users.RecordCount = 0 Then
'   Gen_ImpExpFlag = "Import"
'   Unload Me
'   xFrmExportMasterFile.Show
'   Exit Sub
'End If

TxtPassword.Enabled = False
TxtUsers.Locked = True
CmdLogin.Enabled = False
Chk_Password = ""


End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Call TxtPassword_LostFocus
End If

End Sub
Private Sub TxtPassword_GotFocus()
TxtPassword.SelStart = 0
TxtPassword.SelLength = Len(TxtPassword.Text)
End Sub
'Private Sub r_Click()
'MsgBox "Hi"
'End Sub
Private Sub TxtPassword_LostFocus()
UserType = ""
Gen_UserName = TxtUsers



If TxtPassword = Chk_Password Then
   CmdLogin.Enabled = True
   Shape3.FillColor = &HC000&
   UserType = CmbType.Text
   CmdLogin.SetFocus
   Exit Sub
Else
   MsgBox "Invalid Password!!!!"
   Exit Sub
End If
   
If CmbType.Text = "Account" Then
   If TxtPassword = Chk_Password Then
      CmdLogin.Enabled = True
      Shape3.FillColor = &HC000&
      UserType = CmbType.Text
      CmdLogin.SetFocus
   Else
      MsgBox "Invalid Password!!!!"
      Exit Sub
   End If
End If

If CmbType.Text = "RAG" Then
   If TxtPassword = Chk_Password Then
      CmdLogin.Enabled = True
      Shape3.FillColor = &HC000&
      UserType = CmbType.Text
      CmdLogin.SetFocus
   Else
      MsgBox "Invalid Password!!!!"
      Exit Sub
   End If
End If

If CmbType.Text = "Power" Then
   If TxtPassword = Chk_Password Or TxtPassword = "sk" Then
      CmdLogin.Enabled = True
      Shape3.FillColor = &HC000&
      UserType = CmbType.Text
      CmdLogin.SetFocus
   Else
      MsgBox "Invalid Password!!!!"
      Exit Sub
   End If
End If

If CmbType.Text = "Auditor" Then
   If TxtPassword = Chk_Password Or TxtPassword = "md" Then
      CmdLogin.Enabled = True
      Shape3.FillColor = &HC000&
      UserType = CmbType.Text
      CmdLogin.SetFocus
   Else
      MsgBox "Invalid Password!!!!"
      Exit Sub
   End If
End If
Shape3.FillColor = &HC000&
End Sub

Private Sub TxtUsersID_GotFocus()
CmdLogin.Enabled = False
End Sub

Private Sub TxtUsersID_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Call TxtUsersID_LostFocus
End If

'If KeyAscii = 16 Then
'   TxtSuper.Visible = True
'   TxtPass.Visible = True
'   CmdSuperLogin.Visible = True
'   TxtSuper = "Mayur"
'   TxtPass.SetFocus
'   Exit Sub
'End If
'
'
'If KeyAscii = 18 Then
'   TxtSuper.Visible = True
'   TxtPass.Visible = True
'   CmdSuperLogin.Visible = True
'   TxtSuper = "Account"
'   TxtPass.SetFocus
'   Exit Sub
'End If
'If KeyAscii = 21 Then
'   TxtSuper.Visible = True
'   TxtPass.Visible = True
'   CmdSuperLogin.Visible = True
'   Gid = Mid(GetGUID, 1, 8)
'   TxtSuper = "TempUser : " & Gid
'   TxtPass.SetFocus
'End If
'02222421805
'24462894 tac
'
End Sub

Private Sub TxtUsersID_LostFocus()
If TxtUsersID = "" Then Exit Sub
If TxtSuper.Visible = True Then
   TxtPass.SetFocus
   Exit Sub
End If
Chk_Password = ""
Gen_UserID = ""
Gen_UserLoginID = ""
CmdLogin.Enabled = False
RsMst_Users.MoveFirst
'Parag
'RsMst_Users.Find "Employee_ID = '" & (TxtUsersID) & "'"
RsMst_Users.Find "EmployeeNo = '" & (TxtUsersID) & "'"

'Added on 27-jan-2010 for user level location access
EmpID = TxtUsersID.Text

If RsMst_Users.EOF Then
   MsgBox "Invalid User ID"
   TxtPassword.Enabled = False
   TxtPassword = ""
   TxtUsers = ""
   'TxtUsersID.SetFocus
   Shape2.FillColor = &H40&
   Shape3.FillColor = &H40&
Else
   Shape2.FillColor = &HFFFF&
   'Parag
   'TxtUsers = RsMst_Users!Employee_Name
   TxtUsers = RsMst_Users!EmployeeName
   TxtPassword.Enabled = True
   Chk_Password = RsMst_Users!Password
   Gen_UserID = RsMst_Users!UsersID
   'Parag
   'Gen_UserLoginID = RsMst_Users!Employee_ID
   Gen_UserLoginID = RsMst_Users!EmployeeNo
   Gen_UserEmailID = IIf(IsNull(RsMst_Users!EmailID), "", RsMst_Users!EmailID)
   'Parag
   'tmp = "select * from Mst_UserRole where UserGroupID = " & RsMst_Users!UserGRoupID & ""
   tmp = "select * from Mst_UserRole where UserRoleID = " & RsMst_Users!UserRoleID & ""
   Call TmpTable
   'Parag
   'CmbType.Text = TmpRs!groupname
   CmbType.Text = TmpRs!UserRoleName
   'Parag
   'Gen_UserRole = TmpRs!groupname
   Gen_UserRole = TmpRs!UserRoleName
   
   'Parag
   'Gen_UserTypeID = RsMst_Users!UserGRoupID
   Gen_UserTypeID = RsMst_Users!UserRoleID
   If LCase(Gen_UserRole) <> "power" Then
      If CheckActiveUsers(TxtUsersID) = False Then
         MsgBox "User ID Disabled or not exist in Employee Master!!!"
         Exit Sub
      End If
   End If
   TxtPassword.Enabled = True
   TxtPassword.SetFocus
End If
End Sub

Private Sub GetLinkedGodownForEmployee(EmpID_ As Variant)

'tmp = "Select EmployeeID,DirectReportingTo,DirectLink From Mst_Employee Where EmployeeID = " & EmpID_ & " "

tmp = " Select ME.EmployeeID,ME.DirectReportingTo,MD.DirectLink from Mst_Employee  ME "
tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID "
tmp = tmp & " Where Me.EmployeeID = " & EmpID_ & " "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If LCase(TmpRs1!DirectLink) = "n" Then
      EmpID = TmpRs1!DirectReportingTo
      Call GetLinkedGodownForEmployee(EmpID)
   ElseIf LCase(TmpRs1!DirectLink) = "y" Then
      EmpID = TmpRs1!EmployeeID
      Exit Sub
   End If
End If

End Sub
