VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_StateWiseGSTNforNavision 
   Caption         =   "Pending Client List for Navision"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   7905
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
      TabIndex        =   5
      Top             =   120
      Width           =   15015
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   225
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
         Height          =   375
         Left            =   13845
         TabIndex        =   3
         Top             =   225
         Width           =   1095
      End
      Begin VB.ComboBox CmbExchangeType 
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
         ItemData        =   "FrmRpt_StateWiseGSTNforNavision.frx":0000
         Left            =   1320
         List            =   "FrmRpt_StateWiseGSTNforNavision.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   263
         Width           =   2835
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
         TabIndex        =   8
         Top             =   285
         Width           =   60
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
         TabIndex        =   7
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   165
         TabIndex        =   6
         Top             =   300
         Width           =   1065
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   9240
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8295
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14631
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_StateWiseGSTNforNavision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mR, mC As Double
Dim mExchangeTypeID As Variant
Dim mExcelFilePath As Variant
Dim mListOfEmailEmployeeNo As Variant

Private Sub CmbExchangeType_GotFocus()
tmp = CmbExchangeType.Text
mTmp = CmbExchangeType.Text
Call TableMst_ExchangeType(" Where Flag = 'A' And ExchangeTypeID In (1,2) ")
CmbExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Exchange Type found"
   CmbExchangeType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeType.Text = tmp
End Sub
Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid Exchange Type selection "
   CmbExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType.Fields(0)
End Sub

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then Exit Sub
If LCase(CmbType.Text) <> LCase("Client") And LCase(CmbType.Text) <> LCase("Depository") Then
   MsgBox "Invalid selection!!!"
   CmbType.SetFocus
   Exit Sub
End If
CmbExchangeType.Text = ""
LblTotalRecord.Caption = ""
Call GridHead
If LCase(CmbType.Text) = LCase("Client") Then
    Label2.Visible = True
    CmbExchangeType.Visible = True
ElseIf LCase(CmbType.Text) = LCase("Depository") Then
    Label2.Visible = False
    CmbExchangeType.Visible = False
End If
End Sub

Private Sub CmdClear_Click()
Call ClearData
End Sub

Function ClientPanNoFromHRPro(CRFID_ As Double) As Variant
Dim mResult As Variant

mResult = ""
tmp = "Select CRFID,PAN,FirmName From Txn_CRF Where CRFID = " & CRFID_ & ""
Call TmpTableHRPro
If TmpRsHRPro.RecordCount > 0 Then
    mResult = UCase(IIf(IsNull(TmpRsHRPro!PAN), "", TmpRsHRPro!PAN))
End If

ClientPanNoFromHRPro = mResult
End Function

Private Sub CmdGetReport_Click()
Dim I As Double
If CmbExchangeType.Text = "" Then
    MsgBox "Select Exchange."
    CmbExchangeType.SetFocus
    Exit Sub
End If
ProgressBar1.Value = 0
CmdExcel.Enabled = False
Call GridHead
LblTotalRecord.Caption = ""
MsgBox "Wait till next Message."

''-- Client GST Data
tmp = " SELECT GSL.ClientID,MC.ClientIDRow,MC.ClientName,MC.CRFID,MET.ExchangeType,MSI.StateIndiaID,MSI.StateIndiaCode,MSI.StateIndiaName"
tmp = tmp & " FROM Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Location ML ON GSL.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_Client MC On GSL.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Where MC.Status = 'A' And GSL.BalanceBags > 0 And MET.ExchangeTypeID = " & mExchangeTypeID & ""
tmp = tmp & " And Not Exists ( SELECT A.ClientID,A.StateIndiaID FROM Mst_ClientStateWiseGST AS A  Where A.GSTNavCode+''+A.RGSTNavCode <> '' And GSL.ClientID = A.ClientID And MSI.StateIndiaID = A.StateIndiaID )"
tmp = tmp & " Group By GSL.ClientID,MC.ClientIDRow,MC.ClientName,MC.CRFID,MET.ExchangeType,MSI.StateIndiaID,MSI.StateIndiaCode,MSI.StateIndiaName"
tmp = tmp & "     Union "
tmp = tmp & " Select MSC.ClientID,MC.ClientIDRow,MC.ClientName,MC.CRFID,MET.ExchangeType,MSI.StateIndiaID,MSI.StateIndiaCode,MSI.StateIndiaName"
tmp = tmp & " From Mst_StorageContract MSC"
tmp = tmp & " Inner Join Mst_Client MC On MSC.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_Location ML On MSC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
tmp = tmp & " Where MSC.Flag = 'A' And MET.ExchangeTypeID = " & mExchangeTypeID & ""
tmp = tmp & " And Not Exists (SELECT A.ClientID,A.StateIndiaID FROM Mst_ClientStateWiseGST AS A  Where A.GSTNavCode+''+A.RGSTNavCode <> '' And MSC.ClientID = A.ClientID And MSI.StateIndiaID = A.StateIndiaID )"
tmp = tmp & " Group By MSC.ClientID,MC.ClientIDRow,MC.ClientName,MC.CRFID,MET.ExchangeType,MSI.StateIndiaID,MSI.StateIndiaCode,MSI.StateIndiaName"
tmp = tmp & "     Union "
tmp = tmp & " Select GST.ClientID,MC.ClientIDRow,MC.ClientName,MC.CRFID,MET.ExchangeType,MSI.StateIndiaID,MSI.StateIndiaCode,MSI.StateIndiaName"
tmp = tmp & " From Mst_ClientStateWiseGST GST"
tmp = tmp & " Inner Join Mst_Client MC On GST.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_StateIndia MSI On GST.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
tmp = tmp & " Where MC.Status = 'A' And MET.ExchangeTypeID = " & mExchangeTypeID & ""
tmp = tmp & " And Exists (SELECT A.ClientID,A.StateIndiaID FROM Mst_ClientStateWiseGST AS A  Where A.GSTNavCode+''+A.RGSTNavCode = '' And GST.ClientID = A.ClientID And MSI.StateIndiaID = A.StateIndiaID )"
tmp = tmp & " Group By GST.ClientID,MC.ClientIDRow,MC.ClientName,MC.CRFID,MET.ExchangeType,MSI.StateIndiaID,MSI.StateIndiaCode,MSI.StateIndiaName"
tmp = tmp & " Order By MC.ClientName,MSI.StateIndiaName"
Call TmpTable
With MSHFlexGrid1
    If TmpRs.RecordCount > 0 Then
        ProgressBar1.Max = TmpRs.RecordCount
        For I = 1 To TmpRs.RecordCount
            .TextMatrix(I, 0) = ""
            .TextMatrix(I, 1) = TmpRs!ClientName
            .TextMatrix(I, 2) = ""
            .TextMatrix(I, 3) = ""
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!StateIndiaCode), "", TmpRs!StateIndiaCode)
            .TextMatrix(I, 5) = ""
            .TextMatrix(I, 6) = ""
            .TextMatrix(I, 7) = ""
            .TextMatrix(I, 8) = ClientPanNoFromHRPro(IIf(IsNull(TmpRs!CRFID), 0, TmpRs!CRFID))
            .TextMatrix(I, 9) = IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow)
            .TextMatrix(I, 10) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
            .Rows = .Rows + 1
            TmpRs.MoveNext
            ProgressBar1.Value = I
        Next
    End If
End With

LblTotalRecord.Caption = TmpRs.RecordCount
MsgBox "Done!!!"
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
mListOfEmailEmployeeNo = "'3432','3699','3601','3659','3644','3753','3423'"
''-- EMail List Provided By Manu Agrawal
'1.  Girdhari
'2.  Pawan
'3.  Krishna
'4.  Rakesh Raut
'5.  Rajesh Yadav
'6.  Saddam
'7.  Nagina
Frame1.Enabled = True
Call CmdClear_Click
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_PendingListforNavision.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_PendingListforNavision.xls")
Set oWS = oWB.Worksheets("Sheet1")
mExcelFilePath = Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_PendingListforNavision.xls"
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
           If C <> 9 Then
            oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           Else
            oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
If MSHFlexGrid1.Rows > 2 Then
    Call Gen_Email
End If
oXL.Visible = True
End Sub

Private Sub ClearData()
Call GridHead
LblTotalRecord.Caption = ""
End Sub

Private Sub GridHead()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 11
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Nav Code":                 .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "Name":                     .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "Address":                  .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "Segment":                  .ColWidth(3) = 1200
    .TextMatrix(0, 4) = "State Code":               .ColWidth(4) = 1000
    .TextMatrix(0, 5) = "GST Registration No.":     .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "GSTIN":                    .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "GST Customer Type":        .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "Pan nuumber":              .ColWidth(8) = 1500
    .TextMatrix(0, 9) = "Client ID":                .ColWidth(9) = 1500
    .TextMatrix(0, 10) = "Type":                    .ColWidth(10) = 1800
    .RowHeight(0) = 600
    .Font.Size = 10
    .Refresh
End With
End Sub


Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg As Variant
Dim mR As Variant

mMsg = ""
TOADD = ""
TOCC = ""
mCMPMEmailList = ""

mMsg = "Please find the attachment for Pending List client which are not register in Navision."

''mMsg = "<table border=1 width='80%' align='center'>"
''mMsg = mMsg & "<tr><td width='80%'><b>Agreement ID</b></td><td width='60%'>" & TxtAgreementID.Text & "</td></tr>"
''mMsg = mMsg & "</table>"

tmp = " Select EmployeeNo,EmployeeName,EmployeeEmailID from Mst_Employee Where EmployeeNo In (" & mListOfEmailEmployeeNo & ") ANd Isnull(Flag,'Y') = 'Y' "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   For mR = 1 To TmpRs.RecordCount
    TOADD = TOADD & "," & TmpRs!EmployeeEmailID
    TmpRs.MoveNext
   Next mR
   TOADD = Right(TOADD, Len(TOADD) - 1)
End If

If mForwardID <> Gen_UserID Then
End If

'TOADD = ""  ''"ullhas.p@ncml.com" '' "cmg@ncml.com;accounts@ncml.com"
TOCC = Gen_UserEmailID & ",manu.ag@ncml.com"
mMsgSubject = "Pending Client List for Navision"
Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "N")
End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant, mMailTypeFlg As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession
servertxt = "webmail.ncml.com"
htmlStringhead = "<html>"
htmlStringhead = htmlStringhead & " <head>"
htmlStringhead = htmlStringhead & "<style> body{ color: #000000;margin: 5px;font-family: Candara;} table, th, td { border: 1px solid black; border-collapse: collapse;} th { padding: 5px; text-align: center;} td { padding: 5px; text-align: left;} </style>"
htmlStringhead = htmlStringhead & "</head><body>"
htmlStringtail = "</body></html>"
   
''With oSMTP
''  'connection
''  .Server = servertxt
''  If LCase(mMailTypeFlg) = "n" Then
''     .MailFrom = "info@ncmsl.com"
''  Else
''     .MailFrom = Gen_UserEmailID
''     '.MailFrom = "ullhas.p@ncml.com"
''  End If
'''    .SendTo = "ullhas.p@ncml.com"
'''    .CC = Gen_UserEmailID
''  .SendTo = ToADD_
''  If ToADD_ <> TOCC_ Then
''     .CC = TOCC_
''  End If
''  If LCase(mMailTypeFlg) = "n" Then
''     '.Attachments.Add (mPDFFileName)
''  End If
''  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
''  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
''  .TimeStamp = Now()
''  If LCase(Gen_EmailConfigFlag) = "y" Then
''     .SendEmail
''  End If
''End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
'If LCase(mMailTypeFlg) = "n" Then
'   objSMTP.FromAddr = "info@ncmsl.com"
'Else
   objSMTP.FromAddr = Gen_UserEmailID
'End If
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If
If LCase(mMailTypeFlg) = "n" Then
   'objSMTP.AddAttachment mPDFFileName
   objSMTP.AddAttachment mExcelFilePath
End If
objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If
End Function

