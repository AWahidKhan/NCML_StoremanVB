VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_CMPMonthEndProcess 
   Caption         =   "CMP Month End Process"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   12675
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   11535
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   885
      Width           =   12675
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   12420
         _ExtentX        =   21908
         _ExtentY        =   9340
         _Version        =   393216
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   6555
      Width           =   12735
      Begin VB.CommandButton ExitCmd 
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
         Height          =   495
         Left            =   11325
         TabIndex        =   12
         Top             =   1230
         Width           =   1335
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
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
         Left            =   11325
         TabIndex        =   11
         Top             =   240
         Width           =   1335
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
         Left            =   11325
         TabIndex        =   10
         Top             =   735
         Width           =   1335
      End
      Begin VB.CheckBox ChkSelectAll 
         Caption         =   "Select All"
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
         Left            =   9960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   $"FrmTxn_CMPMonthEndProcess.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   10215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Label2"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Since BOD date is less than month end date, Month End Process cannot be carried out."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   7935
      End
   End
End
Attribute VB_Name = "FrmTxn_CMPMonthEndProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLOcationID, r, mNextMonth As Variant

Private Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 8
MSHFlexGrid2.TextMatrix(0, 0) = "Selection "
MSHFlexGrid2.TextMatrix(0, 1) = "EHOP CMP ID "
MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 3) = "CMP Locked Date"
MSHFlexGrid2.TextMatrix(0, 4) = "CMPID"
MSHFlexGrid2.TextMatrix(0, 5) = "Transaction Completed Upto Date"
MSHFlexGrid2.TextMatrix(0, 6) = "Old CMP Locked Date"
MSHFlexGrid2.TextMatrix(0, 7) = "New CMP Locked Date"

MSHFlexGrid2.ColWidth(0) = 800
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 3000
MSHFlexGrid2.ColWidth(3) = 3000
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 3000
MSHFlexGrid2.ColWidth(6) = 0
MSHFlexGrid2.ColWidth(7) = 0

MSHFlexGrid2.RowHeight(0) = 500
MSHFlexGrid2.WordWrap = True

End Sub

Private Sub ChkSelectAll_Click()
If MSHFlexGrid2.Rows = 2 Then Exit Sub
For r = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(r, 2) = "" Then Exit For
    If ChkSelectAll.Value = 1 Then
       If MSHFlexGrid2.TextMatrix(r, 0) = "" Then
          MSHFlexGrid2.row = r
          MSHFlexGrid2.Col = 0
          If MSHFlexGrid2.CellBackColor <> vbYellow Then
             MSHFlexGrid2.TextMatrix(r, 0) = "*"
          End If
       End If
    Else
       MSHFlexGrid2.TextMatrix(r, 0) = ""
    End If
Next

End Sub

Private Function ChekGrid() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid2.Rows - 2
    If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
       Retval = True
       Exit For
    End If
Next
ChekGrid = Retval
End Function



Private Sub CmdClear_Click()
Call Grid_Head
CmbLocationID.Text = ""
ChkSelectAll.Value = 0
CmbLocationID.SetFocus
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Exit Sub
End If

Call Grid_Head

CmbLocationID = ""

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
r = 0
End Sub
Private Sub MSHFlexGrid2_Click()

r = MSHFlexGrid2.RowSel

If r <= 0 Then Exit Sub
If MSHFlexGrid2.CellBackColor = vbYellow Then Exit Sub

If MSHFlexGrid2.TextMatrix(r, 3) <> "" Then
   If MSHFlexGrid2.TextMatrix(r, 0) = "" Then
      MSHFlexGrid2.TextMatrix(r, 0) = "*"
   Else
      MSHFlexGrid2.TextMatrix(r, 0) = ""
      If ChkSelectAll.Value = 1 Then
         ChkSelectAll.Value = 0
      End If
   End If
End If

End Sub
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
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
   mLOcationID = Null
   Call Grid_Head
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLOcationID <> RsMst_Location!LocationID Then
   Call Grid_Head
   
End If
mLOcationID = RsMst_Location!LocationID
ChkSelectAll.Value = 0
Call FillGrid(mLOcationID)

End Sub
Private Sub FillGrid(mLocationID_)
Dim k As Variant
Call Grid_Head

Call TableMst_CMP(" where CloseDate is null And LocationID = " & mLocationID_ & " ")

If RsMst_CMP.RecordCount > 0 Then
   RsMst_CMP.MoveFirst
   MSHFlexGrid2.Rows = RsMst_CMP.RecordCount + 2
   k = 1
   For I = 0 To RsMst_CMP.RecordCount - 1
       MSHFlexGrid2.TextMatrix(k, 0) = ""
       MSHFlexGrid2.TextMatrix(k, 1) = RsMst_CMP!EHOPCMPID
       MSHFlexGrid2.TextMatrix(k, 2) = RsMst_CMP!CMPName
       MSHFlexGrid2.TextMatrix(k, 3) = Format(IIf(IsNull(RsMst_CMP!CMPLockedDate), "", RsMst_CMP!CMPLockedDate), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(k, 4) = RsMst_CMP!CMPID
       MSHFlexGrid2.TextMatrix(k, 5) = Format(IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(k, 6) = Format(IIf(IsNull(RsMst_CMP!CMPLockedDate), "", RsMst_CMP!CMPLockedDate), "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(k, 7) = ""
       'mNextMonth = Format(Year(CDate(MSHFlexGrid2.TextMatrix(k, 3))) & "-" & Month(CDate(MSHFlexGrid2.TextMatrix(k, 3))) + 1 & "-" & "01", "dd-mmm-yyyy")
       
       'Check BOD  (Current Date) is last date of month or not
       If CDate(MSHFlexGrid2.TextMatrix(k, 5)) >= DateAdd("d", -1, CDate(RsMst_CMP!CMPLockedDate)) Then
          'MSHFlexGrid2.TextMatrix(k, 3) = mNextMonth
       Else
          'If CDate(MSHFlexGrid2.TextMatrix(k, 3)) <> CDate(MSHFlexGrid2.TextMatrix(k, 5)) Then
          MSHFlexGrid2.row = k
          For x = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.Col = x
              MSHFlexGrid2.CellBackColor = vbYellow
          Next
          'End If
       End If
       
       
       k = k + 1
       RsMst_CMP.MoveNext
   Next
Else
   MsgBox "No CMP found for selected location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
'RsMst_CMP.Filter = ""




End Sub

Private Sub SaveCmd_Click()
Dim k, TmpCMPID, mStartDate As Variant

If CmbLocationID.Text = "" Then
   MsgBox "Please select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If ChekGrid = False Then
   MsgBox "Please Select CMP From Grid !!!"
   Exit Sub
End If

If MSHFlexGrid2.Rows <= 2 Then Exit Sub
k = 1
Call TableMst_DR

For I = 0 To RsMst_CMP.RecordCount - 1
    If MSHFlexGrid2.TextMatrix(k, 1) = "" Then Exit For
    
    If MSHFlexGrid2.TextMatrix(k, 0) = "*" Then
       RsMst_CMP.MoveFirst
       RsMst_CMP.Find "CMPID  = " & Val(MSHFlexGrid2.TextMatrix(k, 4))
       If Not RsMst_CMP.EOF Then
          mStartDate = CDate(MSHFlexGrid2.TextMatrix(k, 3))
          RsMst_CMP!CMPLockedDate = DateAdd("m", 1, CDate(MSHFlexGrid2.TextMatrix(k, 3))) 'MSHFlexGrid2.TextMatrix(k, 3)
          RsMst_CMP.Update
          Call SaveInvMonGSLDetail(Val(MSHFlexGrid2.TextMatrix(k, 4)), mStartDate)
          MSHFlexGrid2.TextMatrix(k, 7) = RsMst_CMP!CMPLockedDate
       End If
    End If
    k = k + 1
Next
MsgBox "Data saved successfully!!! "


Call Gen_Mail

Call Grid_Head

End Sub
Private Sub SaveInvMonGSLDetail(TmpCMPID_ As Variant, mStartDate_ As Variant)
Call TableTxn_InvMonGSLDetail
'Call TableMst_GSL(" Where CMPID = " & TmpCMPID_ & " And Flag in('O','J','R') And SpillageFlag=0")

tmp = "Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,"
tmp = tmp & " BalanceBags,BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,"
tmp = tmp & " Flag1 , GrossWt, KRNo, KRNoBatchID, KRDate,NContractID, round((BalanceWeight/Nullif(BalanceBags,0))*1000,2) 'BagSize'"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " Where CMPID = " & TmpCMPID_ & " And Flag in('O','J','R') And SpillageFlag=0"
'tmp = tmp & " And ExchangeTypeID=1"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For mRow = 1 To TmpRs.RecordCount
       RsTxn_InvMonGSLDetail.AddNew
       For mCol = 0 To TmpRs.Fields.Count - 1
           RsTxn_InvMonGSLDetail.Fields(mCol) = TmpRs.Fields(mCol)
       Next
       RsTxn_InvMonGSLDetail!tFlag = "O"
            
       If RsMst_DR.RecordCount > 0 Then
          RsMst_DR.MoveFirst
       End If
       
       RsMst_DR.Find "DRNO =" & RsTxn_InvMonGSLDetail!KRNo & ""
       If Not RsMst_DR.EOF Then
          RsTxn_InvMonGSLDetail!DRNo = RsTxn_InvMonGSLDetail!KRNo
          RsTxn_InvMonGSLDetail!DRDate = RsMst_DR!DRDate
       End If
       
       RsTxn_InvMonGSLDetail!G_UID = GetGUID
       RsTxn_InvMonGSLDetail!YearMonth = mStartDate_
       RsTxn_InvMonGSLDetail!CreatedBy = Gen_UserLoginID
       RsTxn_InvMonGSLDetail!CreatedDate = Now
       RsTxn_InvMonGSLDetail.Update
       TmpRs.MoveNext
   Next
End If
End Sub
Private Function Gen_Mail()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg, mMsgSubject As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsgSubject = "CMP Loked Date Update Details for " & CmbLocationID.Text

mMsg = ""
mMsg = "<tr><td> Location </td><td> " & CmbLocationID.Text & "</td></tr>"

mMsg = mMsg & "<tr><td> Sr.No. </td><td> CMP Name </td><td> Previous CMP Locked Date </td><td> Present CMP Locked Date </td></tr>"

Mcount = 1
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then 'If MSHFlexGrid2.TextMatrix(i, 3) <> MSHFlexGrid2.TextMatrix(i, 6) Then
       mGridMsg = mGridMsg & "<tr><td> " & Mcount & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td>" & Format(MSHFlexGrid2.TextMatrix(I, 6), "dd-mmm-yyyy") & "</td><td>" & Format(MSHFlexGrid2.TextMatrix(I, 3), "dd-mmm-yyyy") & "</td></tr>"
       Mcount = Mcount + 1
    End If
Next

mMsg = mMsg & mGridMsg

TOADD = "cmg@ncml.com" ' mEmployeeEmailID

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)

End Function

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = Gen_UserEmailID  'TOCC_ '"mayur.d@ncml.com"
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "parag.r@ncml.com" '"bhagwat.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  'TOCC_ = TOCC_ '& "," & mEmployeeEmailID
''  If ToADD_ <> TOCC_ Then
''     .CC = TOCC_
''  End If
'  '.BCC = "mayur.d@ncml.com"
'  .MessageSubject = MsgSubject_ '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
''  .SendEmail
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "parag.r@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ '& " (from HR-Pro)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If


End Function



