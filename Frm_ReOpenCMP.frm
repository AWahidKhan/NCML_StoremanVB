VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ReOpenCMP 
   Caption         =   "Reopen CMP"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8025
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   15045
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   14775
         Begin VB.TextBox TxtCMPName 
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
            Height          =   360
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   11
            Top             =   300
            Width           =   5160
         End
         Begin VB.CommandButton CmdAdd2List 
            Caption         =   "Update Date"
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
            Left            =   11520
            TabIndex        =   10
            Top             =   300
            Width           =   1695
         End
         Begin VB.CheckBox Check2 
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
            Left            =   6720
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker TxtReOpenDate 
            Height          =   390
            Left            =   9960
            TabIndex        =   12
            Top             =   285
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   110690305
            CurrentDate     =   39884
         End
         Begin VB.Label LblCMPName 
            Caption         =   "CMP Name"
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
            TabIndex        =   14
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Re-Opened Date"
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
            Left            =   8040
            TabIndex        =   13
            Top             =   360
            Width           =   1770
         End
      End
      Begin VB.ComboBox CmbLocation 
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
         TabIndex        =   4
         Top             =   240
         Width           =   5160
      End
      Begin VB.Frame Frame3 
         Caption         =   "CMP Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   14820
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   5955
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   10504
            _Version        =   393216
            BackColor       =   16777215
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
      Begin VB.Label LbLocation 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   8040
      Width           =   15060
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   13320
         TabIndex        =   2
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton CmdReopen 
         Caption         =   "&Reopen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11925
         TabIndex        =   1
         Top             =   240
         Width           =   1400
      End
   End
End
Attribute VB_Name = "Frm_ReOpenCMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID As Integer
Dim mCMPID As Integer

Private Sub Check2_Click()
For i = 1 To MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.TextMatrix(i, 1) = ""
Next
If Check2.Value = 1 Then
   For i = 1 To MSHFlexGrid2.Rows - 2
       If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
          MSHFlexGrid2.TextMatrix(i, 1) = "*"
       End If
   Next
End If
End Sub

Private Sub CmbLocation_GotFocus()
'Call GridHead1
tmp1 = CmbLocation.Text
CmdReopen.Enabled = False
tmp = "select Distinct Mst_CMP.locationID,LocationName from Mst_CMP" & _
      " inner join Mst_Location on Mst_CMP.locationID=Mst_Location.LocationID" & _
      " Where CloseDate Is Not Null"
Call Tmp0Table
CmbLocation.Clear
If TmpRs0.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To TmpRs0.RecordCount
    CmbLocation.AddItem TmpRs0!LocationName
    TmpRs0.MoveNext
Next
CmbLocation.Text = tmp1
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Exit Sub
End If
TmpRs0.MoveFirst
TmpRs0.Find "LocationName = '" & CmbLocation.Text & "'"
If TmpRs0.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = TmpRs0!LocationID
Call FillCMP(mLocationID)
End Sub
Private Sub FillCMP(LocationID As Integer)
Call TableMst_CMP(" where LocationID=" & LocationID & " And CloseDate is not null ")

Call GridHead1

For i = 1 To RsMst_CMP.RecordCount
    
    MSHFlexGrid2.TextMatrix(i, 0) = RsMst_CMP!CMPID
    MSHFlexGrid2.TextMatrix(i, 2) = RsMst_CMP!CMPName
    MSHFlexGrid2.TextMatrix(i, 3) = RsMst_CMP!Address
    MSHFlexGrid2.TextMatrix(i, 4) = IIf(IsNull(RsMst_CMP!TelephoneNo), "", RsMst_CMP!TelephoneNo)
    MSHFlexGrid2.TextMatrix(i, 5) = IIf(IsNull(RsMst_CMP!StartDate), "", RsMst_CMP!StartDate)
    MSHFlexGrid2.TextMatrix(i, 6) = IIf(IsNull(RsMst_CMP!CloseDate), "", RsMst_CMP!CloseDate)
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    RsMst_CMP.MoveNext
Next

If LCase(Gen_UserRole) = "power" Then
   CmdReopen.Enabled = True
Else
   CmdReopen.Enabled = False
End If

End Sub

Private Sub CmdAdd2List_Click()

If CmbLocation.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

If TxtCMPName.Text = "" Then
   MsgBox "No CMP selected from list!!!"
   Exit Sub
End If

If TxtReOpenDate.Value > Date Then
   MsgBox "Future date is not allowed!!!"
   TxtReOpenDate.SetFocus
   Exit Sub
End If

For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
          If TxtReOpenDate.Value < CDate(MSHFlexGrid2.TextMatrix(i, 6)) Then
             MsgBox "ReOpen Date Must Be Geater than Or Equal to CloseDate!!!"
             Exit Sub
          End If
       End If
       
    End If
Next

If Check2.Value = 1 Then
   For i = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
          MSHFlexGrid2.TextMatrix(i, 7) = TxtReOpenDate.Value
       End If
   Next
ElseIf TxtCMPName.Text <> "" Then
   For i = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
          If MSHFlexGrid2.TextMatrix(i, 0) = mCMPID Then
             MSHFlexGrid2.TextMatrix(i, 7) = TxtReOpenDate.Value
          End If
       End If
   Next
End If

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdReopen_Click()

If CmbLocation.Text = "" Then
   MsgBox "Blank Location Not Allowed."
   CmbLocation.SetFocus
   Exit Sub
End If

If CheckSelCMP = False Then
   MsgBox "Please Select Atleast one CMP."
   Exit Sub
End If
If CheckReopenDate = False Then
   MsgBox "Please Update Reopen Date."
   Exit Sub
End If

For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
          RsMst_CMP.MoveFirst
          RsMst_CMP.Find "CMPID=" & Val(MSHFlexGrid2.TextMatrix(i, 0)) & ""
          RsMst_CMP!CMPID = Val(MSHFlexGrid2.TextMatrix(i, 0))
          RsMst_CMP!CloseDate = Null
          RsMst_CMP.Update
          Call TableLog_CMPReopenHistory
          RsLog_CMPReopenHistory.AddNew
          RsLog_CMPReopenHistory!CMPID = Val(MSHFlexGrid2.TextMatrix(i, 0))
          RsLog_CMPReopenHistory!OldCloseDate = MSHFlexGrid2.TextMatrix(i, 6)
          RsLog_CMPReopenHistory!ReOpenedDate = MSHFlexGrid2.TextMatrix(i, 7)
          RsLog_CMPReopenHistory!UpdatedBy = Gen_UserLoginID
          RsLog_CMPReopenHistory.Update
       End If
    End If
Next
Call Gen_Mail
MsgBox "CMP ReOpen Successfully."
CmbLocation.Text = ""
Call GridHead1
CmdReopen.Enabled = False

End Sub
Private Function CheckSelCMP() As Boolean
Dim Retval As Boolean
Retval = False
For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 1) = "*" Then
          Retval = True
          Exit For
       End If
    End If
Next
CheckSelCMP = Retval
End Function
Private Function CheckReopenDate() As Boolean
Dim Retval As Boolean
Retval = True
For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 1) = "*" Then
          If MSHFlexGrid2.TextMatrix(i, 7) = "" Then
             Retval = False
             Exit For
          End If
       End If
    End If
Next
CheckReopenDate = Retval
End Function
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   ButtonFrm.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
ButtonFrm.Enabled = True
Call GridHead1
CmdReopen.Enabled = False
Check2.Value = False
End Sub
Private Sub GridHead1()

MSHFlexGrid2.Clear
MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 8
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.FixedCols = 3
MSHFlexGrid2.TextMatrix(0, 0) = "CMP ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Select"
MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 3) = "Address"
MSHFlexGrid2.TextMatrix(0, 4) = "Telephone"
MSHFlexGrid2.TextMatrix(0, 5) = "Start Date"
MSHFlexGrid2.TextMatrix(0, 6) = "Close Date"
MSHFlexGrid2.TextMatrix(0, 7) = "ReOpened Date"
MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 700
MSHFlexGrid2.ColWidth(2) = 2400
MSHFlexGrid2.ColWidth(3) = 4600
MSHFlexGrid2.ColWidth(4) = 2000
MSHFlexGrid2.ColWidth(5) = 1400
MSHFlexGrid2.ColWidth(6) = 1400
MSHFlexGrid2.ColWidth(7) = 1800
MSHFlexGrid2.RowHeight(0) = 400

End Sub

Private Sub MSHFlexGrid2_Click()
Dim i As Variant
i = MSHFlexGrid2.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(i, 2) = "" Then Exit Sub

If MSHFlexGrid2.TextMatrix(i, 1) = "*" Then
   MSHFlexGrid2.TextMatrix(i, 1) = ""
   TxtCMPName.Text = ""
   MSHFlexGrid2.TextMatrix(i, 7) = ""
   mCMPID = 0 'Null
Else
   MSHFlexGrid2.TextMatrix(i, 1) = "*"
   TxtCMPName.Text = MSHFlexGrid2.TextMatrix(i, 2)
   mCMPID = MSHFlexGrid2.TextMatrix(i, 0)
End If

End Sub
Private Function Gen_Mail()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg, mMsgSubject As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsgSubject = "CMP Re-Open Details"

mMsg = ""
mMsg = "<tr><td colspan=6>CMP Re-Open Details</td></tr>"
mMsg = mMsg & "<tr><td> Location </td><td colspan=5> " & CmbLocation.Text & "</td></tr>"

mMsg = mMsg & "<tr><td> Sr.No. </td><td> CMP ID </td><td> CMP Name </td><td> Old CloseDate </td><td> Re-Open Date </td><td> Re-OpenedBy Emp No </td></tr>"

For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 2) <> "" Then
       If MSHFlexGrid2.TextMatrix(i, 1) = "*" Then
          mGridMsg = mGridMsg & "<tr><td> " & i & " </td><td> " & MSHFlexGrid2.TextMatrix(i, 0) & " </td><td> " & MSHFlexGrid2.TextMatrix(i, 2) & " </td><td>" & Format(MSHFlexGrid2.TextMatrix(i, 6), "dd-MMM-yyyy") & "</td><td>" & Format(MSHFlexGrid2.TextMatrix(i, 7), "dd-MMM-yyyy") & "</td><td> " & Gen_UserLoginID & "  </td></tr>"
       End If
    End If
Next

mMsg = mMsg & mGridMsg

TOADD = "cmg@ncml.com" ' mEmployeeEmailID
TOCC = "anilkumar.par@NCML.COM" 'Gen_UserEmailID


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
'  .MailFrom = Gen_UserEmailID '"sagar.k@ncml.com" 'TOCC_ '"mayur.d@ncml.com"
''  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
''     ToADD_ = "anil.kumar@ncml.com" '"bhagwat.k@ncml.com"
''  End If
'  .SendTo = ToADD_
'  'TOCC_ = TOCC_ '& "," & mEmployeeEmailID
'  'If ToADD_ <> TOCC_ Then
'  .CC = TOCC_
'  'End If
'  .BCC = "parag.r@ncml.com"
'  .MessageSubject = MsgSubject_ '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  If Day(Date) > "12" Then
'     .TimeStamp = Now()  'Format(Now(), "mm/dd/yyyy hh:mm:ss")
'  End If
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
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

