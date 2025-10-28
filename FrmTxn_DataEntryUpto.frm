VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_DataEntryUpto 
   Caption         =   "Data Entry Upto Which Date Completed"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   8880
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8535
      Left            =   14760
      TabIndex        =   17
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   8820
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   12315
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
            TabIndex        =   1
            Top             =   240
            Width           =   11175
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
            TabIndex        =   16
            Top             =   270
            Width           =   1440
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5535
         Left            =   120
         TabIndex        =   13
         Top             =   885
         Width           =   12375
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   5175
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   12180
            _ExtentX        =   21484
            _ExtentY        =   9128
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   7635
         Width           =   12450
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
            Left            =   10560
            TabIndex        =   5
            Top             =   255
            Width           =   1695
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
            Left            =   8865
            TabIndex        =   4
            Top             =   255
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Yellow colour indicates pending Printing of CDTF / CITF. Such records will not be saved !!!!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   8895
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   6555
         Width           =   12375
         Begin VB.TextBox TxtCMPName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   0
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   8
            Top             =   555
            Visible         =   0   'False
            Width           =   4680
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
            Left            =   10560
            TabIndex        =   3
            Top             =   555
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
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
            Left            =   5160
            TabIndex        =   7
            Top             =   630
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DataUpto 
            Height          =   390
            Left            =   7680
            TabIndex        =   2
            Top             =   555
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
            Format          =   48824321
            CurrentDate     =   39884
         End
         Begin VB.Label LblCMPName 
            Caption         =   "CMP Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Date Upto Which Data entry completed"
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
            Left            =   6480
            TabIndex        =   9
            Top             =   240
            Width           =   4080
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_DataEntryUpto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, r, mLotDetail As Variant
Private Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 6
MSHFlexGrid2.TextMatrix(0, 0) = "Sr No "
MSHFlexGrid2.TextMatrix(0, 1) = "EHOP CMP ID "
MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 3) = "Date Upto which Data Entry Completed"
MSHFlexGrid2.TextMatrix(0, 4) = "CMPID"
MSHFlexGrid2.TextMatrix(0, 5) = "Old Date"


MSHFlexGrid2.ColWidth(0) = 800
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 6500
MSHFlexGrid2.ColWidth(3) = 3000
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 0


End Sub

Private Sub Check1_Click()
Dim Colo As Variant
Colo = vbWhite
If Check1.Value = 1 Then
   Colo = &HFFC0C0
End If

For r = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(r, 3) <> "" Then
       DataUpto.Value = MSHFlexGrid2.TextMatrix(r, 3)
       TxtCMPName = MSHFlexGrid2.TextMatrix(r, 2)
       MSHFlexGrid2.row = r
       MSHFlexGrid2.Col = 1
       MSHFlexGrid2.CellBackColor = Colo
       MSHFlexGrid2.Col = 2
       MSHFlexGrid2.CellBackColor = Colo
    End If
Next
End Sub

Private Sub CmdAdd2List_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Please select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If DataUpto.Value = Date Then
   If Format(Now, "hh") < "19" Then
      MsgBox "Current date is not allowed!!!"
      DataUpto.SetFocus
      Exit Sub
   End If
End If

If DataUpto.Value >= Date Then
   MsgBox "Future date is not allowed!!!"
   DataUpto.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(DataUpto.Value, "yyyy-mm") Then
   MsgBox "'Date Upto' allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   DataUpto.SetFocus
   Exit Sub
End If

'If Check1.Value = 1 Then
'   For i = 1 To MSHFlexGrid2.Rows - 1
'       If MSHFlexGrid2.TextMatrix(i, 5) <> "" Then
'        If CDate(MSHFlexGrid2.TextMatrix(i, 5)) > DataUpto.Value Then
'           MsgBox "Date Enter for CMP " & MSHFlexGrid2.TextMatrix(i, 2) & " (" & MSHFlexGrid2.TextMatrix(i, 5) & ") Is less than"
'           Exit Sub
'        End If
'       End If
'   Next i
'End If

If Check1.Value = 1 Then
   For I = 1 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(I, 5) <> "" Then
          If CDate(MSHFlexGrid2.TextMatrix(I, 5)) < DataUpto.Value Then
             If Gen_DBType = "MDB" Then
                tmp = "SELECT  CxTFNo, TxnType, CxTFDate, Flag, CMPID From Txn_CxTF_Details " & _
                      "WHERE (CxTFDate <= #" & Format(DataUpto.Value, Gen_DatFormat) & "# ) AND (Flag IS NULL) AND (CMPID = " & Val(MSHFlexGrid2.TextMatrix(I, 4)) & ")"
             Else
                tmp = "SELECT  CxTFNo, TxnType, CxTFDate, Flag, CMPID From Txn_CxTF_Details " & _
                      "WHERE (CxTFDate <= '" & Format(DataUpto.Value, Gen_DatFormat) & "' ) AND (Flag IS NULL) AND (CMPID = " & Val(MSHFlexGrid2.TextMatrix(I, 4)) & ")"
             End If
             Call TmpTable
             MSHFlexGrid2.Col = 3
             MSHFlexGrid2.row = I
             If TmpRs.RecordCount = 0 Then
                MSHFlexGrid2.TextMatrix(I, 3) = DataUpto.Value
                MSHFlexGrid2.CellBackColor = vbWhite
             Else
                MSHFlexGrid2.CellBackColor = vbYellow
                Label1.Visible = True
             End If
          End If
       End If
   Next I
   Check1.Value = 0
   TxtCMPName = ""
   Exit Sub
End If

r = MSHFlexGrid2.RowSel
If r = 0 Then Exit Sub

If LCase(Gen_UserRole) <> "power" Then
   If CDate(MSHFlexGrid2.TextMatrix(r, 5)) > DataUpto.Value Then
      MsgBox "Error: Date entered for CMP " & MSHFlexGrid2.TextMatrix(r, 2) & " is less than (" & MSHFlexGrid2.TextMatrix(r, 5) & ")!!!!!"
      Exit Sub
   End If
End If

If Gen_DBType = "MDB" Then
   tmp = "SELECT  CxTFNo, TxnType, CxTFDate, Flag, CMPID From Txn_CxTF_Details " & _
      "WHERE (CxTFDate <= #" & Format(DataUpto.Value, Gen_DatFormat) & "# ) AND (Flag IS NULL) AND (CMPID = " & Val(MSHFlexGrid2.TextMatrix(r, 4)) & ")"
Else
   tmp = "SELECT  CxTFNo, TxnType, CxTFDate, Flag, CMPID From Txn_CxTF_Details " & _
         "WHERE (CxTFDate <= '" & Format(DataUpto.Value, Gen_DatFormat) & "' ) AND (Flag IS NULL) AND (CMPID = " & Val(MSHFlexGrid2.TextMatrix(r, 4)) & ")"
End If

Call TmpTable
If TmpRs.RecordCount = 0 Then
   MSHFlexGrid2.TextMatrix(r, 3) = DataUpto.Value
Else
   MSHFlexGrid2.Col = 3
   MSHFlexGrid2.row = r
   MSHFlexGrid2.CellBackColor = vbYellow
   Label1.Visible = True
End If

CmdAdd2List.Caption = "Update Date"
DataUpto.Value = Date
TxtCMPName = ""

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
'   Call GButtonLockAD(Me)
   MsgBox "Access denied!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Exit Sub
End If

Call Grid_Head
CmbLocationID = ""
DataUpto.Value = Date
Call TableMst_CMP(" Where CloseDate is null ")
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

r = 0
mLotDetail = 0
End Sub
Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid2_Click()
CmdAdd2List.Caption = "Update Date"
Dim Colo, x As Variant



r = MSHFlexGrid2.RowSel
x = r
If r <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(r, 3) <> "" Then
   DataUpto.Value = MSHFlexGrid2.TextMatrix(r, 3)
   TxtCMPName = MSHFlexGrid2.TextMatrix(r, 2)
   MSHFlexGrid2.row = r
   MSHFlexGrid2.Col = 1
   MSHFlexGrid2.CellBackColor = &HFFC0C0
   MSHFlexGrid2.Col = 2
   MSHFlexGrid2.CellBackColor = &HFFC0C0
   
End If


Colo = vbWhite
For r = 1 To MSHFlexGrid2.Rows - 1
    If x <> r Then
       If MSHFlexGrid2.TextMatrix(r, 3) <> "" Then
          MSHFlexGrid2.row = r
          MSHFlexGrid2.Col = 1
          MSHFlexGrid2.CellBackColor = Colo
          MSHFlexGrid2.Col = 2
          MSHFlexGrid2.CellBackColor = Colo
       End If
    End If
Next
MSHFlexGrid2.row = x
End Sub
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
'Call TableMst_Location(" Where  LocationID <> 0 ")

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
   mLocationID = Null
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
If mLocationID <> RsMst_Location!LocationID Then
   Call Grid_Head
   Check1.Value = 0
End If
mLocationID = RsMst_Location!LocationID
Call FillGrid(mLocationID)
CmdAdd2List.Caption = "Update Date"
End Sub
Private Sub FillGrid(mLocationID_)
Dim k As Variant
Call Grid_Head
RsMst_CMP.Filter = "LocationId = " & mLocationID_
If RsMst_CMP.RecordCount > 0 Then
   RsMst_CMP.MoveFirst
   MSHFlexGrid2.Rows = RsMst_CMP.RecordCount + 2
   k = 1
   For I = 0 To RsMst_CMP.RecordCount - 1
       MSHFlexGrid2.TextMatrix(k, 0) = k
       MSHFlexGrid2.TextMatrix(k, 1) = RsMst_CMP!EHOPCMPID
       MSHFlexGrid2.TextMatrix(k, 2) = RsMst_CMP!CMPName
       MSHFlexGrid2.TextMatrix(k, 3) = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate)
       MSHFlexGrid2.TextMatrix(k, 4) = RsMst_CMP!CMPID
       MSHFlexGrid2.TextMatrix(k, 5) = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate)
       k = k + 1
       RsMst_CMP.MoveNext
   Next
Else
   MsgBox "No CMP found for selected location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
RsMst_CMP.Filter = "LocationID <> Null"

End Sub

Private Sub SaveCmd_Click()
Dim k As Variant
If MSHFlexGrid2.Rows <= 2 Then Exit Sub
k = 1

tmp = "Select * from Log_BODUpdate"
Call TmpTable

For I = 0 To RsMst_CMP.RecordCount - 1
    If MSHFlexGrid2.TextMatrix(k, 1) = "" Then Exit For
    RsMst_CMP.MoveFirst
    RsMst_CMP.Find "CMPID  = " & Val(MSHFlexGrid2.TextMatrix(k, 4))
    If Not RsMst_CMP.EOF Then
       RsMst_CMP!CurrentDate = MSHFlexGrid2.TextMatrix(k, 3)
       RsMst_CMP.Update
    End If
        
    TmpRs.AddNew
    TmpRs!CMPID = Val(MSHFlexGrid2.TextMatrix(k, 4))
    TmpRs!NewBODDate = MSHFlexGrid2.TextMatrix(k, 3)
    TmpRs!OldBODDate = MSHFlexGrid2.TextMatrix(k, 5)
    TmpRs!UpdatedBy = Gen_UserLoginID
    TmpRs!UpdatedDate = Now
    TmpRs.Update
    k = k + 1
Next

'Send Mail
MsgBox "Data saved successfully!!! "

Call Gen_Mail

TOADD = ""

'If LCase(CmbLocationID.Text) = "unjha" Then
'
'   tmp = "update mst_gsl set flag = 'E' where BalanceBags <= 0 and flag = 'J'"
'   Call TmpTable
'
'   tmp = "Select ML.LocationName,MC.CMPName,MC.CurrentDate,MG.GodownNo,MGSL.StackNo,Count(MGSL.LotNo) as 'NoofLot',Sum(BalanceBags) 'Bags',Sum(BalanceWeight) 'Weight' "
'   tmp = tmp & " From Mst_GSL MGSL "
'   tmp = tmp & " Inner Join Mst_Godown MG On MGSL.GodownID = MG.GodownID"
'   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
'   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
'   tmp = tmp & " Inner Join Mst_ExchangeType ME On MGSL.ExchangeTypeID = ME.ExchangeTypeID"
'   tmp = tmp & " Where MGSL.Flag not in ('B','E','Z') And MGSL.SpillageFlag = 0  And ME.ExchangeType = 'NCDEX' And ML.LocationName = '" & CmbLocationID.Text & "'"
'   tmp = tmp & " Group By ML.LocationName,MC.CMPName,MC.CurrentDate,MG.GodownNo,MGSL.StackNo  "
'   tmp = tmp & " Having Count(MGSL.LotNo) > 3 "
'   Call TmpTable
'
'   mMsgSubject = "No. of Live Lot for " & CmbLocationID.Text & " "
'   mMsg = ""
'   mGridMsg = ""
'   If TmpRs.RecordCount > 0 Then
'      mMsg = "<tr><td>Location Name</td><td>CMP Name</td><td>BOD Date</td><td>Godown No</td><td>Stack No</td><td>No of Lot</td><td>No of Bags</td><td>Quantity</td></tr>"
'      mGridMsg = ""
'      For I = 1 To TmpRs.RecordCount
'          mGridMsg = mGridMsg & "<tr><td> " & TmpRs!LocationName & " </td><td> " & TmpRs!CMPName & " </td><td> " & Format(TmpRs!CurrentDate, "dd-mmm-yyyy") & " </td><td>" & TmpRs!GodownNo & "</td><td>" & TmpRs!StackNo & "</td><td>" & TmpRs!NoofLot & "</td><td>" & TmpRs!Bags & "</td><td>" & TmpRs!Weight & "</td></tr>"
'          TmpRs.MoveNext
'      Next
'
'      mMsg = mMsg & mGridMsg
'   Else
'      mMsg = "No Live Lot present."
'   End If
'
'
'   'TOADD = "parag.r@ncml.com" ,mayur.d@ncml.com" '"cmg@ncml.com" ' mEmployeeEmailID
'
'   TOADD = "sankar.k@ncml.com,sriram.murthy@ncml.com"
'
'   mLotDetail = 1
'   Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
'
'End If

Call Grid_Head
mLotDetail = 0


End Sub
Private Function Gen_Mail()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg, mMsgSubject As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsgSubject = "BOD Update Details for " & CmbLocationID.Text

mMsg = ""
'mMsg = "<tr>BOD Update Details for " & CmbLocationID.Text & "</tr>"
'mMsg = mMsg & "<tr><td> Location </td><td> " & CmbLocationID.Text & "</td></tr>"

mMsg = "<tr><td> Location </td><td> " & CmbLocationID.Text & "</td></tr>"

mMsg = mMsg & "<tr><td> Sr.No. </td><td> CMP Name </td><td> Previous BOD Date </td><td> Present BOD Date </td></tr>"

mCount = 1
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 3) <> MSHFlexGrid2.TextMatrix(I, 5) Then
       mGridMsg = mGridMsg & "<tr><td> " & mCount & " </td><td> " & MSHFlexGrid2.TextMatrix(I, 2) & " </td><td>" & Format(MSHFlexGrid2.TextMatrix(I, 5), "dd-mmm-yyyy") & "</td><td>" & Format(MSHFlexGrid2.TextMatrix(I, 3), "dd-mmm-yyyy") & "</td></tr>"
       mCount = mCount + 1
    End If
Next

mMsg = mMsg & mGridMsg

TOADD = "cmg@ncml.com" ' mEmployeeEmailID

mLotDetail = 0

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
'  If mLotDetail = 1 Then
'     .MailFrom = "mayur.d@ncml.com" 'Gen_UserEmailID  'TOCC_ '"mayur.d@ncml.com"
'  Else
'     .MailFrom = Gen_UserEmailID
'  End If
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "parag.r@ncml.com" '"bhagwat.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  'TOCC_ = TOCC_ '& "," & mEmployeeEmailID
''  If ToADD_ <> TOCC_ Then
''     .CC = TOCC_
''  End If
'  If mLotDetail = 1 Then
'     .BCC = "mayur.d@ncml.com"
'  End If
'  .MessageSubject = MsgSubject_ '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
'objSMTP.FromAddr = Gen_UserEmailID
If mLotDetail = 1 Then
  objSMTP.FromAddr = "mayur.d@ncml.com" 'Gen_UserEmailID  'TOCC_ '"mayur.d@ncml.com"
Else
  objSMTP.FromAddr = Gen_UserEmailID
End If

If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "parag.r@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If
If mLotDetail = 1 Then
  objSMTP.BCCAddr = "mayur.d@ncml.com"
End If

objSMTP.Subject = MsgSubject_ '& " (from HR-Pro)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If



End Function

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub SetScrollBars()
'    Frame1.Width = 10000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub

Private Sub HScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l

End Sub
