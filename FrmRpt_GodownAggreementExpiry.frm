VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GodownAggreementExpiry 
   Caption         =   "Godown Agreement Expiry Details"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7890
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   15060
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1455
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2566
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7140
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   12594
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   7515
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   -15
      Width           =   15015
      Begin VB.CommandButton CmdMail 
         Caption         =   "Send Mail"
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
         Left            =   13650
         TabIndex        =   7
         Top             =   870
         Width           =   1215
      End
      Begin VB.TextBox TxtExpiryOn 
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
         Height          =   345
         Left            =   9810
         MaxLength       =   200
         TabIndex        =   4
         Top             =   457
         Width           =   1785
      End
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
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   457
         Width           =   3195
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
         Left            =   12600
         TabIndex        =   8
         Top             =   870
         Width           =   1050
      End
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
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   457
         Width           =   3195
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   6585
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   457
         Width           =   3195
      End
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
         Left            =   11505
         TabIndex        =   6
         Top             =   870
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
         Left            =   10320
         TabIndex        =   5
         Top             =   870
         Width           =   1185
      End
      Begin VB.Label lblExpiryOn 
         AutoSize        =   -1  'True
         Caption         =   "Expire Within Days"
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
         Left            =   9765
         TabIndex        =   16
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   1170
         TabIndex        =   14
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   7665
         TabIndex        =   13
         Top             =   195
         Width           =   1035
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
         Caption         =   "Location"
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
         Left            =   4575
         TabIndex        =   12
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   14160
         TabIndex        =   11
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record Found :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11760
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownAggreementExpiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID As Double

Private Sub CmbCMPName_GotFocus()
tmp = CmbCMPName.Text
CmbCMPName.Clear

If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record Found!!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   mCMPID = 0
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text
CmbLocationID.Clear

If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbStateID.SetFocus
   Exit Sub
End If

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID=" & mStateID & "")
Else
   Call TableMst_Location(" Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID

If tmp <> CmbLocationID.Text Then
   CmbCMPName.Text = ""
End If
End Sub
Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" Where " & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   mStateID = 0
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID

If tmp <> CmbStateID.Text Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
End If

End Sub

Private Sub CmdClear_Click()
mStateID = 0
mLocationID = 0
mCMPID = 0
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
Label2.Caption = ""
TxtExpiryOn.Text = ""
Call Grid_Head
Call Grid_Head1
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String

If Trim(TxtExpiryOn) = "" Then
   MsgBox "Enter Days"
   TxtExpiryOn.SetFocus
   Exit Sub
End If

If IsNumeric(TxtExpiryOn) = False Then
   MsgBox "Invalid Format"
   TxtExpiryOn.SetFocus
   Exit Sub
End If


MsgBox "Wait till next message"

Call Grid_Head
wc = ""
If CmbStateID.Text <> "" Then
   wc = "ms.StateID = " & mStateID
End If

If CmbLocationID.Text <> "" Then
   wc = "ML.LocationID = " & mLocationID
End If

If CmbCMPName.Text <> "" Then
   wc = "MC.CMPID = " & mCMPID
End If

tmp = "Select MS.StateName, ML.LocationName, LA.LAID,  case When lower(LA.Flag) ='a' then 'Approved' else case when lower(LA.Flag) ='r' then 'Rejected' else case when lower(LA.Flag) ='p' then 'Pending' end end end as Flag, "
tmp = tmp & " MC.CMPName, MG.GodownNo, MG.Address, MLE.LessorName,"
tmp = tmp & " LA.EntryDate, LA.StartDate, LA.ExpiryDate, TGLAD.RentAmount,"
tmp = tmp & " LA.AgreementRecvAs, LA.AdvanceGiven,"
tmp = tmp & " LA.AdvanceAdjusted, LA.AdvanceReturn, LA.AdvanceBalance,"
tmp = tmp & " MG.Area, MST.Description, LA.Remark, MG.CloseDate,"
tmp = tmp & " '' as GodownStock, MG.GodownCapacity, MG.GodownID, MC.CMPID, ML.LocationID, MS.StateID"
tmp = tmp & " From Mst_Godown  MG"
tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail TGLAD On TGLAD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Txn_LessorAgreement LA on LA.LAID =  TGLAD.LAID"
tmp = tmp & " Left Join Mst_ServiceTax MST on MST.ServiceTaxID = LA.ServiceTaxID"
tmp = tmp & " Inner Join Mst_Lessor MLE on MLE.LessorID = TGLAD.LessorID"
tmp = tmp & " Inner join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Where MG.CloseDate Is Null And MU.UnitTypeID = 1"
tmp = tmp & " And LA.ExpiryDate >= '" & Date & "' And LA.ExpiryDate <= dateadd(day," & TxtExpiryOn & ",'" & Date & "')"

If wc <> "" Then
   tmp = tmp & " And " & wc
End If
tmp = tmp & " Order by MS.StateName, ML.LocationName,MC.CMPName,MG.GodownNo"

Call TmpTable


      
tmp = "SELECT Mst_Godown.GodownID, Sum(Mst_GSL.DepositeWeight)-Sum(Mst_GSL.CDTFWeight) as Stock "
tmp = tmp & " From Mst_Godown"
tmp = tmp & " INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " Where Mst_GSL.Status <> 'E'"
tmp = tmp & " Group By Mst_Godown.GodownID"
tmp = tmp & " Having Sum(Mst_GSL.DepositeWeight) - Sum(Mst_GSL.CDTFWeight) <> 0"
tmp = tmp & " Order By Mst_Godown.GodownID"

Call Tmp3Table

tmp = "Select ECM.LocationID, EMP.EmployeeName+'~'+EMP.EmployeeNo  as 'CMPMgr', EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMName', EMP2.EmployeeEmailID as AMEmailID, EMP2.EmployeeID  as AMEMPID,"
tmp = tmp & " EMP3.EmployeeName +'~'+EMP3.EmployeeNo as 'RMName', EMP3.EmployeeEmailID as RMEmailID, EMP3.EmployeeID  as RMEMPID, EMP.EmployeeID  as CMPMGrEMPID, EMP.EmployeeEmailID as CMPMGrmailID"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID And EMP.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP.DirectReportingTo = EMP2.EmployeeID And EMP2.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID And EMP3.Flag = 'Y'"
tmp = tmp & " Where EMP2.EmployeeName Is Not Null"
tmp = tmp & " Group By ECM.LocationID, EMP2.EmployeeName, EMP2.EmployeeNo"
tmp = tmp & " ,EMP3.EmployeeName ,EMP3.EmployeeNo, EMP2.EmployeeEmailID, EMP3.EmployeeEmailID, EMP2.EmployeeID, EMP3.EmployeeID, EMP.EmployeeName, EMP.EmployeeNo, EMP.EmployeeID, EMP.EmployeeEmailID"

Call Tmp2Table

tmp = "Select * From  Mst_Employee Where DesignationID = 10 And Flag = 'Y'"

Call Tmp4Table

Call dispdata

MsgBox "Done!!!"

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FrmRpt_GodownAggreementExpiry.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "FrmRpt_GodownAggreementExpiry.xls")
Set oWS = oWB.Worksheets("Sheet1")
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
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdMail_Click()
mFlag = True
mAM = ""
mRow = 0
j = 1
If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No Data Exists"
   Exit Sub
End If

Call Grid_Head1
For I = 1 To MSHFlexGrid1.Rows - 1
    mAM = ""
    mFlag = True
    If MSHFlexGrid1.TextMatrix(I, 23) <> "" Then
       mAM = MSHFlexGrid1.TextMatrix(I, 23)
       For j = 1 To MSHFlexGrid2.Rows - 1
           If MSHFlexGrid2.TextMatrix(j, 0) = mAM Then
              mFlag = False
              Exit For
           End If
       Next
       If mFlag = True Then
          mRow = MSHFlexGrid2.Rows - 1
          MSHFlexGrid2.TextMatrix(mRow, 0) = mAM
          MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
        End If
    End If
Next
Call Gen_Mail
MsgBox "Done!!"
End Sub
Private Function Gen_Mail()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg, mMsgSubject, mGridMsgBlank As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim AM_ As String
Dim mFlag As Boolean
mMsgSubject = "Lease Godown agreement Expiry reminder"
mMsg = ""
AM_ = ""
mFlag = True
For I = 1 To MSHFlexGrid2.Rows - 1
    AM_ = ""
    mGridMsg = ""
    TOADD = ""
    TOCC = ""
    mGridMsgBlank = ""
    mEmailId = ""
    AM_ = MSHFlexGrid2.TextMatrix(I, 0)
    For j = 1 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(j, 23) <> "" Then
           If MSHFlexGrid1.TextMatrix(j, 23) = AM_ Then
              mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 7) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 8) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 9) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 10) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 11) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 12) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 13) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 14) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 15) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 16) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 17) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 18) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 19) & " </td>"
              mGridMsg = mGridMsg & "<td> " & MSHFlexGrid1.TextMatrix(j, 20) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 21) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 22) & " </td></tr>"
              'AM
              If InStr(1, TOADD, MSHFlexGrid1.TextMatrix(j, 23)) = 0 Then
                 If TOADD = "" Then
                    TOADD = TOADD & MSHFlexGrid1.TextMatrix(j, 23)
                 Else
                    TOADD = TOADD & "," & MSHFlexGrid1.TextMatrix(j, 23)
                 End If
              End If
              'OC
              If InStr(1, TOADD, MSHFlexGrid1.TextMatrix(j, 25)) = 0 Then
                 If TOADD = "" Then
                    TOADD = TOADD & MSHFlexGrid1.TextMatrix(j, 25)
                 Else
                    TOADD = TOADD & "," & MSHFlexGrid1.TextMatrix(j, 25)
                 End If
              End If
              
              'CMPMgr
              If InStr(1, TOADD, MSHFlexGrid1.TextMatrix(j, 26)) = 0 Then
                 If TOADD = "" Then
                    TOADD = TOADD & MSHFlexGrid1.TextMatrix(j, 26)
                 Else
                    TOADD = TOADD & "," & MSHFlexGrid1.TextMatrix(j, 26)
                 End If
              End If
              
              'RM
              If InStr(1, TOCC, MSHFlexGrid1.TextMatrix(j, 24)) = 0 Then
                 If TOCC = "" Then
                    TOCC = TOCC & MSHFlexGrid1.TextMatrix(j, 24)
                 Else
                    TOCC = TOCC & "," & MSHFlexGrid1.TextMatrix(j, 24)
                 End If
              End If
           End If
        Else
              'mGridMsgBlank = mGridMsgBlank & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td></tr>"
              mGridMsgBlank = mGridMsgBlank & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 7) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 8) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 9) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 10) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 11) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 12) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 13) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 14) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 15) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 16) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 17) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 18) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 19) & " </td>"
              mGridMsgBlank = mGridMsgBlank & "<td> " & MSHFlexGrid1.TextMatrix(j, 20) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 21) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 22) & " </td></tr>"
        End If
    Next
    If mGridMsg <> "" Then
       mMsg = "<tr><td> State </td><td>Location</td><td>Lessor Agreement ID</td><td>Status</td><td>CMP Name</td><td>Godown No</td><td>Godown Address</td><td>Lessor Name</td><td>Agreement Entry Date</td><td>Agreement Start Date</td><td>Agreement Expiry Date</td><td>Rent Amount</td><td>Agreement Rec. As</td><td>Security Deposite Given</td><td>Security Deposite Adjusted</td><td>Security Deposite Returned</td><td>Balance</td><td>Area(in Sq.ft.)</td><td>Service Tax Description</td><td>Remark</td><td>Godown Close Date</td><td>Current Godown Stock</td><td>Capacity</td></tr>"
       mMsg = mMsg & mGridMsg
       'TOADD = "samar.k@ncml.com"
       'TOCC = "parag.r@ncml.com"
        'CMG
        If TOCC = "" Then
           TOCC = TOCC & "cmg@ncml.com"
        Else
           TOCC = TOCC & "," & "cmg@ncml.com"
        End If
        Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
    End If
    If mGridMsgBlank <> "" Then
       mMsg = "<tr><td> State </td><td>Location</td><td>Lessor Agreement ID</td><td>Status</td><td>CMP Name</td><td>Godown No</td><td>Godown Address</td><td>Lessor Name</td><td>Agreement Entry Date</td><td>Agreement Start Date</td><td>Agreement Expiry Date</td><td>Rent Amount</td><td>Agreement Rec. As</td><td>Security Deposite Given</td><td>Security Deposite Adjusted</td><td>Security Deposite Returned</td><td>Balance</td><td>Area(in Sq.ft.)</td><td>Service Tax Description</td><td>Remark</td><td>Godown Close Date</td><td>Current Godown Stock</td><td>Capacity</td></tr>"
       mMsg = mMsg & mGridMsgBlank
       TOADD = "cmg@ncml.com"
       TOCC = ""
       Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
    End If
Next
End Function

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True

mStateID = 0
mLocationID = 0
mCMPID = 0

Call Grid_Head

CmdExcel.Enabled = False
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 27
MSHFlexGrid1.WordWrap = True


MSHFlexGrid1.TextMatrix(0, 0) = "State"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Lessor Agreement ID"
MSHFlexGrid1.TextMatrix(0, 3) = "Status"
MSHFlexGrid1.TextMatrix(0, 4) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 6) = "Godown Address"
MSHFlexGrid1.TextMatrix(0, 7) = "Lessor Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Agreement Entry Date"
MSHFlexGrid1.TextMatrix(0, 9) = "Agreement Start Date"
MSHFlexGrid1.TextMatrix(0, 10) = "Agreement Expiry Date"
MSHFlexGrid1.TextMatrix(0, 11) = "Rent Amount"
MSHFlexGrid1.TextMatrix(0, 12) = "Agreement Rec. As"
MSHFlexGrid1.TextMatrix(0, 13) = "Security Deposite Given"
MSHFlexGrid1.TextMatrix(0, 14) = "Security Deposite Adjusted"
MSHFlexGrid1.TextMatrix(0, 15) = "Security Deposite Returned"
MSHFlexGrid1.TextMatrix(0, 16) = "Balance"
MSHFlexGrid1.TextMatrix(0, 17) = "Area(in Sq.ft.)"
MSHFlexGrid1.TextMatrix(0, 18) = "Service Tax Description"
MSHFlexGrid1.TextMatrix(0, 19) = "Remark"
MSHFlexGrid1.TextMatrix(0, 20) = "Godown Close Date"
MSHFlexGrid1.TextMatrix(0, 21) = "Current Godown Stock"
MSHFlexGrid1.TextMatrix(0, 22) = "Capacity"
MSHFlexGrid1.TextMatrix(0, 23) = "AM"
MSHFlexGrid1.TextMatrix(0, 24) = "RM"
MSHFlexGrid1.TextMatrix(0, 25) = "OC"
MSHFlexGrid1.TextMatrix(0, 26) = "CMPManager"

MSHFlexGrid1.RowHeight(0) = 800
End Sub
Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 1
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.TextMatrix(0, 0) = "AMName"

MSHFlexGrid2.ColWidth(0) = 1800

End Sub
Public Sub dispdata()
ProgressBar1.Value = 0
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount
   For I = 1 To TmpRs.RecordCount
       ProgressBar1.Value = I
       For C = 0 To 22
           If C = 8 Or C = 9 Or C = 10 Or C = 20 Then
              MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", Format(TmpRs.Fields(C), "dd-MMM-yyyy"))
           ElseIf C = 21 Then 'Current Godown Stock
              TmpRs3.MoveFirst
              TmpRs3.Find "GodownID = " & TmpRs!GodownID
              If Not TmpRs3.EOF Then
                 MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs3!Stock), "", TmpRs3!Stock)
              End If
           Else
              MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           End If
       Next
       
       If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
       TmpRs2.Filter = "LocationID = " & TmpRs!LocationID
       If TmpRs2.RecordCount > 0 Then
          'CMPMgr
          For k = 1 To TmpRs2.RecordCount
              If MSHFlexGrid1.TextMatrix(I, 26) = "" Then
                 MSHFlexGrid1.TextMatrix(I, 26) = TmpRs2!CMPMGrmailID
              Else
                 MSHFlexGrid1.TextMatrix(I, 26) = MSHFlexGrid1.TextMatrix(I, 26) & "," & TmpRs2!CMPMGrmailID
              End If
              TmpRs2.MoveNext
          Next
          If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst 'TmpRs2.MoveFirst
          If Not TmpRs2.EOF Then
            'AM / RM
             MSHFlexGrid1.TextMatrix(I, 23) = TmpRs2!AMEmailID
             MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs2!RMEmailID), "", TmpRs2!RMEmailID)
             'OC
             TmpRs4.MoveFirst
             TmpRs4.Filter = "DirectReportingTo = " & TmpRs2!AMEMPID
             If TmpRs4.RecordCount > 0 Then
                For j = 1 To TmpRs4.RecordCount
                    If MSHFlexGrid1.TextMatrix(I, 25) = "" Then
                       MSHFlexGrid1.TextMatrix(I, 25) = TmpRs4!EmployeeEmailID
                    Else
                       MSHFlexGrid1.TextMatrix(I, 25) = MSHFlexGrid1.TextMatrix(I, 25) & "," & TmpRs4!EmployeeEmailID
                    End If
                    TmpRs4.MoveNext
                Next
             End If
             TmpRs4.Filter = ""
          End If
          TmpRs2.Filter = ""
       End If
       TmpRs.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Next
End If
Label2.Caption = TmpRs.RecordCount
CmdExcel.Enabled = True
End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue>This is to inform you that the following godown lease agreements are going to expire within next" & Trim(TxtExpiryOn) & ". If you want the godowns to continue, initiate renewal process for the same, otherwise issue termination notice to the godown owner and intimate CMG.<br><br><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
'ToADD_ = "samar.k@ncml.com"
'TOCC_ = ""
'With oSMTP
'  .Server = servertxt
'  .MailFrom = Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  .SendTo = ToADD_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  .BCC = "parag.r@ncml.com"
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With
'End Function

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function


