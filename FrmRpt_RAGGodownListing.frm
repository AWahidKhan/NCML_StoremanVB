VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_RAGGodownListing 
   Caption         =   "RAG Godown Listing"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8415
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8055
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14208
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbGodownName 
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
         Height          =   360
         Left            =   7920
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   2715
      End
      Begin VB.ComboBox CmbLocation 
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
         Height          =   360
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   2460
      End
      Begin VB.TextBox TxtSBalBags 
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
         Height          =   315
         Left            =   2535
         TabIndex        =   4
         Top             =   915
         Width           =   1365
      End
      Begin VB.ComboBox CmbEmployeeID 
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
         Left            =   10725
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox CmbCMPID 
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
         Height          =   360
         Left            =   5040
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2790
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   2355
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
         Left            =   12390
         TabIndex        =   6
         Top             =   885
         Width           =   1215
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
         Left            =   11175
         TabIndex        =   5
         Top             =   885
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
         Left            =   13605
         TabIndex        =   7
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "Godown No"
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
         Left            =   8737
         TabIndex        =   19
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label Label53 
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
         Height          =   255
         Left            =   3405
         TabIndex        =   17
         Top             =   195
         Width           =   840
      End
      Begin VB.Label Label13 
         Caption         =   "Balance No of Bags "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   915
         Width           =   1920
      End
      Begin VB.Label Label3 
         Caption         =   "RAG Member Name"
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
         Left            =   11805
         TabIndex        =   15
         Top             =   195
         Width           =   1935
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   5895
         TabIndex        =   14
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label LblStateID 
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
         Height          =   255
         Left            =   705
         TabIndex        =   13
         Top             =   195
         Width           =   1200
      End
      Begin VB.Label Label2 
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
         Left            =   5055
         TabIndex        =   10
         Top             =   945
         Width           =   2295
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   7455
         TabIndex        =   9
         Top             =   945
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmRpt_RAGGodownListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mStateID, mLocationID, mCMPID, mGodownID, mEmployeeID As Variant

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
End Sub

Private Sub CmbLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State !!"
   CmbLocation.Clear
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where StateID=" & mStateID & " And " & Gen_WcLocation)
Else
   Call TableMst_Location(" Where StateID=" & mStateID & "")
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocation.Text <> tmp Then
   CmbCMPID.Clear
   CmbCMPID.Text = ""
   mCMPID = 0
   CmbGodownName.Clear
   CmbGodownName.Text = ""
   mGodownID = 0
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub
Private Sub CmbCMPID_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Please Select Location !"
   CmbCMPID.Clear
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMPID.Text
If Gen_WcCMP <> "" Then
   Call TableMst_CMP("Where LocationID = " & mLocationID & " and CloseDate is Null And " & Gen_WcCMP)
Else
   Call TableMst_CMP("Where LocationID = " & mLocationID & " and CloseDate is Null ")
End If
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No CMP found"
   CmbCMPID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub

Private Sub CmbCMPID_LostFocus()

If CmbCMPID.Text = "" Then
   mCMPID = 0
   Exit Sub
End If

' Code for clearing Godown combox in case of change
If CmbCMPID.Text <> tmp Then
   CmbGodownName.Clear
   CmbGodownName.Text = ""
   mGodownID = 0
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid CMP selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub
Private Sub CmbGodownName_Gotfocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP!!!"
   CmbGodownName.Clear
   CmbCMPID.SetFocus
   Exit Sub
End If
tmp = CmbGodownName.Text
Call TableMst_Godown(" Where CMPID = " & mCMPID & " and CloseDate is Null ")
CmbGodownName.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Godown found  !!!!! "
   CmbGodownName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownName.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownName.Text = tmp
End Sub

Private Sub CmbGodownName_LostFocus()
If CmbGodownName.Text = "" Then
   mGodownID = 0
   Exit Sub
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownName.Text & "'"

If RsMst_Godown.EOF Then
   MsgBox "Invalid Godown Selection "
   CmbGodownName.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbEmployeeID_GotFocus()
tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y' And DesignationID=7")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbEmployeeID.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbEmployeeID.Text = tmp
End Sub

Private Sub CmbEmployeeID_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbEmployeeID.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbEmployeeID, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   CmbEmployeeID.SetFocus
   Exit Sub
End If
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbEmployeeID.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbStateID.Text = ""
CmbLocation.Text = ""
CmbCMPID.Text = ""
CmbGodownName.Text = ""
CmbEmployeeID.Text = ""
TxtSBalBags.Text = ""
LblTotalRecords.Caption = "-"
End Sub

Private Sub CmdGetReport_Click()
Dim HWc As String
Dim wc As String
Dim tmp1 As String
Call Grid_Head

MsgBox "Wait till next message!!!"

HWc = "": wc = ""
'ME.Flag='Y' And
wc = " Where   (ME.Flag='Y' OR ME.Flag is Null) And  MC.CloseDate is Null And MGSL.Flag Not In ('Z','B','E') "

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp1 = Replace(Gen_WcState, "StateID", "MS.StateID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbCMPID.Text <> "" Then
   wc = GenWc(wc, mCMPID, "MC.CmpID", "N", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp1 = Replace(Gen_WcCMP, "CMPID", "MC.CMPID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbGodownName.Text <> "" Then
   wc = GenWc(wc, mGodownID, "MG.GodownID", "N", "N")
End If

If CmbEmployeeID.Text <> "" Then
   wc = GenWc(wc, mEmployeeID, "ME1.EmployeeID", "N", "N")
End If

If TxtSBalBags <> "" Then
   If Not IsOprater(TxtSBalBags) Then
      HWc = GenWc(HWc, TxtSBalBags, "Sum(MGSL.BalanceBags)", "N", "N")
   Else
      HWc = GenWc(HWc, TxtSBalBags, "Sum(MGSL.BalanceBags)", "N", "I")
   End If
End If

tmp = "Select MS.StateName,ML.LocationName,MC.CMPID,MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,"
tmp = tmp & " ME.EmployeeName+'~'+ME.EmployeeNo As EmployeeName,ME.Flag,Sum(MGSL.BalanceBags) As BalanceBags,"
tmp = tmp & " ME1.EmployeeName+'~'+ME1.EmployeeNo As RAGName,MG.GodownID"
tmp = tmp & " From Mst_GSL MGSL"
tmp = tmp & " Inner Join Mst_Godown MG On MGSL.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MC.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID=ML.StateID"
tmp = tmp & " Inner Join Mst_UnitType MU On MU.UnitTypeID = MG.UnitTypeID"
tmp = tmp & " Left Join  Mst_EmpCMPMapping MEC On MC.CMPID = MEC.CMPID"
tmp = tmp & " Left Join Mst_Employee ME On MEC.EmployeeID=ME.EmployeeID " 'And ME.Flag='Y'"
tmp = tmp & " Left Join Mst_Employee ME1 On ME1.EmployeeID=ME.ConnectedToRAG"
If wc <> "" Then
   tmp = tmp & wc
End If
tmp = tmp & " Group By MS.StateName,ML.LocationName,MC.CMPID,MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,"
tmp = tmp & " Me.EmployeeName , Me.EmployeeNo, ME1.EmployeeName, ME1.EmployeeNo, Me.Flag,MG.GodownID"
If TxtSBalBags.Text <> "" Then
   tmp = tmp & Replace(HWc, "Where", " Having")
End If
tmp = tmp & " Order By  MS.StateName,ML.LocationName,MC.CMPID,MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType"
Call TmpTable

For I = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 2
        MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If LCase(TmpRs!Flag) = "n" Then
       MSHFlexGrid1.TextMatrix(I, 8) = "De-Active"
    ElseIf LCase(TmpRs!Flag) = "y" Then
       MSHFlexGrid1.TextMatrix(I, 8) = "Active"
    End If
    'Call GetNoOfBags(TmpRs!GodownID, TmpRs!SM_Prefix, TmpRs!CMPID, i)
    Call GetAuditDetails(TmpRs!GodownID, I)
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    TmpRs.MoveNext
Next
LblTotalRecords.Caption = TmpRs.RecordCount
MsgBox "Done !"
End Sub

Private Sub GetNoOfBags(mgodownID_ As Variant, mSM_Prefix As Variant, mCMPID_ As Variant, mRow_ As Variant)
tmp = "Select Sum(BalanceBags) as BalanceBags from Mst_GSL "
tmp = tmp & " Where GodownID =" & mgodownID_ & " "
tmp = tmp & " and SM_PreFix='" & mSM_Prefix & "' and CMPID=" & mCMPID_ & ""

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(mRow_, 9) = IIf(IsNull(TmpRs1!BalanceBags), "", TmpRs1!BalanceBags)
End If
End Sub

Private Sub GetAuditDetails(mgodownID_ As Variant, mRow_ As Variant)
tmp = " Select Max(TA.AuditDate) as ADate"
tmp = tmp & " From Txn_Audit TA"
tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID = TA.EmployeeID"
tmp = tmp & " Where GodownID = " & mgodownID_ '& " And ME.DesignationID = 7 or "

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   TmpRs1.MoveFirst
   MSHFlexGrid1.TextMatrix(mRow_, 11) = Format(TmpRs1!ADate, "dd-mmm-yyyy")
   'MSHFlexGrid1.TextMatrix(mRow_, 11) = IIf(IsNull(TmpRs1!PerceptionRisk), "", TmpRs1!PerceptionRisk)
End If
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 13
     .Font.Size = 10
     .AllowUserResizing = flexResizeBoth
     .WordWrap = True
     .TextMatrix(0, 0) = "State"
     .TextMatrix(0, 1) = "Location"
     .TextMatrix(0, 2) = "CMP ID"
     .TextMatrix(0, 3) = "CMP Name"
     .TextMatrix(0, 4) = "Godown No"
     .TextMatrix(0, 5) = "Godown Address"
     .TextMatrix(0, 6) = "Unit Type"
     .TextMatrix(0, 7) = "CMPM Name"
     .TextMatrix(0, 8) = "Status"
     .TextMatrix(0, 9) = "No. of Bags"
     .TextMatrix(0, 10) = "RAG member"
     .TextMatrix(0, 11) = "Last date of Audit"
     .TextMatrix(0, 12) = "Risk Category"
     .ColWidth(0) = 1500
     .ColWidth(1) = 1500
     .ColWidth(2) = 0
     .ColWidth(3) = 1800
     .ColWidth(4) = 1500
     .ColWidth(5) = 2000
     .ColWidth(6) = 1500
     .ColWidth(7) = 2400
     .ColWidth(8) = 1200
     .ColWidth(9) = 1200
     .ColWidth(10) = 2400
     .ColWidth(11) = 1400
     .ColWidth(12) = 2400
     .RowHeight(0) = 500
     .ColAlignment(4) = vbLeftJustify
End With
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_RAGGodownListing.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_RAGGodownListing.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Call Grid_Head
End Sub


Private Sub TxtSBalBags_LostFocus()
If TxtSBalBags.Text = "" Then Exit Sub
Dim mBul As Boolean

mBul = True

If InStr(2, TxtSBalBags.Text, "<") > 1 Then
   mBul = False
End If

If InStr(1, TxtSBalBags.Text, "<") = False And InStr(2, TxtSBalBags.Text, ">") Then
   mBul = False
End If

If InStr(1, TxtSBalBags.Text, "<") = False And InStr(1, TxtSBalBags.Text, ">") = False Then
   If InStr(2, TxtSBalBags.Text, "=") Then
      mBul = False
   End If
End If

If mBul = False Then
   MsgBox "Pleas enter Proper Condition."
   TxtSBalBags.SetFocus
   Exit Sub
End If
End Sub
