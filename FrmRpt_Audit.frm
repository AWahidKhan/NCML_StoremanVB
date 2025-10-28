VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_Audit 
   Caption         =   "Audit Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbCmp 
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
         Left            =   5715
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   2355
      End
      Begin VB.ComboBox CmbAuditSegment 
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
         Left            =   9855
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   2265
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmRpt_Audit.frx":0000
         Left            =   13080
         List            =   "FrmRpt_Audit.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1230
         Width           =   1815
      End
      Begin VB.ComboBox CmbAuditType 
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
         Left            =   12120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox CmbAuditPoint 
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
         Left            =   3855
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1230
         Width           =   9180
      End
      Begin VB.ComboBox CmbGodownName 
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
         Left            =   8100
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1710
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
         Left            =   3165
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   2535
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
         Height          =   495
         Left            =   2535
         TabIndex        =   11
         Top             =   1230
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
         Height          =   495
         Left            =   105
         TabIndex        =   10
         Top             =   1230
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
         Height          =   495
         Left            =   1320
         TabIndex        =   14
         Top             =   1230
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   61276161
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   61276161
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CMP"
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
         Left            =   6675
         TabIndex        =   23
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label8 
         Caption         =   "Audit Segment"
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
         Left            =   10320
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Status"
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
         Left            =   13680
         TabIndex        =   21
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Audit Type"
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
         Left            =   12885
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "To Audit Date"
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
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "From Audit  Date"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Audit Point"
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
         Left            =   7958
         TabIndex        =   17
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Godown"
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
         Left            =   8528
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7875
      Left            =   120
      TabIndex        =   12
      Top             =   1785
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_Audit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxRow, mGodownID, mCMPID, temp, mLocationID, mSegmentTypeID, mAuditTypeID As Variant

Dim strSplitString() As String
Dim strObservation As String
Dim strComment As String

Private Sub CmbAuditPoint_GotFocus()
tmp = CmbAuditPoint.Text
'If CmbAuditSegment.Text = "" Then Exit Sub
'If CmbAuditType.Text = "" Then Exit Sub
Call TableMst_AuditPoint("") '("Where SegmentTypeID = " & mSegmentTypeID & " And AuditTypeID = " & mAuditTypeID & " ")
CmbAuditPoint.Clear
If RsMst_AuditPoint.RecordCount = 0 Then
   MsgBox "No Audit Points found"
   CmbAuditPoint.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditPoint.RecordCount
    CmbAuditPoint.AddItem RsMst_AuditPoint!AuditPoint
    RsMst_AuditPoint.MoveNext
Next
CmbAuditPoint.Text = tmp
End Sub

Private Sub CmbAuditPoint_LostFocus()
If CmbAuditPoint.Text = "" Then
   mAuditPointID = Null
   Exit Sub
End If

RsMst_AuditPoint.MoveFirst
RsMst_AuditPoint.Find "AuditPoint = '" & CmbAuditPoint.Text & "'"

If RsMst_AuditPoint.EOF Then
   MsgBox "Invalid point Selection"
   CmbAuditPoint.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbCMP_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location."
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMP.Text

Call TableMst_CMP("Where LocationID =" & mLocationID & "")

CmbCMP.Clear

If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No CMP found  !!!!! "
   CmbCMP.SetFocus
   Exit Sub
End If
RsMst_CMP.MoveFirst
For I = 1 To RsMst_CMP.RecordCount
    CmbCMP.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMP.Text = tmp
End Sub

Private Sub CmbCMP_LostFocus()
If tmp <> CmbCMP.Text Then
   CmbGodownName.Clear
   CmbGodownName.Text = ""
End If
If CmbCMP.Text = "" Then
   mCMPID = 0
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"


If RsMst_CMP.EOF Then
   MsgBox "Invalid CMP Selection "
   CmbCMP.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub


Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) = "open" Or LCase(CmbStatus.Text) = "close" Then
Else
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()

Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmbLocation.Text = ""
CmbGodownName.Text = ""
CmbAuditPoint.Text = ""
CmbAuditType.Text = ""
CmbStatus.Text = ""
CmdExcel.Enabled = False
From_GenDate = Date
TxtFromDate.Value = CDate(From_GenDate) - 30 'Date
TxtToDate.Value = Date
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
'   CmdGetReport.Enabled = False
'   CmdExcel.Enabled = False
'   CmdClear.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Call CmdClear_Click
'Call Grid_Head
CmdExcel.Enabled = False
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
MSHFlexGrid1.Cols = 25 '30
MSHFlexGrid1.WordWrap = True


MSHFlexGrid1.TextMatrix(0, 0) = "Audit ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Audit Date"
MSHFlexGrid1.TextMatrix(0, 2) = "State"
MSHFlexGrid1.TextMatrix(0, 3) = "Location"
MSHFlexGrid1.TextMatrix(0, 4) = "Segment Type"
MSHFlexGrid1.TextMatrix(0, 5) = "CMP ID"
MSHFlexGrid1.TextMatrix(0, 6) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 7) = "Godown Code"
MSHFlexGrid1.TextMatrix(0, 8) = "Godown Address"
MSHFlexGrid1.TextMatrix(0, 9) = "Designation"
MSHFlexGrid1.TextMatrix(0, 10) = "Employee No"
MSHFlexGrid1.TextMatrix(0, 11) = "Auditor Name"
MSHFlexGrid1.TextMatrix(0, 12) = "Audit Type"
MSHFlexGrid1.TextMatrix(0, 13) = "Audit Point"
MSHFlexGrid1.TextMatrix(0, 14) = "Previous Obseravtion"
MSHFlexGrid1.TextMatrix(0, 15) = "Present Obseravtion"
MSHFlexGrid1.TextMatrix(0, 16) = "Previous Comment"
MSHFlexGrid1.TextMatrix(0, 17) = "Present Comment"
MSHFlexGrid1.TextMatrix(0, 18) = "Required"
MSHFlexGrid1.TextMatrix(0, 19) = "Actual"
MSHFlexGrid1.TextMatrix(0, 20) = "Required Date"
MSHFlexGrid1.TextMatrix(0, 21) = "Required Value"
MSHFlexGrid1.TextMatrix(0, 22) = "Start Date"
MSHFlexGrid1.TextMatrix(0, 23) = "Close Date"
MSHFlexGrid1.TextMatrix(0, 24) = "Flag"
'MSHFlexGrid1.TextMatrix(0, 25) = "CMP Manager"
'MSHFlexGrid1.TextMatrix(0, 26) = "CC"
'MSHFlexGrid1.TextMatrix(0, 27) = "AM"
'MSHFlexGrid1.TextMatrix(0, 28) = "RM"
'MSHFlexGrid1.TextMatrix(0, 29) = "Managment"

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 0
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 0
MSHFlexGrid1.ColWidth(5) = 0
MSHFlexGrid1.ColWidth(6) = 2000
MSHFlexGrid1.ColWidth(7) = 2000
MSHFlexGrid1.ColWidth(8) = 2200
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.ColWidth(10) = 0
MSHFlexGrid1.ColWidth(11) = 1600
MSHFlexGrid1.ColWidth(12) = 2400
MSHFlexGrid1.ColWidth(13) = 3000
MSHFlexGrid1.ColWidth(14) = 3000
MSHFlexGrid1.ColWidth(15) = 1200
MSHFlexGrid1.ColWidth(16) = 3000
MSHFlexGrid1.ColWidth(17) = 1200

MSHFlexGrid1.ColWidth(18) = 1200
MSHFlexGrid1.ColWidth(19) = 1600
MSHFlexGrid1.ColWidth(20) = 1600
MSHFlexGrid1.ColWidth(21) = 1600
MSHFlexGrid1.ColWidth(22) = 1200
MSHFlexGrid1.ColWidth(23) = 1200
MSHFlexGrid1.ColWidth(24) = 0
'MSHFlexGrid1.ColWidth(25) = 2400
'MSHFlexGrid1.ColWidth(26) = 2200
'MSHFlexGrid1.ColWidth(27) = 2200
'MSHFlexGrid1.ColWidth(28) = 2200
'MSHFlexGrid1.ColWidth(29) = 2800



'MSHFlexGrid1.RowHeight(0) = 1600

End Sub

Private Sub CmdGetReport_Click()
Dim Sdat, Edat As Variant

Call Grid_Head


MsgBox "Wait till next message!!!"

Sdat = Format(TxtFromDate, Gen_DatFormat)
Edat = Format(TxtToDate, Gen_DatFormat)
    
    
tmp = " SELECT Txn_Audit.AuditID, Txn_Audit.AuditDate,Mst_State.StateName,Mst_Location.LocationName, Mst_AuditSegment.SegmentType,Mst_CMP.CMPID,Mst_CMP.CMPName,"
tmp = tmp & " Mst_Godown.GodownNo,Mst_Godown.Address, Mst_Designation.Designation, Mst_Employee.EmployeeNo, Mst_Employee.EmployeeName, "
tmp = tmp & " Mst_AuditType.AuditType, Mst_AuditPoint.AuditPoint, Txn_AuditDetails.Observation, Txn_AuditDetails.Comment,"
tmp = tmp & " Txn_AuditDetails.Required, Txn_AuditDetails.Actual, Txn_AuditDetails.RequiredDate, Txn_AuditDetails.RequiredValue,"
tmp = tmp & " Txn_AuditDetails.StartDate , Txn_AuditDetails.CloseDate, Txn_AuditDetails.Flag"
tmp = tmp & " FROM Txn_Audit INNER JOIN "
tmp = tmp & " Txn_AuditDetails ON Txn_Audit.AuditID = Txn_AuditDetails.AuditID INNER JOIN"
tmp = tmp & " Mst_AuditPoint ON Txn_AuditDetails.AuditPointID = Mst_AuditPoint.AuditPointID INNER JOIN"
tmp = tmp & " Mst_AuditSegment ON Txn_Audit.SegmentTypeID = Mst_AuditSegment.SegmentTypeID INNER JOIN"
tmp = tmp & " Mst_AuditType ON Txn_Audit.AuditTypeID = Mst_AuditType.AuditTypeID INNER JOIN"
tmp = tmp & " Mst_Designation ON Txn_Audit.DesignationID = Mst_Designation.DesignationID INNER JOIN"
tmp = tmp & " Mst_Employee ON Txn_Audit.EmployeeID = Mst_Employee.EmployeeID INNER JOIN"
tmp = tmp & " Mst_Godown ON Txn_Audit.GodownID = Mst_Godown.GodownID INNER JOIN"
tmp = tmp & " Mst_CMP ON Mst_Godown.CMPID = Mst_CMP.CMPID INNER JOIN"
tmp = tmp & " Mst_Location ON Mst_CMP.LocationID  = Mst_Location.LocationID INNER JOIN"
tmp = tmp & " Mst_State ON Mst_Location.StateID = Mst_State.StateID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " WHERE AuditDate >= #" & Sdat & "# AND AuditDate <= #" & Edat & "#  "
Else
      tmp = tmp & " WHERE AuditDate >= '" & Sdat & "' AND AuditDate <= '" & Edat & "'  "
End If

If CmbLocation.Text <> "" Then
   tmp = tmp & " And Mst_Location.LocationName = '" & CmbLocation.Text & "' "
End If

If CmbCMP.Text <> "" Then
   tmp = tmp & " And Mst_CMP.CMPName = '" & CmbCMP.Text & "'"
End If

If CmbGodownName.Text <> "" Then
   tmp = tmp & " And Mst_Godown.GodownNo = '" & CmbGodownName.Text & "'"
End If

If CmbAuditType.Text <> "" Then
   tmp = tmp & " And Mst_AuditType.AuditType = '" & CmbAuditType.Text & "'  "
End If

If CmbAuditPoint.Text <> "" Then
   tmp = tmp & " And Mst_AuditPoint.AuditPoint = '" & CmbAuditPoint.Text & "'"
End If

If CmbStatus.Text <> "" Then
   If LCase(CmbStatus.Text) = "open" Then
      tmp = tmp & " And Txn_AuditDetails.CloseDate IS NULL "
   ElseIf LCase(CmbStatus.Text) = "close" Then
      tmp = tmp & " And Txn_AuditDetails.CloseDate IS NOT NULL"
   End If
End If
tmp = tmp & " Order by Mst_State.StateName,Mst_Location.LocationName,Mst_CMP.CMPName"


Call TmpTable

Call dispdata
'If TmpRs.RecordCount > 0 Then
'Else
'   MsgBox "No record found for given Date!!!"
'   Call Grid_Head
'   Exit Sub
'End If

If MSHFlexGrid1.Rows > 2 Then
   MSHFlexGrid1.FixedRows = 1
   MSHFlexGrid1.FixedCols = 4
End If
CmdExcel.Enabled = True
MsgBox "Done!!!"

End Sub
Public Sub dispdata()
Dim x As Double

tmp = "Select ECM.CMPID,EMP.EmployeeName +'~'+ D.Designation  as CM,EMP1.EmployeeName  +'~'+ D1.Designation  as CC,"
tmp = tmp & " EMP2.EmployeeName +'~'+ D2.Designation as AM ,EMP3.EmployeeName +'~'+ D3.Designation as RM,"
tmp = tmp & " EMP4.EmployeeName +'~'+ D4.Designation as Managment"
tmp = tmp & " From Mst_EmpCMPMapping ECM"
tmp = tmp & " Inner Join Mst_Employee EMP on ECM.EmployeeID = EMP.EmployeeID"
tmp = tmp & " Left Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Join Mst_employee EMP2 on EMP1.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Join Mst_employee EMP3 ON EMP2.DirectReportingTo = EMP3.EmployeeID"
tmp = tmp & " Left Join Mst_employee EMP4 ON EMP3.DirectReportingTo = EMP4.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation D on EMP.DesignationID = D.DesignationID"
tmp = tmp & " Inner Join Mst_Designation D1 on EMP1.DesignationID = D1.DesignationID"
tmp = tmp & " Inner Join Mst_Designation D2 on EMP2.DesignationID = D2.DesignationID"
tmp = tmp & " Inner Join Mst_Designation D3 on EMP3.DesignationID = D3.DesignationID"
tmp = tmp & " Inner Join Mst_Designation D4 on EMP4.DesignationID = D4.DesignationID"
tmp = tmp & " Where EMP.Flag = 'Y' "
tmp = tmp & " Group By ECM.CMPID,EMP.EmployeeName,EMP1.EmployeeName,EMP2.EmployeeName,EMP3.EmployeeName,EMP4.EmployeeName,"
tmp = tmp & " D.Designation , D1.Designation, D2.Designation, D3.Designation, D4.Designation"


Call Tmp1Table



x = 1

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'       For c = 0 To 24
'           MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'       Next
       
       MSHFlexGrid1.TextMatrix(x, 0) = IIf(IsNull(TmpRs!AuditID), "", TmpRs!AuditID)
       MSHFlexGrid1.TextMatrix(x, 1) = IIf(IsNull(TmpRs!AuditDate), "", TmpRs!AuditDate)
       MSHFlexGrid1.TextMatrix(x, 2) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(x, 3) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid1.TextMatrix(x, 4) = IIf(IsNull(TmpRs!SegmentType), "", TmpRs!SegmentType)
       MSHFlexGrid1.TextMatrix(x, 5) = IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID)
       MSHFlexGrid1.TextMatrix(x, 6) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
       MSHFlexGrid1.TextMatrix(x, 7) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
       MSHFlexGrid1.TextMatrix(x, 8) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
       MSHFlexGrid1.TextMatrix(x, 9) = IIf(IsNull(TmpRs!Designation), "", TmpRs!Designation)
       MSHFlexGrid1.TextMatrix(x, 10) = IIf(IsNull(TmpRs!EmployeeNo), "", TmpRs!EmployeeNo)
       MSHFlexGrid1.TextMatrix(x, 11) = IIf(IsNull(TmpRs!EmployeeName), "", TmpRs!EmployeeName)
       MSHFlexGrid1.TextMatrix(x, 12) = IIf(IsNull(TmpRs!AuditType), "", TmpRs!AuditType)
       MSHFlexGrid1.TextMatrix(x, 13) = IIf(IsNull(TmpRs!AuditPoint), "", TmpRs!AuditPoint)
       
       MSHFlexGrid1.TextMatrix(x, 14) = IIf(IsNull(TmpRs!Observation), "", TmpRs!Observation)
       
       If MSHFlexGrid1.TextMatrix(x, 14) <> "" Then
          strSplitString = Split(MSHFlexGrid1.TextMatrix(x, 14), "~")
          strObservation = strSplitString(0)
          MSHFlexGrid1.TextMatrix(x, 15) = strObservation
       Else
          MSHFlexGrid1.TextMatrix(x, 15) = MSHFlexGrid1.TextMatrix(x, 14)
       End If
              
       MSHFlexGrid1.TextMatrix(x, 16) = IIf(IsNull(TmpRs!Comment), "", TmpRs!Comment)
       
       If MSHFlexGrid1.TextMatrix(x, 16) <> "" Then
          strSplitString = Split(MSHFlexGrid1.TextMatrix(x, 16), "~")
          strComment = strSplitString(0)
          MSHFlexGrid1.TextMatrix(x, 17) = strComment
       Else
          MSHFlexGrid1.TextMatrix(x, 17) = MSHFlexGrid1.TextMatrix(x, 16)
       End If
       
       MSHFlexGrid1.TextMatrix(x, 18) = IIf(IsNull(TmpRs!Required), "", TmpRs!Required)
       MSHFlexGrid1.TextMatrix(x, 19) = IIf(IsNull(TmpRs!Actual), "", TmpRs!Actual)
       MSHFlexGrid1.TextMatrix(x, 20) = IIf(IsNull(TmpRs!RequiredDate), "", TmpRs!RequiredDate)
       MSHFlexGrid1.TextMatrix(x, 21) = IIf(IsNull(TmpRs!RequiredValue), "", TmpRs!RequiredValue)
       MSHFlexGrid1.TextMatrix(x, 22) = IIf(IsNull(TmpRs!StartDate), "", TmpRs!StartDate)
       MSHFlexGrid1.TextMatrix(x, 23) = IIf(IsNull(TmpRs!CloseDate), "", TmpRs!CloseDate)
       MSHFlexGrid1.TextMatrix(x, 24) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
       
       
       
'       Call AddGData(TmpRs!CMPID, x)
       
       
       
       TmpRs.MoveNext
   Next
End If

End Sub
'Sub AddGData(mCMPID_ As Double, x1 As Double)
'Dim strSplitString() As String
'Dim EmpName_, EmpDesig_ As String
'
'TmpRs1.MoveFirst
'TmpRs1.Filter = "CMPID = '" & mCMPID_ & "'"
'For i1 = 1 To TmpRs1.RecordCount
'
'    MSHFlexGrid1.TextMatrix(x1, 25) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
'    If IsNull(TmpRs1.Fields(2)) = False Then
'       strSplitString = Split(TmpRs1.Fields(2), "~")
'       EmpName_ = Trim(strSplitString(0))
'       EmpDesig_ = Trim(strSplitString(1))
'       If LCase(EmpDesig_) = "am" Then
'          MSHFlexGrid1.TextMatrix(x1, 27) = IIf(IsNull(TmpRs1.Fields(2)), "", TmpRs1.Fields(2))
'       ElseIf LCase(EmpDesig_) = "cc" Then
'          MSHFlexGrid1.TextMatrix(x1, 26) = IIf(IsNull(TmpRs1.Fields(2)), "", TmpRs1.Fields(2))
'       End If
'    End If
'
'    If IsNull(TmpRs1.Fields(3)) = False Then
'       strSplitString = Split(TmpRs1.Fields(3), "~")
'       EmpName_ = Trim(strSplitString(0))
'       EmpDesig_ = Trim(strSplitString(1))
'       If LCase(EmpDesig_) = "rm" Then
'          MSHFlexGrid1.TextMatrix(x1, 28) = IIf(IsNull(TmpRs1.Fields(3)), "", TmpRs1.Fields(3))
'       ElseIf LCase(EmpDesig_) = "am" Then
'          MSHFlexGrid1.TextMatrix(x1, 27) = IIf(IsNull(TmpRs1.Fields(3)), "", TmpRs1.Fields(3))
'       End If
'    End If
'
'    If IsNull(TmpRs1.Fields(4)) = False Then
'       strSplitString = Split(TmpRs1.Fields(4), "~")
'       EmpName_ = Trim(strSplitString(0))
'       EmpDesig_ = Trim(strSplitString(1))
'       If LCase(EmpDesig_) = "rm" Then
'          MSHFlexGrid1.TextMatrix(x1, 28) = IIf(IsNull(TmpRs1.Fields(4)), "", TmpRs1.Fields(4))
'       ElseIf LCase(EmpDesig_) = "managment" Then
'          MSHFlexGrid1.TextMatrix(x1, 29) = IIf(IsNull(TmpRs1.Fields(4)), "", TmpRs1.Fields(4))
'       End If
'    End If
'
'    If TmpRs1.RecordCount > 1 Then
'       x1 = MSHFlexGrid1.Rows
'       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'    End If
'
'    TmpRs1.MoveNext
'
'Next
'
'
'If TmpRs1.RecordCount > 1 Then
'   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
'End If
'
'TmpRs1.Filter = ""
'
'End Sub

Private Sub CmdExcel_Click()
If MSHFlexGrid1.Rows < 2 Then Exit Sub

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_Audit.xls")
'Call Gen_Create_Xls("Rpt_Audit.xls")
Call Xls_Detail_Rpt

End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
'Set oWB = oXL.Workbooks.Open(Pat & "\Rpt_Audit.xls")
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_Audit.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
Call TableMst_Location("")

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

Private Sub CmbLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocation")
End If
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text <> tmp Then
   CmbCMP.Clear
   CmbGodownName.Clear
   CmbCMP.Text = ""
   CmbGodownName.Text = ""
End If

If CmbLocation.Text = "" Then
   mLocationID = 0
   Exit Sub
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

Private Sub CmbGodownName_Gotfocus()
If CmbCMP.Text = "" Then
   MsgBox "Select CMP."
   CmbCMP.SetFocus
   Exit Sub
End If
'
temp = CmbGodownName.Text
Call TableMst_Godown("Where CMPID = " & mCMPID & " And CloseDate Is Null")
CmbGodownName.Clear

If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Godown found  !!!!! "
   CmbGodownName.SetFocus
   Exit Sub
End If
RsMst_Godown.MoveFirst
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownName.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownName.Text = temp

End Sub

Private Sub CmbGodownName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbGodownName")
End If
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
End Sub

Private Sub CmbAuditType_GotFocus()
tmp = CmbAuditType.Text
Call TableMst_AuditType("")
CmbAuditType.Clear
If RsMst_AuditType.RecordCount = 0 Then
   MsgBox "No Audit Type found"
   CmbAuditType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditType.RecordCount
    CmbAuditType.AddItem RsMst_AuditType!AuditType
    RsMst_AuditType.MoveNext
Next
CmbAuditType.Text = tmp
End Sub

Private Sub CmbAuditType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAuditType")
End If
End Sub

Private Sub CmbAuditType_LostFocus()
If CmbAuditType.Text = "" Then
   mAuditTypeID = Null
   Exit Sub
End If

RsMst_AuditType.MoveFirst
RsMst_AuditType.Find "AuditType = '" & CmbAuditType.Text & "'"

If RsMst_AuditType.EOF Then
   MsgBox "Invalid Audit type Selection "
   CmbAuditType.SetFocus
   Exit Sub
End If
mAuditTypeID = RsMst_AuditType!AuditTypeID
End Sub

Private Sub CmbAuditSegment_GotFocus()
tmp = CmbAuditSegment.Text
Call TableMst_AuditSegment("")
CmbAuditSegment.Clear
If RsMst_AuditSegment.RecordCount = 0 Then
   MsgBox "No Audit Segment found"
   CmbAuditSegment.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AuditSegment.RecordCount
    CmbAuditSegment.AddItem RsMst_AuditSegment!SegmentType
    RsMst_AuditSegment.MoveNext
Next
CmbAuditSegment.Text = tmp
End Sub

Private Sub CmbAuditSegment_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbAuditSegment")
End If
End Sub

Private Sub CmbAuditSegment_LostFocus()
If CmbAuditSegment.Text = "" Then
   mSegmentTypeID = Null
   Exit Sub
End If

RsMst_AuditSegment.MoveFirst
RsMst_AuditSegment.Find "SegmentType = '" & CmbAuditSegment.Text & "'"

If RsMst_AuditSegment.EOF Then
   MsgBox "Invalid Segment Selection "
   CmbAuditSegment.SetFocus
   Exit Sub
End If
mSegmentTypeID = RsMst_AuditSegment!SegmentTypeID

End Sub
