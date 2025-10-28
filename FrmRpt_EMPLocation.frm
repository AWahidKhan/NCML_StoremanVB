VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_EMPLocation 
   Caption         =   "Employee Location Connectivity"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   9090
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   135
      TabIndex        =   7
      Top             =   0
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
         Left            =   10575
         TabIndex        =   4
         Top             =   465
         Width           =   1215
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   4215
      End
      Begin VB.ComboBox CmbDesignation 
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
         Left            =   4440
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox CmbFlag 
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
         ItemData        =   "FrmRpt_EMPLocation.frx":0000
         Left            =   6960
         List            =   "FrmRpt_EMPLocation.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2265
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
         Left            =   9360
         TabIndex        =   3
         Top             =   465
         Width           =   1215
      End
      Begin VB.Label LblTotalRec 
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
         Left            =   14280
         TabIndex        =   12
         Top             =   525
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
         Left            =   11880
         TabIndex        =   11
         Top             =   525
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Employee Name"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   135
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Designation"
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
         Left            =   5100
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label11 
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
         Height          =   285
         Left            =   7785
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8415
      Left            =   120
      TabIndex        =   5
      Top             =   975
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8055
         Left            =   120
         TabIndex        =   6
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
End
Attribute VB_Name = "FrmRpt_EMPLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStatus As Variant
Dim mEmployeeID, mDesignationID As Variant

Private Sub CmbEmployeeID_GotFocus()
tmp = CmbEmployeeID.Text
Call TableMst_Employee("Where Flag = 'Y'")
CmbEmployeeID.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbEmployeeID.SetFocus
   'RsMst_Employee.MoveFirst
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
Private Sub CmbDesignation_GotFocus()
tmp = CmbDesignation.Text
CmbDesignation.Clear
Call TableMst_Designation
If RsMst_Designation.RecordCount = 0 Then
   MsgBox "No Record Found !!!"
   CmbDesignation.SetFocus
   Exit Sub
End If

For I = 1 To RsMst_Designation.RecordCount
    CmbDesignation.AddItem RsMst_Designation!Designation
    RsMst_Designation.MoveNext
Next

CmbDesignation.Text = tmp
End Sub

Private Sub CmbDesignation_LostFocus()
If CmbDesignation.Text = "" Then Exit Sub
RsMst_Designation.MoveFirst
RsMst_Designation.Find "Designation='" & CmbDesignation.Text & "'"
If RsMst_Designation.EOF Then
   MsgBox "Invalid Designation Selection"
   CmbDesignation.SetFocus
   Exit Sub
End If
mDesignationID = RsMst_Designation!DesignationID
End Sub

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then
   mStatus = Null
   Exit Sub
End If

If LCase(CmbFlag.Text) = "active" Then
   mStatus = "Y"
ElseIf LCase(CmbFlag.Text) = "de-active" Then
   mStatus = "N"
Else
   MsgBox "Invalid Selection!!!"
   CmbFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeLocationMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeLocationMapping.xls")
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

CmbFlag.Text = ""
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If

CmdGetReport.Enabled = True
Call Grid_Head
CmdExcel.Enabled = False
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
MSHFlexGrid1.Cols = 9

MSHFlexGrid1.TextMatrix(0, 0) = "Employee ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Employee No"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Designation ID"
MSHFlexGrid1.TextMatrix(0, 4) = "Designation"
MSHFlexGrid1.TextMatrix(0, 5) = "Status"
MSHFlexGrid1.TextMatrix(0, 6) = "Reporting To"
MSHFlexGrid1.TextMatrix(0, 7) = "Status"
MSHFlexGrid1.TextMatrix(0, 8) = "Location"


MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 2200
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 0
MSHFlexGrid1.ColWidth(4) = 2400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 3000
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1400


End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
wc = ""
Call Grid_Head
MsgBox "Wait till next message !!!"

If CmbEmployeeID.Text <> "" Then
   wc = GenWc(wc, mEmployeeID, "ME.EmployeeID", "N", "N")
End If

If CmbDesignation.Text <> "" Then
   wc = GenWc(wc, mDesignationID, "ME.DesignationID", "N", "N")
End If

If CmbFlag.Text <> "" Then
   wc = GenWc(wc, mStatus, "ME.Flag", "S", "N")
End If

tmp = "Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,MD.DesignationID,MD.Designation,ME.Flag," & _
      " ME1.EmployeeName +'~'+ ME1.EmployeeNo as ReportingTo, ME1.Flag " & _
      " From Mst_Employee ME " & _
      " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID " & _
      " Left Outer Join Mst_Employee ME1 On ME.DirectReportingTo = ME1.EmployeeID  "
'tmp = tmp & Wc
tmp1 = " Order By MD.DesignationID "

tmp = tmp & " " & wc & " " & tmp1

Call TmpTable

Call dispdata

If MSHFlexGrid1.Rows > 1 Then
   MSHFlexGrid1.FixedRows = 1
End If

LblTotalRec.Caption = MSHFlexGrid1.Rows - 1

MsgBox "Done!!!"

End Sub

Public Sub dispdata()
Dim x As Double
x = 0
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       x = MSHFlexGrid1.Rows - 1
       For c = 0 To 7
           MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       If TmpRs.Fields(5) = "Y" Then
          MSHFlexGrid1.TextMatrix(x, 5) = "Active"
       Else
          MSHFlexGrid1.TextMatrix(x, 5) = "De-Active"
       End If
       If TmpRs.Fields(7) = "Y" Then
          MSHFlexGrid1.TextMatrix(x, 7) = "Active"
       ElseIf TmpRs.Fields(7) = "N" Then
          MSHFlexGrid1.TextMatrix(x, 7) = "De-Active"
       Else
          MSHFlexGrid1.TextMatrix(x, 7) = ""
       End If
       Call AddGData(TmpRs!EmployeeID, TmpRs!Designation, x)
       'x = MSHFlexGrid1.Rows - 1
       TmpRs.MoveNext
   Next
'   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If

End Sub
Sub AddGData(mEmployeeID_ As Double, mDesignation_ As Variant, x1 As Double)
Dim mEmployeeIDDummy_ As Variant

mEmployeeIDDummy_ = mEmployeeID_
If LCase(mDesignation_) = "qqs" Or LCase(mDesignation_) = "drc" Then
   tmp = "Select DirectReportingTo From Mst_Employee Where EmployeeID = " & mEmployeeIDDummy_ & ""
   Call Tmp1Table
   If TmpRs1.RecordCount <> 0 Then
      mEmployeeIDDummy_ = IIf(IsNull(TmpRs1!DirectReportingTo), 0, TmpRs1!DirectReportingTo)
   End If
End If
tmp = "Select ECM.LocationID,ML.LocationName" & _
      " From Mst_EMPCMPMapping ECM " & _
      " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID " & _
      " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID " & _
      " Left Outer Join Mst_Employee EMP6 On EMP.ConnectedToRAG = EMP6.EmployeeID " & _
      " Left Outer Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
      " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
      " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
      " Where (EMP.EmployeeID = " & mEmployeeIDDummy_ & " Or EMP1.EmployeeID = " & mEmployeeIDDummy_ & " Or EMP2.EmployeeID = " & mEmployeeIDDummy_ & " Or EMP3.EmployeeID = " & mEmployeeIDDummy_ & " Or EMP4.EmployeeID = " & mEmployeeIDDummy_ & " Or EMP6.EmployeeID = " & mEmployeeIDDummy_ & ") " & _
      " Group By ECM.LocationID,ML.LocationName " & _
      " Order By ECM.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then

   For i1 = 1 To TmpRs1.RecordCount
       For r = 0 To 7
           MSHFlexGrid1.TextMatrix(x1, r) = IIf(IsNull(TmpRs.Fields(r)), "", TmpRs.Fields(r))
       Next
       If TmpRs.Fields(5) = "Y" Then
          MSHFlexGrid1.TextMatrix(x1, 5) = "Active"
       Else
          MSHFlexGrid1.TextMatrix(x1, 5) = "De-Active"
       End If
       If TmpRs.Fields(7) = "Y" Then
          MSHFlexGrid1.TextMatrix(x1, 7) = "Active"
       ElseIf TmpRs.Fields(7) = "N" Then
          MSHFlexGrid1.TextMatrix(x1, 7) = "De-Active"
       Else
          MSHFlexGrid1.TextMatrix(x1, 7) = ""
       End If
       MSHFlexGrid1.TextMatrix(x1, 8) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
       TmpRs1.MoveNext
       If TmpRs1.RecordCount > 1 Then
          x1 = MSHFlexGrid1.Rows
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       End If
   Next
   If TmpRs1.RecordCount > 1 Then
     MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
   End If
End If

TmpRs1.Close



End Sub

