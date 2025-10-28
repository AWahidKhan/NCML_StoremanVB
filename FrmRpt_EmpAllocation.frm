VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_EmpAllocation 
   Caption         =   "Employee Allocation"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   9555
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   7
      Top             =   975
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   135
      TabIndex        =   4
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbReportingTo 
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
         Width           =   4215
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
         Left            =   9990
         TabIndex        =   3
         Top             =   435
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
         Left            =   8790
         TabIndex        =   2
         Top             =   435
         Width           =   1215
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
         ItemData        =   "FrmRpt_EmpAllocation.frx":0000
         Left            =   14520
         List            =   "FrmRpt_EmpAllocation.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Reporting To"
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
         Left            =   5940
         TabIndex        =   12
         Top             =   135
         Width           =   1215
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
         Left            =   11640
         TabIndex        =   5
         Top             =   480
         Width           =   2415
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
         TabIndex        =   6
         Top             =   480
         Width           =   90
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
         TabIndex        =   11
         Top             =   135
         Width           =   1575
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
         Left            =   13920
         TabIndex        =   10
         Top             =   450
         Visible         =   0   'False
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmRpt_EmpAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStatus, mEmployeeID, mReportingToID As Variant


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
Private Sub CmbReportingTo_GotFocus()
tmp1 = CmbReportingTo.Text

tmp = "Select EmployeeID,EmployeeNo,EmployeeName From Mst_Employee"
tmp = tmp & " Where EmployeeID in("
tmp = tmp & "    Select Distinct DirectReportingTo From Mst_Employee)"
tmp = tmp & " And Flag='Y'"

Call Tmp4Table

CmbReportingTo.Clear
If TmpRs4.RecordCount = 0 Then
   MsgBox "No Direct Reporting To Found!!!"
   CmbReportingTo.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs4.RecordCount
    CmbReportingTo.AddItem TmpRs4!EmployeeName & "~" & TmpRs4!EmployeeNo
    TmpRs4.MoveNext
Next

CmbReportingTo.Text = tmp1

End Sub

Private Sub CmbReportingTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbReportingTo.Text = "" Then
   mReportingToID = Null
   Exit Sub
End If

strSplitString = Split(CmbReportingTo, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


TmpRs4.MoveFirst
TmpRs4.Find "EmployeeNo = '" & EmpNo_ & "'"

If TmpRs4.EOF Then
   MsgBox "Invalid Selection "
   CmbReportingTo.SetFocus
   Exit Sub
End If
mReportingToID = TmpRs4!EmployeeID
End Sub

Private Sub CmdGetReport_Click()
Dim Wc As String

Wc = ""


Call Grid_Head

tmp = "Select distinct EMP.EmployeeNo,EMP.EmployeeName,EMP.Flag,Emp.EmployeeID"
tmp = tmp & "     From Mst_Employee EMP where EMP.Flag = 'Y' "
tmp = tmp & "     order by EMP.EmployeeName"
Call Tmp3Table

''''Add Filter Here
If CmbEmployeeID.Text <> "" Then
   Wc = GenWc(Wc, mEmployeeID, "EMP.EmployeeID", "N", "N")
End If



If CmbReportingTo.Text <> "" Then
   Wc = GenWc(Wc, mReportingToID, "EMP.DirectReportingTo", "N", "N")
End If

tmp = "Select distinct  EMP.EmployeeNo,EMP.EmployeeName,EMP.Flag,EMP.DirectReportingTo "
tmp = tmp & "     From Mst_EMPCMPMapping ECM"
tmp = tmp & "     inner Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID"
tmp = tmp & "     Inner Join Mst_Designation MD On EMP.DesignationID = MD.DesignationID  "
If Wc <> "" Then
   tmp = tmp & Wc & " And EMP.Flag = 'Y' "
Else
   tmp = tmp & " where EMP.Flag = 'Y'"
End If
tmp = tmp & "     order by EMP.EmployeeName"

Call TmpTable

tmp = "Select distinct EMP.EmployeeName,EMP1.EmployeeName + ' ~ ' + EMP1.EmployeeNo 'EMPCon' ,MD1.Designation,EMP1.Flag  "
tmp = tmp & "     From Mst_EMPCMPMapping ECM"
tmp = tmp & "     inner Join Mst_Location ML On ECM.LocationID = ML.LocationID"
tmp = tmp & "     inner Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID and EMP.Flag = 'Y'"
tmp = tmp & "     Left Outer Join Mst_Employee EMP1 On EMP.EmployeeID  = EMP1.DirectReportingTo and EMP1.Flag = 'Y'"
tmp = tmp & "     Inner Join Mst_Designation MD On EMP.DesignationID = MD.DesignationID"
tmp = tmp & "     Inner Join Mst_Designation MD1 On EMP1.DesignationID = MD1.DesignationID where EMP1.Flag = 'Y'"
tmp = tmp & "     order by EMP.EmployeeName"

Call Tmp1Table

tmp = "Select distinct EMP.EmployeeName,ML.LocationName"
tmp = tmp & "     From Mst_EMPCMPMapping ECM"
tmp = tmp & "     inner Join Mst_Location ML On ECM.LocationID = ML.LocationID"
tmp = tmp & "     inner Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID where EMP.Flag = 'Y'"
tmp = tmp & "     order by EMP.EmployeeName"
Call Tmp2Table


For o = 1 To TmpRs.RecordCount
    TmpRs1.Filter = "EmployeeName = '" & TmpRs!EmployeeName & "'"
    TmpRs2.Filter = "EmployeeName = '" & TmpRs!EmployeeName & "'"
    TmpRs3.Filter = "EmployeeID = " & TmpRs!DirectReportingTo & ""
    
    Call dispdata
    TmpRs.MoveNext
Next
Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
End Sub
Public Sub dispdata()
Dim MaxR, x, mFlag As Double
x = 1
MaxR = 0
If TmpRs1.RecordCount > 0 Then
   MaxR = TmpRs1.RecordCount
End If

If TmpRs2.RecordCount > MaxR Then
   MaxR = TmpRs2.RecordCount
End If
mFlag = 0

If MaxR > 0 Then
   For I = 1 To MaxR
       x = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(x, 0) = x
       
       
       If mFlag = 0 Then
          MSHFlexGrid1.TextMatrix(x, 1) = IIf(IsNull(TmpRs.Fields(0)), "", TmpRs.Fields(0))
          MSHFlexGrid1.TextMatrix(x, 2) = IIf(IsNull(TmpRs.Fields(1)), "", TmpRs.Fields(1))
          If TmpRs.Fields(2) = "Y" Then
             MSHFlexGrid1.TextMatrix(x, 3) = "Active"
          Else
             MSHFlexGrid1.TextMatrix(x, 3) = "De-Active"
          End If
          If Not TmpRs3.EOF Then
             tmp = IIf(IsNull(TmpRs3.Fields(1)), "", TmpRs3.Fields(1))
             tmp = tmp & " ~ " & IIf(IsNull(TmpRs3.Fields(0)), "", TmpRs3.Fields(0))
             MSHFlexGrid1.TextMatrix(x, 5) = tmp
          End If
       
       End If
       
       If Not TmpRs2.EOF Then
          MSHFlexGrid1.TextMatrix(x, 4) = IIf(IsNull(TmpRs2.Fields(1)), "", TmpRs2.Fields(1))
       End If
       
       
       
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(x, 6) = IIf(IsNull(TmpRs1.Fields(1)), "", TmpRs1.Fields(1))
          MSHFlexGrid1.TextMatrix(x, 7) = IIf(IsNull(TmpRs1.Fields(2)), "", TmpRs1.Fields(2))
          If TmpRs1.Fields(3) = "Y" Then
             MSHFlexGrid1.TextMatrix(x, 8) = "Active"
          Else
             MSHFlexGrid1.TextMatrix(x, 8) = "De-Active"
          End If
       
       End If
       
       
       If Not TmpRs1.EOF Then
          TmpRs1.MoveNext
       End If
       If Not TmpRs2.EOF Then
          TmpRs2.MoveNext
       End If
       mFlag = mFlag + 1
       
   Next
   
Else
   Call Grid_Head
End If

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeAllocation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeAllocation.xls")
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
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If

Call Grid_Head
CmdExcel.Enabled = False
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9

MSHFlexGrid1.TextMatrix(0, 0) = "Sr_No"
MSHFlexGrid1.TextMatrix(0, 1) = "Employee No"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee Name "
MSHFlexGrid1.TextMatrix(0, 3) = "Status"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "Reporting To "
MSHFlexGrid1.TextMatrix(0, 6) = "CMP-Staff Name"
MSHFlexGrid1.TextMatrix(0, 7) = "CMP-Staff Designation"
MSHFlexGrid1.TextMatrix(0, 8) = "CMP-Staff Status"


MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1800
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 2400
MSHFlexGrid1.ColWidth(5) = 3000
MSHFlexGrid1.ColWidth(6) = 3200
MSHFlexGrid1.ColWidth(7) = 2000

End Sub


