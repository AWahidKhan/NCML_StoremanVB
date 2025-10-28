VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_AuditorGodownMIS 
   Caption         =   "Auditor Godown MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   8055
      Left            =   120
      TabIndex        =   15
      Top             =   1470
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   16
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   15015
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
         Left            =   8775
         TabIndex        =   7
         Top             =   353
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
         Left            =   7560
         TabIndex        =   6
         Top             =   353
         Width           =   1215
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
         TabIndex        =   5
         Top             =   353
         Width           =   1215
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
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   360
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
         Format          =   61538305
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         Top             =   360
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
         Format          =   61538305
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         Left            =   13800
         TabIndex        =   14
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
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
         Left            =   11280
         TabIndex        =   13
         Top             =   413
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
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
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
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
         Left            =   2040
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.Label LblSelection 
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
         Left            =   4860
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.OptionButton Option2 
         Caption         =   "Godown not Audited"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Auditor wise MIS"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_AuditorGodownMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSplitString() As String
Dim EmpNo, EmpName As Variant

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Option1.Value = False And Option2.Value = False Then Exit Sub

CmbLocation.Clear
If Option1.Value = True Then
   Call TableMst_Employee("Where DesignationID = 7")
   If RsMst_Employee.RecordCount = 0 Then
      MsgBox "No Auditor found !!!! "
      Exit Sub
   End If
   For I = 1 To RsMst_Employee.RecordCount
       CmbLocation.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
       RsMst_Employee.MoveNext
   Next
End If

If Option2.Value = True Then
   If Gen_WcLocation <> "" Then
      Call TableMst_Location(" Where " & Gen_WcLocation)
   Else
      Call TableMst_Location
   End If
   If RsMst_Location.RecordCount = 0 Then
      MsgBox "No Location found !!!! "
      Exit Sub
   End If
   For I = 1 To RsMst_Location.RecordCount
       CmbLocation.AddItem RsMst_Location!LocationName
       RsMst_Location.MoveNext
   Next
End If

CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()



If CmbLocation.Text = "" Then
   Exit Sub
End If

If Option1.Value = True Then
   strSplitString = Split(CmbLocation, "~")
   EmpName = Trim(strSplitString(0))
   EmpNo = Trim(strSplitString(1))

   RsMst_Employee.MoveFirst
   RsMst_Employee.Find "EmployeeNo = '" & EmpNo & "'"

   If RsMst_Employee.EOF Then
      MsgBox "Invalid Selection "
      CmbLocation.SetFocus
      Exit Sub
   End If
End If

If Option2.Value = True Then
   RsMst_Location.MoveFirst
   RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
   If RsMst_Location.EOF Then
      MsgBox "Invalid Location  selection "
      CmbLocation.SetFocus
      Exit Sub
   End If
End If


End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Label22.Caption = ""
CmdExcel.Enabled = False
End Sub

Private Sub CmdGetReport_Click()
Dim wc, DispFlag, Filt, Sdat, Edat As Variant
Call Grid_Head

If Option2.Value = True Then
   If TxtFromDate.Value > TxtToDate.Value Then
      MsgBox "'From Date' must be less than 'To Date' always!!!"
      TxtFromDate.SetFocus
      Exit Sub
   End If
End If

If Option1.Value = True Then
   tmp = " SELECT Mst_State.StateName, Mst_Location.LocationName,Mst_CMP.CMPName,Max(Txn_Audit.AuditDate) as AuditDate,Mst_Employee.EmployeeNo, "
   tmp = tmp & " Mst_Employee.EmployeeName,Mst_Designation.Designation, "
   tmp = tmp & " COUNT(Mst_Godown.GodownNo) As [No of Godown] "
   tmp = tmp & " From Txn_Audit "
   tmp = tmp & " INNER JOIN  Mst_Employee ON Txn_Audit.EmployeeID = Mst_Employee.EmployeeID "
   tmp = tmp & " INNER JOIN  Mst_Employee EMP ON Mst_Employee.DirectReportingTo = EMP.EmployeeID "
   tmp = tmp & " INNER JOIN  Mst_Godown ON Txn_Audit.GodownID = Mst_Godown.GodownID "
   tmp = tmp & " INNER JOIN  Mst_CMP ON Mst_Godown.CMPID = Mst_CMP.CMPID "
   tmp = tmp & " INNER JOIN  Mst_Location ON Mst_CMP.LocationID = Mst_Location.LocationID "
   tmp = tmp & " INNER JOIN  Mst_State ON Mst_Location.StateID = Mst_State.StateID  "
   tmp = tmp & " INNER JOIN  Mst_Designation ON Mst_Employee.DesignationID = Mst_Designation.DesignationID "

ElseIf Option2.Value = True Then
   tmp = " Select MS.StateName,ML.LocationName,MC.CMPName,MG.GodownNo,MG.Address,MC.CMPID,MG.GodownID "
   tmp = tmp & " From Mst_Godown MG  "
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID  "
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID  "
   tmp = tmp & " INNER JOIN Mst_State MS ON ML.StateID = MS.StateID  "
   tmp = tmp & " Where "
   If Gen_DBType = "MDB" Then
      tmp = tmp & " MG.GodownID Not in (Select GodownID From Txn_Audit Where AuditTypeID = 1 And AuditDate Between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " MG.GodownID Not in (Select GodownID From Txn_Audit Where AuditTypeID = 1 And AuditDate Between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "') "
   End If
   tmp = tmp & " And MG.CloseDate is Null  "

Else
   MsgBox " Please Select atleast one option from above!!!"
   Exit Sub
End If

wc = ""

If CmbLocation.Text <> "" Then
   If Option1.Value = True Then
      wc = " Where Mst_Employee.EmployeeNo =  '" & EmpNo & "' "
   End If
   If Option2.Value = True Then
      wc = " And ML.LocationName =  '" & CmbLocation.Text & "' "
   End If
Else
   If Option2.Value = True Then
      If Gen_WcLocation <> "" Then
         wc = " And ML." & Gen_WcLocation
      End If
   End If
End If
tmp1 = ""
If Option1.Value = True Then
   tmp1 = " GROUP BY Mst_State.StateName, Mst_Location.LocationName,Mst_CMP.CMPName, " & _
         " Mst_Employee.EmployeeNo , Mst_Employee.EmployeeName, Mst_Designation.Designation "
End If

If Option2.Value = True Then
   tmp1 = " Order By MS.StateName,ML.LocationName,MC.CMPName "
End If

tmp = tmp & wc & tmp1

Call TmpTable

MsgBox "Wait till next message"

Call dispdata
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

MsgBox "Done"

If MSHFlexGrid1.Rows > 2 Then
   MSHFlexGrid1.FixedRows = 2
End If

CmdExcel.Enabled = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True
Call Grid_Head
CmdExcel.Enabled = False
Option1.Value = False
Option2.Value = False
Call Option1_Click

End Sub
Public Sub dispdata()
Dim mEmpNames, mDesig As Variant
Dim x As Variant
Dim mRAGEmpNames As String
x = 0
If TmpRs.RecordCount > 0 Then
   If Option2.Value = True Then
      tmp = " Select MEC.EmployeeID ,ME.EmployeeName,MDes.Designation,ME1.EmployeeID as RAGEmployeeID,ME1.EmployeeName As RAGEmployeeName,MEC.CMPID "
      tmp = tmp & " From Mst_CMP MC "
      tmp = tmp & " Left JOIN Mst_EmpCMPMapping MEC ON MC.CMPID = MEC.CMPID"
      tmp = tmp & " Left join Mst_Employee ME on MEC.EmployeeID=ME.EmployeeID"
      tmp = tmp & " Left join Mst_Employee ME1 on ME.ConnectedToRAG=ME1.EmployeeID"
      tmp = tmp & " INNER JOIN Mst_Designation MDes ON ME.DesignationID = MDes.DesignationID"
      tmp = tmp & " Where ME.Flag='Y'"
         
      Call Tmp2Table
      
      tmp = "Select Max(AuditDate) As AuditDate ,GodownID From Txn_Audit Group By GodownID"
      
      Call Tmp3Table
   End If
   
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       If Option2.Value = True Then
          For c = 0 To 4
              MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
          Next
          TmpRs2.Filter = "CMPID=" & TmpRs!CMPID & "  "
          mEmpNames = ""
          mRAGEmpNames = ""
          mDesig = ""
          For r1 = 1 To TmpRs2.RecordCount
              If mEmpNames <> "" Then
                 mEmpNames = mEmpNames & " ~ " & TmpRs2!EmployeeName
              Else
                 mEmpNames = TmpRs2!EmployeeName
              End If
              If mRAGEmpNames <> "" Then
                 mRAGEmpNames = mRAGEmpNames & " ~ " & IIf(IsNull(TmpRs2!RAGEmployeeName), "", TmpRs2!RAGEmployeeName)
              Else
                 mRAGEmpNames = IIf(IsNull(TmpRs2!RAGEmployeeName), "", TmpRs2!RAGEmployeeName)
              End If
              mDesig = TmpRs2!Designation
              TmpRs2.MoveNext
          Next
          MSHFlexGrid1.TextMatrix(x, 5) = mEmpNames
          MSHFlexGrid1.TextMatrix(x, 6) = mDesig
          MSHFlexGrid1.TextMatrix(x, 7) = mRAGEmpNames
          TmpRs2.Filter = ""
          If TmpRs3.RecordCount > 0 Then
             TmpRs3.MoveFirst
          End If
          TmpRs3.Find "GodownID= " & TmpRs!GodownID
          If Not TmpRs3.EOF Then
             MSHFlexGrid1.TextMatrix(x, 8) = Format(TmpRs3!AuditDate, "dd-MMM-yyyy")
          End If
       Else
          For c = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
          Next
          
       End If
       If Option1.Value = True Then
          MSHFlexGrid1.TextMatrix(x, 3) = Format(CDate(MSHFlexGrid1.TextMatrix(x, 3)), "dd-MMM-yyyy")
          MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(x, 7))
       End If
       TmpRs.MoveNext
    Next
End If
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
If Option1 = True Then
   MSHFlexGrid1.Cols = 8
ElseIf Option2 = True Then
   MSHFlexGrid1.Cols = 9
Else
   MSHFlexGrid1.Cols = 4
End If

MSHFlexGrid1.WordWrap = True

If Option1.Value = True Or Option2.Value = True Then
   MSHFlexGrid1.TextMatrix(1, 0) = "State Name"
   MSHFlexGrid1.TextMatrix(1, 1) = "Location Name"
   MSHFlexGrid1.TextMatrix(1, 2) = "CMP Name"
End If

If Option1.Value = True Then
   MSHFlexGrid1.TextMatrix(1, 3) = "Audit Date"
   MSHFlexGrid1.TextMatrix(1, 4) = "RAG (Employee no)"
   MSHFlexGrid1.TextMatrix(1, 5) = "RAG Employee Name"
   MSHFlexGrid1.TextMatrix(1, 6) = "Designation"
   MSHFlexGrid1.TextMatrix(1, 7) = "No Of Godown"
ElseIf Option2.Value = True Then
   MSHFlexGrid1.TextMatrix(1, 3) = "Godown No"
   MSHFlexGrid1.TextMatrix(1, 4) = "Godown Address"
   MSHFlexGrid1.TextMatrix(1, 5) = "Employee Names"
   MSHFlexGrid1.TextMatrix(1, 6) = "Employee Designation"
   MSHFlexGrid1.TextMatrix(1, 7) = "RAG Employee Names"
   MSHFlexGrid1.TextMatrix(1, 8) = "Last Audit Date"
End If
MSHFlexGrid1.ColWidth(0) = 2500
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 3000
If Option1.Value = True Then
   MSHFlexGrid1.ColWidth(3) = 1400
   MSHFlexGrid1.ColWidth(4) = 1400
   MSHFlexGrid1.ColWidth(5) = 1800
   MSHFlexGrid1.ColWidth(6) = 1000
   MSHFlexGrid1.ColWidth(7) = 1000
End If
If Option2.Value = True Then
   MSHFlexGrid1.ColWidth(0) = 1800
   MSHFlexGrid1.ColWidth(1) = 1800
   MSHFlexGrid1.ColWidth(2) = 1800
   MSHFlexGrid1.ColWidth(3) = 1000
   MSHFlexGrid1.ColWidth(4) = 3400
   MSHFlexGrid1.ColWidth(5) = 2400
   MSHFlexGrid1.ColWidth(6) = 1400
   MSHFlexGrid1.ColWidth(7) = 2400
   MSHFlexGrid1.ColWidth(8) = 1400
End If


MSHFlexGrid1.RowHeight(1) = 700

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Rpt_AuditorGodownMIS.xls")
'Call Gen_Create_Xls("Rpt_AuditorGodownMIS.xls")
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
'Set oWB = oXL.Workbooks.Open(Pat & "\" & "Rpt_AuditorGodownMIS.xls")
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Rpt_AuditorGodownMIS.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
   LblSelection.Caption = "Auditor"
ElseIf Option2.Value = True Then
   TxtFromDate.Enabled = True
   TxtToDate.Enabled = True
   TxtFromDate.Value = Date - 30
   TxtToDate.Value = Date
   LblSelection.Caption = "Location"
Else
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
   LblSelection.Caption = "Location"
End If
CmbLocation.Text = ""
End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub



