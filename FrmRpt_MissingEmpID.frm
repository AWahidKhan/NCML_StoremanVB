VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_MissingEmpID 
   Caption         =   "Missing Employee ID"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   12765
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
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
         Height          =   495
         Left            =   3150
         TabIndex        =   5
         Top             =   360
         Width           =   1455
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
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
         Left            =   1695
         TabIndex        =   3
         Top             =   360
         Width           =   1455
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
         Left            =   5760
         TabIndex        =   7
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
         Left            =   8280
         TabIndex        =   6
         Top             =   480
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   1110
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14420
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_MissingEmpID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClear_Click()
Call Grid_Head1
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_MissingEmpIDs.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_MissingEmpIDs.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head1
Label2.Caption = ""
'tmp = "Select ME.EmployeeNo,ME.EmployeeName,MD.Designation,ME.EmployeeEmailID,ME.MobileNo " & _
'      " From Mst_Employee ME " & _
'      " inner join Mst_Designation MD on ME.DesignationID=MD.DesignationID " & _
'      " Where EmployeeNo Not in(select Employee_ID from Mst_Users) And Flag='Y' Order By EmployeeName"

tmp = "select MU.EmployeeNo,MU.EmployeeName,MUR.GroupName,ML.LocationName,MU.EmailID,MU.ContactNo from Mst_Users MU"
tmp = tmp & " inner join Mst_Location ML on MU.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_UserRole MUR on MU.UserGroupID=MUR.UserGroupID"
tmp = tmp & " where MU.EmployeeNo Not in (select EmployeeNo from Mst_Employee) And (MU.Flag = 'Y' Or MU.Flag is null) "
tmp = tmp & " Order By MU.EmployeeName"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       TmpRs.MoveNext
   Next
   CmdExcel.Enabled = True
   Label2.Caption = TmpRs.RecordCount
   Label2.Refresh
   MsgBox "Done!!"
Else
   MsgBox "Record Not Found."
End If

End Sub

Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 6
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 1
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = " Employee ID "
MSHFlexGrid1.TextMatrix(0, 1) = " Employee Name"
MSHFlexGrid1.TextMatrix(0, 2) = " User Role"
MSHFlexGrid1.TextMatrix(0, 3) = " Location"
MSHFlexGrid1.TextMatrix(0, 4) = " Email ID"
MSHFlexGrid1.TextMatrix(0, 5) = " Contact No"

MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 2600
MSHFlexGrid1.ColWidth(2) = 2200
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 2600
MSHFlexGrid1.ColWidth(5) = 1800

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

Call Grid_Head1
CmdExcel.Enabled = False
End Sub
