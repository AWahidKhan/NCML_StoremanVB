VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_EmpAssetMap 
   Caption         =   "Employee Asset  Report"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11490
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8340
      Left            =   120
      TabIndex        =   9
      Top             =   1515
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13996
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbAssetType 
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
         Left            =   8976
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   3015
      End
      Begin VB.ComboBox CmbAssetStatus 
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
         Left            =   12060
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   2505
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
         ItemData        =   "FrmRpt_EmpAssetMap.frx":0000
         Left            =   6764
         List            =   "FrmRpt_EmpAssetMap.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   2145
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
         Left            =   4402
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   2295
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
         Top             =   585
         Width           =   4215
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
         Left            =   11895
         TabIndex        =   6
         Top             =   1050
         Width           =   1335
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
         Left            =   10560
         TabIndex        =   5
         Top             =   1050
         Width           =   1335
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
         Left            =   13230
         TabIndex        =   7
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Asset Type"
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
         Left            =   9936
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Asset Status"
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
         Left            =   12705
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Employee Status"
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
         Left            =   7049
         TabIndex        =   15
         Top             =   240
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
         Left            =   5002
         TabIndex        =   14
         Top             =   240
         Width           =   1095
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
         TabIndex        =   13
         Top             =   240
         Width           =   1575
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
         Left            =   9240
         TabIndex        =   12
         Top             =   1110
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
         Left            =   6720
         TabIndex        =   11
         Top             =   1110
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_EmpAssetMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mEmployeeID, mDesignationID, mStatus As Variant
Dim mAssetStatusID, mAssetTypeId As Variant
Private Sub CmbAssetType_GotFocus()

tmp = CmbAssetType.Text
Call TableMst_AssetType("")
CmbAssetType.Clear
If RsMst_AssetType.RecordCount = 0 Then
   MsgBox "No Asset found"
   CmbAssetType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AssetType.RecordCount
    CmbAssetType.AddItem RsMst_AssetType!AssetType
    RsMst_AssetType.MoveNext
Next
CmbAssetType.Text = tmp

End Sub
Private Sub CmbAssetType_LostFocus()

If CmbAssetType.Text = "" Then
   mAssetTypeId = Null
   Exit Sub
End If

RsMst_AssetType.MoveFirst
RsMst_AssetType.Find "AssetType = '" & CmbAssetType.Text & "'"

If RsMst_AssetType.EOF Then
   MsgBox "Invalid Asset  selection "
   CmbAssetType.SetFocus
   Exit Sub
End If
mAssetTypeId = RsMst_AssetType.Fields(0)

End Sub
Private Sub CmbAssetStatus_GotFocus()

tmp = CmbAssetStatus.Text
Call TableMst_AssetStatus("")
CmbAssetStatus.Clear
If RsMst_AssetStatus.RecordCount = 0 Then
   MsgBox "No Asset Status found"
   CmbAssetStatus.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_AssetStatus.RecordCount
    CmbAssetStatus.AddItem RsMst_AssetStatus!AssetStatus
    RsMst_AssetStatus.MoveNext
Next
CmbAssetStatus.Text = tmp

End Sub
Private Sub CmbAssetStatus_LostFocus()
If CmbAssetStatus.Text = "" Then
   mAssetStatusID = Null
   Exit Sub
End If

RsMst_AssetStatus.MoveFirst
RsMst_AssetStatus.Find "AssetStatus = '" & CmbAssetStatus.Text & "'"

If RsMst_AssetStatus.EOF Then
   MsgBox "Invalid Status selection "
   CmbAssetStatus.SetFocus
   Exit Sub
End If
mAssetStatusID = RsMst_AssetStatus.Fields(0)

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


Private Sub CmdClear_Click()
Call Grid_Head
CmdExcel.Enabled = False
CmbEmployeeID.Text = ""
CmbDesignation.Text = ""
CmbFlag.Text = ""
CmbAssetStatus.Text = ""
CmbAssetType.Text = ""
Label2.Caption = "-"
End Sub

Private Sub CmdGetReport_Click()
Dim Wc As String
Wc = ""

Call Grid_Head

MsgBox "Wait till next message!!!"

If CmbEmployeeID.Text <> "" Then
   Wc = GenWc(Wc, mEmployeeID, "EMP.EmployeeID", "N", "N")
End If

If CmbDesignation.Text <> "" Then
   Wc = GenWc(Wc, mDesignationID, "D.DesignationID", "N", "N")
End If

If CmbFlag.Text <> "" Then
   Wc = GenWc(Wc, mStatus, "EMP.Flag", "S", "N")
End If

If CmbAssetType.Text <> "" Then
   Wc = GenWc(Wc, mAssetTypeId, "MAT.AssetsID", "N", "N")
End If

If CmbAssetStatus.Text <> "" Then
   Wc = GenWc(Wc, mAssetStatusID, "MAS.AssetsStatusID", "N", "N")
End If

tmp = "Select EMP.EmployeeID,EMP.EmployeeName+'~'+EMP.EmployeeNo as Employee,D.Designation,EMP.Flag,MAT.AssetType, " & _
      " EAM.AssetQty , EAM.Description, MAS.AssetStatus, EAM.VerifiedDate,EAM.Remarks,EAM.AMCExpDate,EAM.AMCAmount " & _
      " From Mst_EmpAssetMap EAM " & _
      " Left Join Mst_Employee EMP on EAM.EmployeeID = EMP.EmployeeID " & _
      " Inner Join Mst_Designation D on EMP.DesignationID = D.DesignationID " & _
      " Left Join Mst_AssetsType MAT On EAM.AssetID = MAT.AssetsID " & _
      " Left Join Mst_AssetStatus MAS On EAM.AssetStatusID = MAS.AssetsStatusID  "
tmp = tmp & Wc

Call TmpTable

Call dispdata

If MSHFlexGrid1.Rows > 1 Then
   MSHFlexGrid1.FixedRows = 1
End If

Label2.Refresh
Label2.Caption = TmpRs.RecordCount
CmdExcel.Enabled = True
MsgBox "Done!!!"


End Sub
Public Sub dispdata()
Dim x As Double
x = 0
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(x, 0) = I
       For c = 2 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c - 1)), "", TmpRs.Fields(c - 1))
       Next
       If LCase(MSHFlexGrid1.TextMatrix(x, 4)) = "y" Then
          MSHFlexGrid1.TextMatrix(x, 4) = "Active"
       Else
          MSHFlexGrid1.TextMatrix(x, 4) = "De-Active"
       End If
       If LCase(MSHFlexGrid1.TextMatrix(x, 3)) = "cmpm" Then
          Call AddROData(TmpRs!EmployeeID, x)
       End If
       MSHFlexGrid1.TextMatrix(x, 9) = Format(MSHFlexGrid1.TextMatrix(x, 9), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(x, 11) = Format(MSHFlexGrid1.TextMatrix(x, 11), "dd-MMM-yyyy")
       
       TmpRs.MoveNext
    Next
End If
End Sub
Sub AddROData(mEmpID_ As Double, x1 As Double)

tmp = " Select Distinct(ML.LocationName) " & _
      " From Mst_EMPCMPMapping ECM " & _
      " Inner Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
      " Where EmployeeID = " & mEmpID_ & ""
Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
    If MSHFlexGrid1.TextMatrix(x1, 1) = "" Then
       MSHFlexGrid1.TextMatrix(x1, 1) = IIf(IsNull(TmpRs1.Fields(0)), "", TmpRs1.Fields(0))
    Else
       MSHFlexGrid1.TextMatrix(x1, 1) = MSHFlexGrid1.TextMatrix(x1, 1) & "," & IIf(IsNull(TmpRs1.Fields(0)), "", TmpRs1.Fields(0))
    End If
    TmpRs1.MoveNext
Next
TmpRs1.Close
End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Label2.Caption = ""
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Label2.Caption = ""
Call Grid_Head
CmdExcel.Enabled = False


End Sub


Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
MSHFlexGrid1.Cols = 13
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "SrNo"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee"
MSHFlexGrid1.TextMatrix(0, 3) = "Designation"
MSHFlexGrid1.TextMatrix(0, 4) = "Status"
MSHFlexGrid1.TextMatrix(0, 5) = "Asset Type"
MSHFlexGrid1.TextMatrix(0, 6) = "Asset Qty"
MSHFlexGrid1.TextMatrix(0, 7) = "Description"
MSHFlexGrid1.TextMatrix(0, 8) = "Asset Status"
MSHFlexGrid1.TextMatrix(0, 9) = "Asset Verified On"
MSHFlexGrid1.TextMatrix(0, 10) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 11) = "AMC Exp. Date"
MSHFlexGrid1.TextMatrix(0, 12) = "AMC Amount"

MSHFlexGrid1.ColWidth(0) = 500
MSHFlexGrid1.ColWidth(1) = 1600
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 2200
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400

'MSHFlexGrid1.RowHeight(0) = 800

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeAsset Map.xls")

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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_EmployeeAsset Map.xls")
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




