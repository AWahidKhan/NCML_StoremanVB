VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_VendorClaimDetails 
   Caption         =   "Vendor Claim Details"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   9495
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   15000
      Begin VB.ComboBox CmbExpenseTypeID 
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
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1050
         Width           =   6300
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
         Left            =   13530
         TabIndex        =   9
         Top             =   1035
         Width           =   1125
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
         Left            =   2925
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   390
         Width           =   2445
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
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
         Left            =   11280
         TabIndex        =   7
         Top             =   1035
         Width           =   1125
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
         Left            =   12405
         TabIndex        =   8
         Top             =   1035
         Width           =   1125
      End
      Begin VB.ComboBox CmbSForwardTo 
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
         Left            =   9090
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   390
         Width           =   3675
      End
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmRpt_VendorClaimDetails.frx":0000
         Left            =   12780
         List            =   "FrmRpt_VendorClaimDetails.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   390
         Width           =   1875
      End
      Begin VB.ComboBox CmbSVendor 
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
         Left            =   5445
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   3585
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7815
         Left            =   60
         TabIndex        =   11
         Top             =   1455
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   13785
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   360
         Left            =   1530
         TabIndex        =   1
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
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
         Format          =   22544385
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   360
         Left            =   75
         TabIndex        =   0
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
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
         Format          =   22544385
         CurrentDate     =   36494
      End
      Begin VB.Label Label3 
         Caption         =   "Expense Type"
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
         Left            =   2460
         TabIndex        =   20
         Top             =   780
         Width           =   1530
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   3705
         TabIndex        =   19
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label31 
         Caption         =   "Forwarded To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10290
         TabIndex        =   18
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label16 
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
         Height          =   300
         Left            =   13335
         TabIndex        =   17
         Top             =   105
         Width           =   720
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
         Left            =   6600
         TabIndex        =   16
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label22 
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
         Left            =   9120
         TabIndex        =   15
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label Label45 
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
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Vendor Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6600
         TabIndex        =   13
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label44 
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
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmRpt_VendorClaimDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSStatus, mLocationID As Variant

Private Sub CmbExpenseTypeID_GotFocus()
Dim wc As String
tmp = CmbExpenseTypeID.Text
Call TableMst_ExpenseType(" Where (AllocateTo = 'B' Or AllocateTo = 'V') And Flag = 'A' And  OtherDeductionFlag = 'N' ")
CmbExpenseTypeID.Clear
If RsMst_ExpenseType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExpenseType.RecordCount
    CmbExpenseTypeID.AddItem RsMst_ExpenseType!ExpenseType
    RsMst_ExpenseType.MoveNext
Next
CmbExpenseTypeID.Text = tmp
End Sub
'----
Private Sub CmbExpenseTypeID_LostFocus()
If CmbExpenseTypeID.Text = "" Then
   mExpenseTypeID = Null
   Exit Sub
End If
RsMst_ExpenseType.MoveFirst
RsMst_ExpenseType.Find "ExpenseType = '" & CmbExpenseTypeID.Text & "'"
If RsMst_ExpenseType.EOF Then
   MsgBox "Invalid selection "
   CmbExpenseTypeID.SetFocus
   Exit Sub
End If
mExpenseTypeID = RsMst_ExpenseType!ExpenseTypeID
End Sub

Private Sub CmbSForwardTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSForwardTo.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbSForwardTo, "~")
If UBound(strSplitString) = 1 Then
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
Else
   MsgBox "Invalid Selection!!!"
   CmbSForwardTo.SetFocus
   Exit Sub
End If
RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSForwardTo.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()

SFrom.Value = CDate(From_GenDate) - 30
STo.Value = Date
CmbLocation.Text = ""
CmbSVendor.Text = ""
CmbSForwardTo.Text = ""
CmbSStatus.Text = ""
Label22.Caption = ""
CmbExpenseTypeID.Text = ""
Call Grid_Head

End Sub

Private Sub CmdGo_Click()

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

Call Grid_Head

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

Label22.Caption = ""

tmp = " Select VC.ClaimID,VC.ClaimDate,VC.ExpenseFrom,VC.ExpenseTo,VC.Amount as 'TotalClaim',ML.LocationName,"
tmp = tmp & " MV.VendorName,VC.Flag,VC.Remarks,VC.ApprovalRemarks,VC.ApprovalDate,EMP.EmployeeName +'~'+ EMP.EmployeeNo as ForwardedTo," 'VC.G_UID,
tmp = tmp & " ME.ExpenseType,TVC.Particulars,TVC.Amount,TVC.ServiceTax,'',TVC.InvoiceDate,VC.BatchID "
tmp = tmp & " From Txn_VendorClaim VC"
tmp = tmp & " Inner Join Mst_Vendor MV On VC.VendorID = MV.VendorID"
tmp = tmp & " Inner Join Txn_VendorClaimDetail TVC on TVC.ClaimID=VC.ClaimID"
tmp = tmp & " Inner Join Mst_Location ML on TVC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_ExpenseType ME on ME.ExpensetypeID=TVC.ExpenseTypeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP On VC.ForwardToID = EMP.EmployeeID"


If Gen_DBType = "MDB" Then
   tmp = tmp & " Where VC.ExpenseFrom Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where VC.ExpenseFrom Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
End If


If CmbLocation.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

If CmbSVendor.Text <> "" Then
   tmp = tmp & " AND MV.VendorName = '" & CmbSVendor.Text & "' "
End If

If CmbSStatus.Text <> "" Then
   tmp = tmp & " AND VC.Flag = '" & mSStatus & "' "
End If

If CmbSForwardTo.Text <> "" Then
   strSplitString = Split(CmbSForwardTo, "~")
   EmpName_ = Trim(strSplitString(0))
   EmpNo_ = Trim(strSplitString(1))
 
   tmp = tmp & " AND EMP.EmployeeNo = '" & EmpNo_ & "' "
'Else
'For Checking
'   If LCase(Gen_UserRole) <> "power" Then
'      If Gen_DesignationID <> 1 Then
'         If Gen_WcEmployeeID <> "" Then
'            tmp = tmp & " AND EMP." & Gen_WcEmployeeID
'         End If
'      End If
'   End If
End If

If CmbExpenseTypeID.Text <> "" Then
   tmp = tmp & " AND ME.ExpenseType = '" & CmbExpenseTypeID.Text & "' "
End If


Call TmpTable


If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   
   For mRow = 1 To TmpRs.RecordCount
   
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(mRow, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(mRow, 1) = Format(MSHFlexGrid1.TextMatrix(mRow, 1), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(mRow, 2) = Format(MSHFlexGrid1.TextMatrix(mRow, 2), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(mRow, 3) = Format(MSHFlexGrid1.TextMatrix(mRow, 3), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(mRow, 10) = Format(MSHFlexGrid1.TextMatrix(mRow, 10), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(mRow, 17) = Format(MSHFlexGrid1.TextMatrix(mRow, 17), "dd-MMM-yyyy")
       
       MSHFlexGrid1.TextMatrix(mRow, 16) = Val(MSHFlexGrid1.TextMatrix(mRow, 14)) + Val(MSHFlexGrid1.TextMatrix(mRow, 15))
       
       If LCase(MSHFlexGrid1.TextMatrix(mRow, 7)) = "a" Then
          MSHFlexGrid1.TextMatrix(mRow, 7) = "Approved"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(mRow, 7)) = "p" Then
          MSHFlexGrid1.TextMatrix(mRow, 7) = "Pending"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(mRow, 7)) = "r" Then
          MSHFlexGrid1.TextMatrix(mRow, 7) = "Rejected"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(mRow, 7)) = "f" Then
          MSHFlexGrid1.TextMatrix(mRow, 7) = "Frozen"
       End If
       TmpRs.MoveNext
   Next
  
   MsgBox "Done !"
Else
   MsgBox " No Record Found !!"
End If

Label22.Refresh
Label22.Caption = TmpRs.RecordCount


End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 19
     
     .FixedRows = 1
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(0, 0) = "Claim ID"
     .TextMatrix(0, 1) = "Claim Date"
     .TextMatrix(0, 2) = "Expense From"
     .TextMatrix(0, 3) = "Expense To"
     .TextMatrix(0, 4) = "Total Claim"
     .TextMatrix(0, 5) = "Location Name"
     .TextMatrix(0, 6) = "Vendor Name"
     .TextMatrix(0, 7) = "Flag"
     .TextMatrix(0, 8) = "Remarks"
     .TextMatrix(0, 9) = "Approval Remarks"
     .TextMatrix(0, 10) = "Approval Date"
     .TextMatrix(0, 11) = "Forwarded To"
     .TextMatrix(0, 12) = "Expense Type"
     .TextMatrix(0, 13) = "Particulars"
     .TextMatrix(0, 14) = "Amount"
     .TextMatrix(0, 15) = "STax Amount"
     .TextMatrix(0, 16) = "Gross Amount"
     .TextMatrix(0, 17) = "Invoice Date"
     .TextMatrix(0, 18) = "Batch No"
        
     .ColWidth(0) = 800
     .ColWidth(1) = 1000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1200
     .ColWidth(5) = 1500
     .ColWidth(6) = 2000
     .ColWidth(7) = 1000
     .ColWidth(8) = 2000
     .ColWidth(9) = 1800
     .ColWidth(10) = 1000
     .ColWidth(11) = 2000
     .ColWidth(12) = 2000
     .ColWidth(13) = 2000
     .ColWidth(14) = 1200
     .ColWidth(15) = 1000
     .ColWidth(16) = 1000
     .ColWidth(17) = 1000
     .ColWidth(18) = 1000
     .RowHeight(0) = 500
     .WordWrap = True
End With
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_VendorClaimDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_VendorClaimDetails.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame0.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame0.Enabled = True
Call CmdClear_Click

End Sub


Private Sub CmbSVendor_GotFocus()
tmp = CmbSVendor.Text
CmbSVendor.Clear
Call GetSearchVendorDataonGotFocus
CmbSVendor.Text = tmp
End Sub
Private Sub GetSearchVendorDataonGotFocus()

Call TableMst_Vendor '("Where Flag = 'Y'")
CmbSVendor.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Vendor Found!!!"
   CmbSVendor.SetFocus
   'RsMst_Vendor.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
    CmbSVendor.AddItem RsMst_Vendor!VendorName
    RsMst_Vendor.MoveNext
Next

End Sub

Private Sub CmbSVendor_LostFocus()

Call GetSearchVendorDataOnLostFocus

End Sub
Private Sub GetSearchVendorDataOnLostFocus()

If CmbSVendor.Text = "" Then
   mSVendorID = Null
   Exit Sub
End If

RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbSVendor.Text & "'"

If RsMst_Vendor.EOF Then
   MsgBox "Invalid Selection "
   CmbSVendor.SetFocus
   Exit Sub
End If
mSVendorID = RsMst_Vendor!VendorID

End Sub


Private Sub CmbSForwardTo_GotFocus()
tmp = CmbSForwardTo.Text

If Gen_WcEmployeeID <> "" Then
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" Then
      Call TableMst_Employee("Where " & Gen_WcEmployeeID)
   Else
      Call TableMst_Employee
   End If
End If

CmbSForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSForwardTo.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSForwardTo.Text = tmp

End Sub


Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
ElseIf LCase(CmbSStatus.Text) = "rejected" Then
   mSStatus = "R"
ElseIf LCase(CmbSStatus.Text) = "frozen" Then
   mSStatus = "F"
Else
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   mLocationID = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If


mLocationID = RsMst_Location!LocationID

End Sub

