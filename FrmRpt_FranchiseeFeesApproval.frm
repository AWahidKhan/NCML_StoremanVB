VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_FranchiseeFeesApproval 
   Caption         =   "Franchisee Fees Approval "
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   15000
      Begin VB.OptionButton Option1 
         Caption         =   "Claim Date"
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
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Expense Date"
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
         Left            =   1500
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         Width           =   1575
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
         ItemData        =   "FrmRpt_FranchiseeFeesApproval.frx":0000
         Left            =   13620
         List            =   "FrmRpt_FranchiseeFeesApproval.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   645
         Width           =   1305
      End
      Begin VB.ComboBox CmbLessor 
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
         Left            =   10110
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   660
         Width           =   3495
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
         Left            =   6525
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   3495
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
         Left            =   13695
         TabIndex        =   9
         Top             =   1155
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
         Left            =   11265
         TabIndex        =   7
         Top             =   1155
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
         Left            =   12480
         TabIndex        =   8
         Top             =   1155
         Width           =   1215
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
         Left            =   2940
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   660
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1530
         TabIndex        =   2
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   48300033
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   48300033
         CurrentDate     =   36494
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Claim Status"
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
         Left            =   13710
         TabIndex        =   20
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   315
         TabIndex        =   19
         Top             =   435
         Width           =   945
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
         Left            =   1785
         TabIndex        =   18
         Top             =   420
         Width           =   855
      End
      Begin VB.Label LblGodown 
         AutoSize        =   -1  'True
         Caption         =   "Franchisee Name"
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
         Left            =   11160
         TabIndex        =   16
         Top             =   360
         Width           =   1605
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
         Left            =   7860
         TabIndex        =   15
         Top             =   345
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   7200
         TabIndex        =   14
         Top             =   1200
         Width           =   2325
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "--"
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
         Left            =   9600
         TabIndex        =   13
         Top             =   1200
         Width           =   165
      End
      Begin VB.Label Label6 
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
         Left            =   4155
         TabIndex        =   12
         Top             =   375
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7620
      Left            =   120
      TabIndex        =   0
      Top             =   1665
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   150
         TabIndex        =   17
         Top             =   7215
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12091
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   3
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
   End
End
Attribute VB_Name = "FrmRpt_FranchiseeFeesApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mLessorID, mStatus As Variant

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 26 '25
    .Rows = 2
    .WordWrap = True
    .TextMatrix(0, 0) = "State Name"
    .TextMatrix(0, 1) = "Location Name"
    .TextMatrix(0, 2) = "Claim ID"
    .TextMatrix(0, 3) = "Claim Date"
    .TextMatrix(0, 4) = "Franchisee Fee From"
    .TextMatrix(0, 5) = "Franchisee Fee To"
    .TextMatrix(0, 6) = "Franchisee Name"
    .TextMatrix(0, 7) = "Remark:CMG"
    .TextMatrix(0, 8) = "Remark : Approval"
    .TextMatrix(0, 9) = "Approval Date"
    .TextMatrix(0, 10) = "Handling Officer"
    .TextMatrix(0, 11) = "Claim Status"
    .TextMatrix(0, 12) = "Updated By FAG"
    .TextMatrix(0, 13) = "FAG Posted Date"
    .TextMatrix(0, 14) = "CMP ID"
    .TextMatrix(0, 15) = "Godown No"
    .TextMatrix(0, 16) = "Address"
    .TextMatrix(0, 17) = "Franchisee for Godown"
    .TextMatrix(0, 18) = "Bank Name"
    .TextMatrix(0, 19) = "Cheque No"
    .TextMatrix(0, 20) = "Total franchisee fees to franchisee [Actual Fee]"
    .TextMatrix(0, 21) = "TDS"
    .TextMatrix(0, 22) = "Cheque/RTGS Amount"
    .TextMatrix(0, 23) = "Cheque/RTGS Date"
    .TextMatrix(0, 24) = "Area in (Sq. Ft.)"
    .TextMatrix(0, 25) = "Franchisee Bill No"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 2500
    .ColWidth(7) = 2500
    .ColWidth(8) = 2500
    .ColWidth(9) = 1500
    .ColWidth(10) = 2000
    .ColWidth(11) = 1000
    .ColWidth(12) = 1800
    .ColWidth(13) = 1500
    .ColWidth(14) = 1000
    .ColWidth(15) = 1800
    .ColWidth(16) = 2000
    .ColWidth(17) = 2000
    .ColWidth(18) = 2000
    .ColWidth(19) = 1500
    .ColWidth(20) = 2000
    .ColWidth(21) = 1500
    .ColWidth(22) = 1500
    .ColWidth(23) = 1500
    .ColWidth(24) = 1500
    .ColWidth(25) = 1800
    .RowHeight(0) = 700
End With
End Sub

Private Sub CmbLessor_GotFocus()
tmp = CmbLessor.Text
Call TableMst_Vendor(" Order by VendorName")

CmbLessor.Clear
If RsMst_Vendor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Vendor.RecordCount
    CmbLessor.AddItem RsMst_Vendor!VendorName
    RsMst_Vendor.MoveNext
Next
CmbLessor.Text = tmp
End Sub

Private Sub CmbLessor_LostFocus()
If CmbLessor.Text = "" Then
   Exit Sub
End If
RsMst_Vendor.MoveFirst
RsMst_Vendor.Find "VendorName = '" & CmbLessor.Text & "'"
If RsMst_Vendor.EOF Then
   MsgBox "Invalid selection "
   CmbLessor.SetFocus
   Exit Sub
End If
mLessorID = RsMst_Vendor!VendorID
End Sub
Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" where " & Gen_WcState & "")
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
   CmbLocation.Clear
   Exit Sub
End If
If CmbStateID.Text <> tmp Then
   CmbLocation.Clear
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
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text

Call TableMst_Location(" Where LocationID <> 0 And StateID= " & mStateID & "")

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
Private Sub CmbStatus_LostFocus()
mStatus = ""
If CmbStatus.Text = "" Then Exit Sub
If LCase(CmbStatus.Text) <> "pending" And LCase(CmbStatus.Text) <> "approved" And LCase(CmbStatus.Text) <> "frezz" And LCase(CmbStatus.Text) <> "rejected" Then
   MsgBox "Invalid selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
If LCase(CmbStatus.Text) = "pending" Then
   mStatus = "P"
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
ElseIf LCase(CmbStatus.Text) = "approved" Then
   mStatus = "A"
ElseIf LCase(CmbStatus.Text) = "frezz" Then
   mStatus = "F"
End If
End Sub

Private Sub CmdClear_Click()
Label5.Caption = ""
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbStateID.Text = ""
CmbLessor.Text = ""
CmbStatus.Text = ""
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_FranchiseeFeesApproval.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_FranchiseeFeesApproval.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
       If C = 25 Then
          oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
       Else
          oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
       End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
wc = ""
Label5.Caption = "--"
ProgressBar1.Value = 0
Call Grid_Head

MsgBox "Please wait till next Message."

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date must be less than To Date"
   TxtFromDate.SetFocus
   Exit Sub
End If

If Option1.Value = True Then
   wc = wc & " Where TFR.ClaimDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "'"
Else
   wc = wc & " Where TFR.ExpenseFrom Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "'"
End If

If CmbStateID.Text <> "" Then
   wc = wc & " And MS.StateID = " & mStateID
End If

If CmbLocation.Text <> "" Then
   wc = wc & " And ML.LocationID = " & mLocationID
End If

If CmbLessor.Text <> "" Then
   wc = wc & " And MLE.VendorID = " & mLessorID
End If


If CmbStatus.Text <> "" Then
   wc = wc & " And TFR.Flag = '" & mStatus & "'"
End If

''''tmp = "Select MS.StateName, ML.LocationName, TFR.ClaimID, TFR.ClaimDate, TFR.ExpenseFrom, TFR.ExpenseTo, MV.VendorName, TFR.Remarks, TFR.ApprovalRemarks, TFR.ApprovalDate,"
''''tmp = tmp & " ME.EmployeeName+ '~'+ ME.EmployeeNo 'Handling Officer', TFR.Flag, ME1.Employee_Name+ '~'+ TFR.UpdatedByFAG 'Updated By FAG', TFR.ApprovalDate, MC.CMPName, MG.GodownNo,"
''''tmp = tmp & " MG.Address , Sum(IsNull(TIMGSLD.FranchInvoiceAmount, 0)) 'Franchisee for Godown',TFR.BankName, TFR.ChequeNo,"
''''tmp = tmp & " TFR.ActualFee,  TFR.TDS, TFR.PaidAmount, TFR.PaidDate, MG.Area, TFR.Amount, TFR.VendorID, MG.CMPID, TIMGSLD.GodownID"
''''tmp = tmp & " From Txn_FranchiseeRent as TFR"
'''''tmp = tmp & " Inner Join Mst_Lessor as MLE on MLE.LessorID = TFR.VendorID"
''''tmp = tmp & " Inner Join Mst_Vendor as MV on MV.VendorID = TFR.VendorID"
''''tmp = tmp & " Inner Join Mst_Employee as ME on ME.EmployeeID = TFR.ForwardToID"
''''tmp = tmp & " Inner Join Mst_Location As ML on ML.LocationID = TFR.LocationID"
''''tmp = tmp & " Inner Join Mst_State As MS on MS.StateID = ML.StateID"
''''tmp = tmp & " Left Join Mst_Users as ME1 on ME1.Employee_ID = TFR.UpdatedByFAG"
''''tmp = tmp & " Inner Join Txn_InvMonGSLDetail as TIMGSLD on TIMGSLD.FranchClaimID = TFR.ClaimID"
''''tmp = tmp & " Inner Join Mst_Godown as MG on MG.GodownID = TIMGSLD.GodownID"
''''tmp = tmp & " Inner Join Mst_CMP as MC on MC.CMPID = MG.CMPID"
''''tmp = tmp & wc
''''tmp = tmp & " Group By"
''''tmp = tmp & " MS.StateName, ML.LocationName, TFR.ClaimID, TFR.ClaimDate, TFR.ExpenseFrom, TFR.ExpenseTo,"
''''tmp = tmp & " TFR.VendorID, MV.VendorName, TFR.Remarks, TFR.ApprovalRemarks, TFR.ApprovalDate,"
''''tmp = tmp & " ME.EmployeeName, ME.EmployeeNo, TFR.Flag, ME1.EmployeeName, TFR.UpdatedByFAG,"
''''tmp = tmp & " MG.CMPID, MC.CMPName, MG.GodownNo, MG.Address,TIMGSLD.GodownID, TFR.BAnkName, TFR.ChequeNo,"
''''tmp = tmp & " TFR.TDS , TFR.PaidAmount, TFR.PaidDate, MG.Area, TFR.ActualFee, TFR.Amount"
''''tmp = tmp & " Order by MS.StateName, ML.LocationName, TFR.ClaimID"

tmp = "Select MS.StateName, ML.LocationName, TFR.ClaimID, TFR.ClaimDate, TFR.ExpenseFrom, TFR.ExpenseTo, MLE.VendorName, TFR.Remarks, TFR.ApprovalRemarks, TFR.ApprovedDate,"
tmp = tmp & " ME.EmployeeName+ '~'+ ME.EmployeeNo 'Handling Officer', TFR.Flag, ME1.EmployeeName+ '~'+ TFR.UpdatedByFAG 'Updated By FAG', TFR.ApprovedDate, MC.CMPName, MG.GodownNo,"
tmp = tmp & " MG.Address , Sum(IsNull(TIMGSLD.FranchInvoiceAmount, 0)) 'Franchisee for Godown',TFR.BankName, TFR.ChequeNo,"
tmp = tmp & " TFR.ActualFee , TFR.TDS, TFR.PaidAmount, TFR.PaidDate, MG.Area,TFR.FranchInvoiceNo, TFR.Amount, TFR.VendorID, MG.CMPID, TIMGSLD.GodownID"
tmp = tmp & " From Txn_FranchiseeRent as TFR"
tmp = tmp & " Inner Join Mst_Vendor as MLE on MLE.VendorID = TFR.VendorID"
tmp = tmp & " Inner Join Mst_Employee as ME on ME.EmployeeID = TFR.ForwardToID"
tmp = tmp & " Inner Join Mst_Location As ML on ML.LocationID = TFR.LocationID"
tmp = tmp & " Inner Join Mst_State As MS on MS.StateID = ML.StateID"
tmp = tmp & " Left Join Mst_Users as ME1 on ME1.EmployeeNo = TFR.UpdatedByFAG"
tmp = tmp & " Inner Join Txn_InvMonGSLDetail as TIMGSLD on TIMGSLD.FranchClaimID = TFR.ClaimID"
tmp = tmp & " Inner Join Mst_Godown as MG on MG.GodownID = TIMGSLD.GodownID"
tmp = tmp & " Inner Join Mst_CMP as MC on MC.CMPID = MG.CMPID"
tmp = tmp & wc '---Where ML.LocationName ='Ganjbasoda'
tmp = tmp & " Group By"
tmp = tmp & " MS.StateName, ML.LocationName, TFR.ClaimID, TFR.ClaimDate, TFR.ExpenseFrom, TFR.ExpenseTo,"
tmp = tmp & " TFR.VendorID, MLE.VendorName, TFR.Remarks, TFR.ApprovalRemarks, TFR.ApprovedDate,"
tmp = tmp & " ME.EmployeeName, ME.EmployeeNo, TFR.Flag, ME1.EmployeeName, TFR.UpdatedByFAG,"
tmp = tmp & " MG.CMPID, MC.CMPName, MG.GodownNo, MG.Address,TIMGSLD.GodownID, TFR.BAnkName, TFR.ChequeNo,"
tmp = tmp & " TFR.TDS , TFR.PaidAmount, TFR.PaidDate, MG.Area, TFR.ActualFee, TFR.Amount,TFR.FranchInvoiceNo"
tmp = tmp & " Order by MS.StateName, ML.LocationName, TFR.ClaimID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount
   Label5.Caption = TmpRs.RecordCount
   For I = 1 To TmpRs.RecordCount
       ProgressBar1.Value = I
       For C = 0 To MSHFlexGrid1.Cols - 1
           If C = 3 Or C = 4 Or C = 5 Or C = 9 Or C = 13 Or C = 23 Then
              MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", Format(TmpRs.Fields(C), "dd-MMM-yyyy"))
           Else
              MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           End If
       Next
       If MSHFlexGrid1.TextMatrix(I, 11) = "P" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(I, 11) = "R" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Rejected"
       ElseIf MSHFlexGrid1.TextMatrix(I, 11) = "A" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Approved"
       ElseIf MSHFlexGrid1.TextMatrix(I, 11) = "S" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Settled"
       ElseIf MSHFlexGrid1.TextMatrix(I, 11) = "F" Then
          MSHFlexGrid1.TextMatrix(I, 11) = "Frezz"
       End If
       TmpRs.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Next
   MsgBox "Done!!"
Else
   MsgBox "No Records Found."
End If
ProgressBar1.Value = ProgressBar1.Max
CmdExcel.Enabled = True
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
TxtFromDate.Value = Date
TxtToDate.Value = Date
Grid_Head
CmdExcel.Enabled = False
End Sub

