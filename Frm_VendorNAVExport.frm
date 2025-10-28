VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_VendorNAVExport 
   Caption         =   "Vendor NAV Export"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   135
      ScaleHeight     =   795
      ScaleWidth      =   14955
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbApprovedBy 
         Height          =   315
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox CmbBatch 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Old"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Get Data"
         Height          =   435
         Left            =   10620
         TabIndex        =   2
         Top             =   195
         Width           =   1335
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         Height          =   435
         Left            =   11940
         TabIndex        =   1
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Approved By"
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   270
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8655
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15266
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Frm_VendorNAVExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpFlag, MaxID As Variant
Dim mForwardToID As Double
Private Sub CmbApprovedBy_GotFocus()

tmp1 = CmbApprovedBy.Text
'Call TableMst_Employee("Where Flag = 'Y' And DesignationID=")
tmp = "Select EmployeeID,EmployeeNo,EmployeeName From Mst_Employee ME"
tmp = tmp & " Inner join Mst_Users MU On MU.EmployeeNo=ME.EmployeeNo Where MU.UserGroupID=7 And ME.Flag = 'Y'"

Call TmpTable
CmbApprovedBy.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Approved By Found!!!"
   CmbApprovedBy.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbApprovedBy.AddItem TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
    TmpRs.MoveNext
Next

CmbApprovedBy.Text = tmp1

End Sub

Private Sub CmbApprovedBy_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbApprovedBy.Text = "" Then
   mForwardToID = 0
   Exit Sub
End If

strSplitString = Split(CmbApprovedBy, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbApprovedBy.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


TmpRs.MoveFirst
TmpRs.Find "EmployeeNo = '" & EmpNo_ & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Selection "
   CmbApprovedBy.SetFocus
   Exit Sub
End If
mForwardToID = TmpRs!EmployeeID

End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant
Path = App.Path
If Option1.Value = True Then
   strFilName = "\Navision\VendorNavUpload_SMClaim" & MaxID & ".csv"
Else
   strFilName = "\Navision\VendorNavUpload_SMClaim" & CmbBatch.Text & ".csv"
End If
strFilLocation = Path + strFilName
If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
Open strFilLocation For Output As #1

For I = 1 To MSHFlexGrid1.Rows - 1
    tmp = ""
    For C = 0 To MSHFlexGrid1.Cols - 1
        If tmp = "" Then
           tmp = tmp & MSHFlexGrid1.TextMatrix(I, C)
        Else
           tmp = tmp & "," & MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    Print #1, tmp
Next
If Option1.Value = True Then
'   tmp = "Update Txn_VendorClaim Set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A'"
'   Call TmpTable

   If CmbApprovedBy.Text = "" Then
      tmp = "Update Txn_VendorClaim set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A'"
   Else
      tmp = "Update Txn_VendorClaim set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A' And ForwardToID = " & mForwardToID & ""
   End If
   Call TmpTable

End If

MsgBox "Done !!!"
Close #1
End Sub

Private Sub CmdGo_Click()

ExpFlag = ""
Call Grid_Head
If Option1.Value = True Then
        
        tmp = "Select VC.ApprovalDate,'Invoice' As Inv,'G/L Account' As GLAccount,ET.NAVCode 'ExpenseTypeNAVCode',"
        tmp = tmp & " CD.Amount 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode',  "
        tmp = tmp & " 'Being expenses incurred against Storeman Batch ID " & MaxID & "' As 'ApprovalRemarks', "
        tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim , "
        tmp = tmp & " '' 'EmployeeNAVCode', VC.BatchID,CD.InvoiceDate,CD.ClaimID  From Txn_VendorClaimDetail CD "
        tmp = tmp & " Inner Join Txn_VendorClaim VC on CD.ClaimID = VC.ClaimID  "
        tmp = tmp & " Inner Join Mst_Vendor E on VC.VendorID = E.VendorID "
        tmp = tmp & " Inner Join  Mst_Location L on CD.LocationID = L.LocationID "
        tmp = tmp & " Inner Join Mst_ExpenseType ET on CD.ExpenseTypeId = ET.ExpenseTypeID "
        tmp = tmp & " left  Join Mst_ServiceTax ST on ET.ServiceTaxID = ST.ServiceTaxID "
        tmp = tmp & " Where VC.BatchID =  0 and  VC.Flag = 'A'"
        If CmbApprovedBy.Text <> "" Then
           tmp = tmp & " And VC.ForwardToID=" & mForwardToID & ""
        End If
'        tmp = tmp & " Group by VC.ApprovalDate,ET.NAVCode,ST.NAVCode ,L.NAVCode, VC.BatchID,CD.InvoiceDate "
        tmp = tmp & " Order By VC.ApprovalDate"

        
        'tmp = tmp & " Group By VC.ApprovalDate,E.NAVCode,L.NAVCode , ET.NAVCode  ,VC.ApprovalRemarks,ST.NAVCode,VC.BatchID"
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
        
        tmp = "Select VC.ApprovalDate,'Invoice' As Inv,'G/L Account' As GLAccount,ET.NAVCode 'ExpenseTypeNAVCode',"
        tmp = tmp & " CD.Amount 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode', "
        tmp = tmp & " 'Being expenses incurred against Storeman Batch ID " & CmbBatch.Text & "' As 'ApprovalRemarks', "
        tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim , "
        tmp = tmp & " '' 'VendorNAVCode', VC.BatchID,CD.InvoiceDate,CD.ClaimID  From Txn_VendorClaimDetail CD "
        tmp = tmp & " Inner Join Txn_VendorClaim VC on CD.ClaimID = VC.ClaimID  "
        tmp = tmp & " Inner Join Mst_Vendor E on VC.VendorID = E.VendorID "
        tmp = tmp & " Inner Join  Mst_Location L on CD.LocationID = L.LocationID "
        tmp = tmp & " Inner Join Mst_ExpenseType ET on CD.ExpenseTypeId = ET.ExpenseTypeID "
        tmp = tmp & " left  Join Mst_ServiceTax ST on ET.ServiceTaxID = ST.ServiceTaxID "
        tmp = tmp & " Where VC.BatchID =  " & Val(CmbBatch.Text) & " and  VC.Flag = 'A'"
'        tmp = tmp & " Group by VC.ApprovalDate,ET.NAVCode,ST.NAVCode ,L.NAVCode, VC.BatchID,CD.InvoiceDate "
        tmp = tmp & " Order By VC.ApprovalDate "
        'tmp = tmp & " Group By VC.ApprovalDate,E.NAVCode,L.NAVCode , ET.NAVCode  ,VC.ApprovalRemarks,ST.NAVCode,VC.BatchID"
End If
Call TmpTable


If Option1.Value = True Then
    tmp = "Select VC.ApprovalDate,'Invoice' As Inv,'Vendor' As GLAccount,E.NAVCode 'ExpenseTypeNAVCode',"
    tmp = tmp & " (VC.Amount) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode', "
    tmp = tmp & " VC.ApprovalRemarks + ' - Claim ID:' + convert(varchar(10),VC.ClaimID) As 'ApprovalRemarks', "
    tmp = tmp & " '201001' 'LocationNAVCode','CMP' As DeptDim , "
    tmp = tmp & " E.NAVCode 'VendorNAVCode', VC.BatchID ,'' As InvoiceDate,VC.InvoiceNo,VC.ClaimID From  Txn_VendorClaim VC "
    tmp = tmp & " Inner Join Mst_Vendor E on VC.VendorID = E.VendorID "
    tmp = tmp & " Inner Join Mst_TDS T on E.TDSID = T.TDSID  "
    tmp = tmp & " Where VC.BatchID =  0 and  VC.Flag = 'A'"
    If CmbApprovedBy.Text <> "" Then
       tmp = tmp & " And VC.ForwardToID=" & mForwardToID & ""
    End If
    tmp = tmp & " Order By VC.ApprovalDate"
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
   tmp = "Select VC.ApprovalDate,'Invoice' As Inv,'Vendor' As GLAccount,E.NAVCode 'ExpenseTypeNAVCode',"
   tmp = tmp & " (VC.Amount) 'Amount', T.NavCode As TDS,'' As 'ServiceTaxNAVCode', "
   tmp = tmp & " VC.ApprovalRemarks + ' - Claim ID:' + convert(varchar(10),VC.ClaimID) As 'ApprovalRemarks', "
   tmp = tmp & " '201001'As 'LocationNAVCode','CMP' As DeptDim , "
   tmp = tmp & " '' As 'VendorNAVCode', VC.BatchID ,'' As InvoiceDate,VC.InvoiceNo,VC.ClaimID From  Txn_VendorClaim VC "
   tmp = tmp & " Inner Join Mst_Vendor E on VC.VendorID = E.VendorID "
   tmp = tmp & " Inner Join Mst_TDS T on E.TDSID = T.TDSID  "
   tmp = tmp & " Where VC.BatchID =  " & Val(CmbBatch.Text) & " and  VC.Flag = 'A'"
   tmp = tmp & " Order By VC.ApprovalDate"
End If
Call Tmp1Table
Dim mDat_ As Variant
If TmpRs.RecordCount > 0 And TmpRs1.RecordCount > 0 Then
    Dim x As Variant
    x = 0
    ExpFlag = "Y"
    mDat_ = IIf(IsNull(TmpRs.Fields(0)), "", TmpRs.Fields(0))
    For x = 1 To TmpRs.RecordCount
        If TmpRs!ExpenseTypeNAVCode = "" Or IsNull(TmpRs!ExpenseTypeNAVCode) Or TmpRs!LocationNAVCode = "" Or IsNull(TmpRs!LocationNAVCode) Then
           ExpFlag = "N"
        End If
        I = MSHFlexGrid1.Rows - 1
        For C = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        Next
        MSHFlexGrid1.TextMatrix(I, 0) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(I, 0))
        MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
        MSHFlexGrid1.TextMatrix(I, 12) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(I, 12))
        
        If CmbBatch.Text = "" Then
           MSHFlexGrid1.TextMatrix(I, 11) = "SMV" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
        Else
           MSHFlexGrid1.TextMatrix(I, 11) = "SMV" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
        End If
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        mDat_ = TmpRs!ClaimID
        
        TmpRs.MoveNext
        If TmpRs.EOF Then
           CrEntry (mDat_)
           Exit For
        End If
        If mDat_ <> IIf(IsNull(TmpRs!ClaimID), "", TmpRs!ClaimID) Then
           Call CrEntry(mDat_)
           mDat_ = IIf(IsNull(TmpRs!ClaimID), "", TmpRs!ClaimID)
        End If
        
    Next

Else
   ExpFlag = "N"
End If

If ExpFlag = "N" Then
   CmdExport.Enabled = False
Else
   CmdExport.Enabled = True
End If

MSHFlexGrid1.Col = 11
MSHFlexGrid1.Sort = flexSortNumericDescending
End Sub

Sub CrEntry(mPDat_ As Variant)

TmpRs1.MoveFirst
TmpRs1.Find "ClaimID = " & mPDat_ 'TmpRs!ClaimID
'For x = 1 To TmpRs1.RecordCount
If TmpRs1.EOF = False Then
    If TmpRs1!ExpenseTypeNAVCode = "" Or IsNull(TmpRs1!ExpenseTypeNAVCode) Or TmpRs1!LocationNAVCode = "" Or IsNull(TmpRs1!LocationNAVCode) Then
       ExpFlag = "N"
    End If
    I = MSHFlexGrid1.Rows - 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(I, 0) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(I, 0))
    MSHFlexGrid1.TextMatrix(I, 4) = Val(MSHFlexGrid1.TextMatrix(I, 4)) * -1
    MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
    
'    If CmbBatch.Text = "" Then
'       MSHFlexGrid1.TextMatrix(I, 11) = "SMV" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'    Else
'       MSHFlexGrid1.TextMatrix(I, 11) = "SMV" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'    End If

'        If CmbBatch.Text = "" Then
           MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!InvoiceNo), "", "Invoice " & TmpRs1!InvoiceNo) '"SMV" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'        Else
'           MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo) 'TmpRs!InvoiceNo  '"SMV" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
'        End If
    
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'    TmpRs1.MoveNext
'    If TmpRs1.EOF Then Exit For
'    If mPDat_ <> IIf(IsNull(TmpRs1.Fields(0)), "", TmpRs1.Fields(0)) Then Exit For
'Next
End If

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Picture1.Enabled = False
   MSHFlexGrid1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Picture1.Enabled = True
MSHFlexGrid1.Enabled = True
Call Grid_Head
tmp = "Select Distinct BatchID from Txn_VendorClaim Where BatchID > 0  order by BatchID"
Call Tmp4Table
CmbBatch.Clear
If TmpRs4.RecordCount > 0 Then
   For I = 1 To TmpRs4.RecordCount
       CmbBatch.AddItem TmpRs4.Fields(0)
       MaxID = TmpRs4.Fields(0) + 1
       TmpRs4.MoveNext
   Next
Else
   MaxID = 1
End If
End Sub
Private Sub Option2_Click()
CmdExport.Enabled = False
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbApprovedBy.Text = ""
   CmbApprovedBy.Enabled = False
Else
   CmbBatch.Enabled = False
   CmbApprovedBy.Text = ""
   CmbApprovedBy.Enabled = True
End If
End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Public Sub Grid_Head()

CmdExport.Enabled = False

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13

MSHFlexGrid1.TextMatrix(0, 0) = "Posting Date"
MSHFlexGrid1.TextMatrix(0, 1) = "Document Type"
MSHFlexGrid1.TextMatrix(0, 2) = "Account Type"
MSHFlexGrid1.TextMatrix(0, 3) = "Account No."
MSHFlexGrid1.TextMatrix(0, 4) = "Amount"
MSHFlexGrid1.TextMatrix(0, 5) = "TDS Nature of Deduction"
MSHFlexGrid1.TextMatrix(0, 6) = "Service Tax Group Code"
MSHFlexGrid1.TextMatrix(0, 7) = "Narration"
MSHFlexGrid1.TextMatrix(0, 8) = "RegionDim"
MSHFlexGrid1.TextMatrix(0, 9) = "DeptDim"
MSHFlexGrid1.TextMatrix(0, 10) = "VendorDim"
MSHFlexGrid1.TextMatrix(0, 11) = "External Document No."
MSHFlexGrid1.TextMatrix(0, 12) = "Invoice Date"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 800
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1000

End Sub




