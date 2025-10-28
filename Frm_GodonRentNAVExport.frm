VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_GodonRentNAVExport 
   Caption         =   "Godown Rent NAV Export"
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
      Height          =   1215
      Left            =   135
      ScaleHeight     =   1155
      ScaleWidth      =   14955
      TabIndex        =   7
      Top             =   120
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
         Height          =   350
         Left            =   13320
         TabIndex        =   12
         Top             =   720
         Width           =   1140
      End
      Begin VB.ComboBox CmbJournalCode 
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
         Left            =   10560
         TabIndex        =   4
         Top             =   240
         Width           =   2160
      End
      Begin VB.ComboBox CmbApprovedBy 
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
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13860
         TabIndex        =   6
         Top             =   195
         Width           =   975
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Get Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12780
         TabIndex        =   5
         Top             =   195
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
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
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Old"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox CmbBatch 
         Enabled         =   0   'False
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
         ItemData        =   "Frm_GodonRentNAVExport.frx":0000
         Left            =   2280
         List            =   "Frm_GodonRentNAVExport.frx":0002
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Journal Code"
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
         Left            =   9240
         TabIndex        =   11
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Approved By"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   300
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8055
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14208
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "Frm_GodonRentNAVExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpFlag, MaxID As Variant
Dim mForwardToID As Double
Dim mJournalBatchName, mJournalType As Variant
Dim mDocumentNo As Variant

Dim mWarehouseRentExpenseID As Double

Private Sub CmbBatch_GotFocus()
tmp = "Select Distinct BatchID from Txn_GodownRent Where BatchID > 0  order by BatchID"
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

Private Sub CmbJournalCode_GotFocus()
Dim mR As Variant
tmp1 = CmbJournalCode
CmbJournalCode.Clear
tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '" & Gen_UserLoginID & "' Order by JCM.JournalTypeID"
Call TmpTableHRPro
If TmpRsHRPro.RecordCount > 0 Then
For mR = 1 To TmpRsHRPro.RecordCount
   CmbJournalCode.AddItem IIf(IsNull(TmpRsHRPro!JournalCode), "", TmpRsHRPro!JournalCode)
   TmpRsHRPro.MoveNext
Next
Else
   MsgBox "No Record found !!!! "
   Exit Sub
End If
CmbJournalCode = tmp1
End Sub

Private Sub CmbJournalCode_LostFocus()
Dim mSplit As Variant
If CmbJournalCode.Text = "" Then
   mJournalBatchName = ""
   mJournalType = ""
   Exit Sub
End If
mSplit = Split(CmbJournalCode.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection!!! "
   CmbJournalCode.SetFocus
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "JournalCode = '" & CmbJournalCode.Text & "'"
If TmpRsHRPro.EOF Then
   MsgBox "Invalid selection "
   CmbJournalCode.SetFocus
   Exit Sub
End If
mJournalType = UCase(mSplit(0))
mJournalBatchName = UCase(mSplit(1))
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExportGodownRent.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExportGodownRent.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant
Dim mUserBatchID, mInvoiceType As Variant

Path = App.Path
''-- On 08 Aug 2016
'mJournalType = "GENERAL"
'mUserBatchID = GetGUID
'mJournalBatchName = GetJournalDetailsFromHRPro(Gen_UserLoginID, mJournalType)
'If Trim(mJournalBatchName) = "" Then
'   MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'   Exit Sub
'End If

''---- Ullhas Patil 12 Oct 2017
''-- On 05 Aug 2016
If Option1.Value = True Then
    If CmbJournalCode.Text = "" Then
       MsgBox "Journal Code not selected"
       CmbJournalCode.SetFocus
       Exit Sub
    End If
    mUserBatchID = GetGUID
'    If Trim(mJournalType) = "" Or Trim(mJournalBatchName) = "" Then
'       MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'       Exit Sub
'    End If
End If

mInvoiceType = "GODOWNRENT"

If Option1.Value = True Then
   tmp = "Select Max(BatchID) 'BatchID' from Txn_GodownRent Where BatchID > 0 "
   Call Tmp4Table
   MaxID = IIf(IsNull(TmpRs4!BatchID), 1, TmpRs4!BatchID + 1)
   strFilName = "\Navision\GodownRentNavUpload_SMClaim" & MaxID & ".csv"
Else
   strFilName = "\Navision\GodownRentNavUpload_SMClaim" & CmbBatch.Text & ".csv"
End If

''-- On 08 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
    ''---- Ullhas Patil 12 Oct 2017
    Call TableTxn_SecurityBillForNavision(" Where G_UID is null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
End If

With MSHFlexGrid1
For I = 1 To MSHFlexGrid1.Rows - 1
    tmp = ""
    If Option1.Value = True Then
       If .TextMatrix(I, 0) <> "" Then
''---- Ullhas Patil 12 Oct 2017
        RsTxn_SecurityBillForNavision.AddNew
        RsTxn_SecurityBillForNavision!PostingDate = CDate(.TextMatrix(I, 0))
        RsTxn_SecurityBillForNavision!DocumentType = Trim(.TextMatrix(I, 1))
        RsTxn_SecurityBillForNavision!AccountType = Trim(.TextMatrix(I, 2))
        RsTxn_SecurityBillForNavision!AccountNo = Trim(.TextMatrix(I, 3))
        RsTxn_SecurityBillForNavision!Amount = Trim(.TextMatrix(I, 4))
        RsTxn_SecurityBillForNavision!TDSNatureofDeduction = Trim(.TextMatrix(I, 5))
        RsTxn_SecurityBillForNavision!ServiceTaxGroupCode = Trim(.TextMatrix(I, 6))
        RsTxn_SecurityBillForNavision!Narration = Trim(.TextMatrix(I, 7))
        RsTxn_SecurityBillForNavision!RegionDim = Trim(.TextMatrix(I, 8))
        RsTxn_SecurityBillForNavision!DeptDim = Trim(.TextMatrix(I, 9))
        RsTxn_SecurityBillForNavision!VendorDim_EmpDim = Trim(.TextMatrix(I, 10))
        RsTxn_SecurityBillForNavision!ExternalDocumentNo = Trim(.TextMatrix(I, 11))
        RsTxn_SecurityBillForNavision!InvoiceDate = IIf(Trim(.TextMatrix(I, 12)) = "", Null, Trim(.TextMatrix(I, 12)))
        RsTxn_SecurityBillForNavision!InvoiceType = mInvoiceType
        'RsTxn_SecurityBillForNavision!PINo = Left(Trim(.TextMatrix(I, 11)), InStr(Trim(.TextMatrix(I, 11)), "-") - 1)
        RsTxn_SecurityBillForNavision!CreatedBy = Gen_UserLoginID
        RsTxn_SecurityBillForNavision!G_UID = GetGUID
        RsTxn_SecurityBillForNavision!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
        RsTxn_SecurityBillForNavision!JournalBatchType = Trim(mJournalType)
        RsTxn_SecurityBillForNavision!JournalBatchName = Trim(mJournalBatchName)
        RsTxn_SecurityBillForNavision!DocumentNo = Trim(.TextMatrix(I, 11))
        RsTxn_SecurityBillForNavision!UserBatchID = mUserBatchID
        
        RsTxn_SecurityBillForNavision!ProjectDim = Trim(.TextMatrix(I, 13))
        RsTxn_SecurityBillForNavision!GSTGroupCode = Trim(.TextMatrix(I, 14))
        RsTxn_SecurityBillForNavision!HSNNo = Trim(.TextMatrix(I, 15))
        RsTxn_SecurityBillForNavision!AccType = Trim(.TextMatrix(I, 16))
        
        'RsTxn_SecurityBillForNavision!DataTransfer = "X"
        
        RsTxn_SecurityBillForNavision.Update

'        For c = 0 To MSHFlexGrid1.Cols - 1
'            If tmp = "" Then
'               tmp = tmp & MSHFlexGrid1.TextMatrix(i, c)
'            Else
'               tmp = tmp & "," & MSHFlexGrid1.TextMatrix(i, c)
'            End If
'        Next
'        Print #1, tmp

       End If
    Else
    For C = 0 To MSHFlexGrid1.Cols - 1
        If tmp = "" Then
           tmp = tmp & MSHFlexGrid1.TextMatrix(I, C)
        Else
           tmp = tmp & "," & MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    Print #1, tmp
    End If
Next
End With

If Option1.Value = True Then
'   tmp = "Update Txn_GodownRent Set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A'"
'   Call TmpTable
   If CmbApprovedBy.Text = "" Then
      tmp = "Update Txn_GodownRent set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A'"
   Else
      tmp = "Update Txn_GodownRent set BatchID = " & MaxID & " Where BatchID = 0 and  Flag = 'A' And ForwardToID = " & mForwardToID & ""
   End If
   Call TmpTable
End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Close #1
Call Grid_Head
End Sub

Private Sub CmdGo_Click()
ExpFlag = ""
Call Grid_Head

If Option1.Value = True Then
   tmp = "Select Max(BatchID) 'BatchID' from Txn_GodownRent Where BatchID > 0 "
   Call Tmp4Table
   MaxID = IIf(IsNull(TmpRs4!BatchID), 1, TmpRs4!BatchID + 1)
End If

If Option1.Value = True Then

''   tmp = "Select CD.ApprovalDate,'Invoice' As Inv,'G/L Account' As GLAccount,ET.NAVCode 'ExpenseTypeNAVCode',"
''   tmp = tmp & " Sum (GD.Amount + GD.Arrears) 'Amount', '' As TDS,'' 'ServiceTaxNAVCode'," 'ST.NAVCode 'ServiceTaxNAVCode',"
''   tmp = tmp & " 'Being warehouse rent payable against Storeman Batch ID " & MaxID & "' As 'ApprovalRemarks',"
''   tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
''   tmp = tmp & " '' 'EmployeeNAVCode', CD.BatchID,CD.ApprovalDate , GD.ClaimID From Txn_GodownRent CD"
''   tmp = tmp & " Inner Join Txn_GodownRentDetail GD on CD.ClaimID = GD.ClaimID"
''   tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
''   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
''   tmp = tmp & " Inner Join Mst_Lessor E on CD.LessorID = E.LessorID"
''   tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
''   tmp = tmp & " Inner Join Mst_ExpenseType ET on ET.ExpenseTypeId = 1"
''   tmp = tmp & " left join Mst_TDS MT on E.TDSCode=MT.TDSID"
''   tmp = tmp & " left  Join Mst_ServiceTax ST on ET.ServiceTaxID = ST.ServiceTaxID"
''   tmp = tmp & " Where CD.BatchID =  0 and  CD.Flag = 'A'"
''   If CmbApprovedBy.Text <> "" Then
''      tmp = tmp & " And CD.ForwardToID=" & mForwardToID & ""
''   End If
''   tmp = tmp & " Group by CD.ApprovalDate,ET.NAVCode,ST.NAVCode ,L.NAVCode, CD.BatchID,CD.ApprovalDate, MT.NavCode,GD.ClaimID"
''   tmp = tmp & " Order By GD.ClaimID,CD.ApprovalDate"

   tmp = "Select CD.ApprovalDate,'Invoice' As Inv,'G/L Account' As GLAccount,ET.NAVCode 'ExpenseTypeNAVCode',"
   tmp = tmp & " Sum (GD.Amount + GD.Arrears) 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode',"
   tmp = tmp & " 'Being warehouse rent payable against Storeman Batch ID " & MaxID & "' As 'ApprovalRemarks',"
   tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
   tmp = tmp & " '' 'EmployeeNAVCode', CD.BatchID,CD.ApprovalDate , GD.ClaimID,ST.Description "
   tmp = tmp & " From Txn_GodownRent CD "
   tmp = tmp & " Inner Join Txn_GodownRentDetail GD on CD.ClaimID = GD.ClaimID"
   tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Lessor E on CD.LessorID = E.LessorID"
   tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
   tmp = tmp & " Inner Join Txn_LessorAgreement TLA On GD.LAID = TLA.LAID "
   tmp = tmp & " Inner Join Mst_ExpenseType ET on ET.ExpenseTypeId = 1"
   tmp = tmp & " left join Mst_TDS MT on E.TDSID=MT.TDSID"
   ''tmp = tmp & " left  Join Mst_ServiceTax ST on ET.ServiceTaxID = ST.ServiceTaxID"
   tmp = tmp & " left  Join Mst_ServiceTax ST on TLA.ServiceTaxID = ST.ServiceTaxID "
   tmp = tmp & " Where CD.BatchID =  0 and  CD.Flag = 'A'"
   If CmbApprovedBy.Text <> "" Then
      tmp = tmp & " And CD.ForwardToID=" & mForwardToID & ""
   End If
   tmp = tmp & " Group by CD.ApprovalDate,ET.NAVCode,ST.NAVCode ,L.NAVCode, CD.BatchID,CD.ApprovalDate, MT.NavCode,GD.ClaimID,ST.Description"
   tmp = tmp & " Order By GD.ClaimID,CD.ApprovalDate"

        
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
        
''   tmp = "Select CD.ApprovalDate,'Invoice' As Inv,'G/L Account' As GLAccount,ET.NAVCode 'ExpenseTypeNAVCode',"
''   tmp = tmp & " Sum (GD.Amount + GD.Arrears) 'Amount', '' As TDS,'' 'ServiceTaxNAVCode'," 'ST.NAVCode 'ServiceTaxNAVCode',"
''   tmp = tmp & " 'Being warehouse rent payable against Storeman Batch ID " & CmbBatch.Text & "' As 'ApprovalRemarks',"
''   tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
''   tmp = tmp & " '' 'EmployeeNAVCode', CD.BatchID,CD.ApprovalDate,GD.ClaimID  From Txn_GodownRent CD"
''   tmp = tmp & " Inner Join Txn_GodownRentDetail GD on CD.ClaimID = GD.ClaimID"
''   tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
''   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
''   tmp = tmp & " Inner Join Mst_Lessor E on CD.LessorID = E.LessorID"
''   tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
''   tmp = tmp & " Inner Join Mst_ExpenseType ET on ET.ExpenseTypeId = 1"
''   tmp = tmp & " left join Mst_TDS MT on E.TDSCode=MT.TDSID"
''   tmp = tmp & " left  Join Mst_ServiceTax ST on ET.ServiceTaxID = ST.ServiceTaxID"
''   tmp = tmp & " Where CD.BatchID = " & Val(CmbBatch.Text) & " and  CD.Flag = 'A'"
''   tmp = tmp & " Group by CD.ApprovalDate,ET.NAVCode,ST.NAVCode ,L.NAVCode, CD.BatchID,CD.ApprovalDate, MT.NavCode,GD.ClaimID"
''   tmp = tmp & " Order By GD.ClaimID,CD.ApprovalDate"
   
   tmp = "Select CD.ApprovalDate,'Invoice' As Inv,'G/L Account' As GLAccount,ET.NAVCode 'ExpenseTypeNAVCode',"
   tmp = tmp & " Sum (GD.Amount + GD.Arrears) 'Amount', '' As TDS,ST.NAVCode 'ServiceTaxNAVCode',"
   tmp = tmp & " 'Being warehouse rent payable against Storeman Batch ID " & CmbBatch.Text & "' As 'ApprovalRemarks',"
   tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
   tmp = tmp & " '' 'EmployeeNAVCode', CD.BatchID,CD.ApprovalDate,GD.ClaimID,ST.Description "
   tmp = tmp & " From Txn_GodownRent CD "
   tmp = tmp & " Inner Join Txn_GodownRentDetail GD on CD.ClaimID = GD.ClaimID"
   tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Lessor E on CD.LessorID = E.LessorID"
   tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID"
   tmp = tmp & " Inner Join Txn_LessorAgreement TLA On GD.LAID = TLA.LAID "
   tmp = tmp & " Inner Join Mst_ExpenseType ET on ET.ExpenseTypeId = 1"
   tmp = tmp & " left join Mst_TDS MT on E.TDSID=MT.TDSID"
   ''tmp = tmp & " left  Join Mst_ServiceTax ST on ET.ServiceTaxID = ST.ServiceTaxID"
   tmp = tmp & " left  Join Mst_ServiceTax ST on TLA.ServiceTaxID = ST.ServiceTaxID "
   tmp = tmp & " Where CD.BatchID = " & Val(CmbBatch.Text) & " and  CD.Flag in ('S', 'A') "
   tmp = tmp & " Group by CD.ApprovalDate,ET.NAVCode,ST.NAVCode ,L.NAVCode, CD.BatchID,CD.ApprovalDate, MT.NavCode,GD.ClaimID,ST.Description"
   tmp = tmp & " Order By GD.ClaimID,CD.ApprovalDate"
   
End If
Call TmpTable

If Option1.Value = True Then
    
''    tmp = "Select GR.ApprovalDate,'Invoice' As Inv,'Vendor' As GLAccount,ML.NAVCode 'ExpenseTypeNAVCode',"
''    tmp = tmp & " sum(GD.Amount + GD.Arrears) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode' ,"
''    tmp = tmp & " GR.ApprovalRemarks + ' - Claim ID:' + convert(varchar(10),GR.ClaimID) As 'ApprovalRemarks',"
''    tmp = tmp & " L.NAVCode  'LocationNAVCode','CMP' As DeptDim ,"
''    tmp = tmp & " ML.NAVCode 'LessorNAVCode', GR.BatchID ,'' As InvoiceDate,GD.ClaimID From  Txn_GodownRent GR"
''    tmp = tmp & " Inner Join Mst_Lessor ML on GR.LessorID = ML.LessorID"
''    tmp = tmp & " Left Join Mst_TDS T on ML.TDSCode = T.TDSID"
''    tmp = tmp & " inner Join Txn_GodownRentDetail GD on GR.ClaimID = GD.ClaimID "
''    tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID "
''    tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID "
''    tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID "
''    tmp = tmp & " Where GR.BatchID =  0 and  GR.Flag = 'A'"
''    If CmbApprovedBy.Text <> "" Then
''       tmp = tmp & " And GR.ForwardToID=" & mForwardToID & ""
''    End If
''    tmp = tmp & " group by GR.ApprovalDate,ML.NAVCode ,T.NavCode  ,GR.ApprovalRemarks ,GR.ClaimID, L.NAVCode ,  GR.BatchID,GD.ClaimID"
''    tmp = tmp & " Order By GD.ClaimID,GR.ApprovalDate"

    tmp = "Select GR.ApprovalDate,'Invoice' As Inv,'Vendor' As GLAccount,MLGST.GSTNavCode 'ExpenseTypeNAVCode',"
    tmp = tmp & " sum(GD.Amount + GD.Arrears) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode' ,"
    tmp = tmp & " GR.ApprovalRemarks + ' - Claim ID:' + convert(varchar(10),GR.ClaimID) As 'ApprovalRemarks',"
    tmp = tmp & " L.NAVCode  'LocationNAVCode','CMP' As DeptDim ,"
    tmp = tmp & " ML.NAVCode 'LessorNAVCode', GR.BatchID ,'' As InvoiceDate,GD.ClaimID "
    tmp = tmp & " From  Txn_GodownRent GR "
    tmp = tmp & " Inner Join Mst_Lessor ML on GR.LessorID = ML.LessorID"
    tmp = tmp & " Left Join Mst_TDS T on ML.TDSID = T.TDSID"
    tmp = tmp & " inner Join Txn_GodownRentDetail GD on GR.ClaimID = GD.ClaimID "
    tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID "
    tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID "
    tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID "
    ''-- Ullhas Patil 26 Oct 2017
    ''tmp = tmp & " Inner Join Mst_State MS On L.StateID = MS.StateID "
    ''tmp = tmp & " Left Join Mst_LessorStateWiseGST MLGST On MSI.StateIndiaID = MLGST.StateIndiaID And ML.LessorID = MLGST.LessorID "
    tmp = tmp & " Left Join Mst_LessorStateWiseGST MLGST On GR.BillingStateIndiaID = MLGST.StateIndiaID And ML.LessorID = MLGST.LessorID"
    ''--------------------------------------
    
    tmp = tmp & " Where GR.BatchID =  0 and  GR.Flag = 'A'"
    If CmbApprovedBy.Text <> "" Then
       tmp = tmp & " And GR.ForwardToID=" & mForwardToID & ""
    End If
    tmp = tmp & " group by GR.ApprovalDate,MLGST.GSTNavCode,T.NavCode,GR.ApprovalRemarks,GR.ClaimID,L.NAVCode,ML.NAVCode,GR.BatchID,GD.ClaimID "
    tmp = tmp & " Order By GD.ClaimID,GR.ApprovalDate"

Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If

''    tmp = "Select GR.ApprovalDate,'Invoice' As Inv,'Vendor' As GLAccount,ML.NAVCode 'ExpenseTypeNAVCode',"
''    tmp = tmp & " Sum(GD.Amount  + GD.Arrears) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',"
''    tmp = tmp & " GR.ApprovalRemarks + ' - Claim ID:' + convert(varchar(10),GR.ClaimID) As 'ApprovalRemarks',"
''    tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
''    tmp = tmp & " ML.NAVCode 'LessorNAVCode', GR.BatchID ,'' As InvoiceDate,GD.ClaimID From  Txn_GodownRent GR"
''    tmp = tmp & " Inner Join Mst_Lessor ML on GR.LessorID = ML.LessorID"
''    tmp = tmp & " Left Join Mst_TDS T on ML.TDSCode = T.TDSID"
''    tmp = tmp & " inner Join Txn_GodownRentDetail GD on GR.ClaimID = GD.ClaimID "
''    tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID "
''    tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID "
''    tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID "
''    tmp = tmp & " Where GR.BatchID = " & Val(CmbBatch.Text) & " and  GR.Flag in ('S', 'A')"
''    tmp = tmp & " group by GR.ApprovalDate,ML.NAVCode ,T.NavCode  ,GR.ApprovalRemarks ,GR.ClaimID, L.NAVCode ,  GR.BatchID,GD.ClaimID"
''    tmp = tmp & " Order By GD.ClaimID,GR.ApprovalDate"

    tmp = "Select GR.ApprovalDate,'Invoice' As Inv,'Vendor' As GLAccount,MLGST.GSTNavCode 'ExpenseTypeNAVCode',"
    tmp = tmp & " Sum(GD.Amount  + GD.Arrears) 'Amount', T.NavCode  As TDS,'' As 'ServiceTaxNAVCode',"
    tmp = tmp & " GR.ApprovalRemarks + ' - Claim ID:' + convert(varchar(10),GR.ClaimID) As 'ApprovalRemarks',"
    tmp = tmp & " L.NAVCode 'LocationNAVCode','CMP' As DeptDim ,"
    tmp = tmp & " ML.NAVCode 'LessorNAVCode', GR.BatchID ,'' As InvoiceDate,GD.ClaimID From  Txn_GodownRent GR"
    tmp = tmp & " Inner Join Mst_Lessor ML on GR.LessorID = ML.LessorID"
    tmp = tmp & " Left Join Mst_TDS T on ML.TDSID = T.TDSID"
    tmp = tmp & " inner Join Txn_GodownRentDetail GD on GR.ClaimID = GD.ClaimID "
    tmp = tmp & " Inner Join Mst_Godown MG on GD.GodownID = MG.GodownID "
    tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID "
    tmp = tmp & " Inner Join Mst_Location L on MC.LocationID = L.LocationID "
    ''--- Ullhas Patil 26 Oct 2017
    ''tmp = tmp & " Inner Join Mst_State MS On L.StateID = MS.StateID "
    ''tmp = tmp & " Left Join Mst_LessorStateWiseGST MLGST On MSI.StateIndiaID = MLGST.StateIndiaID And ML.LessorID = MLGST.LessorID "
    tmp = tmp & " Left Join Mst_LessorStateWiseGST MLGST On GR.BillingStateIndiaID = MLGST.StateIndiaID And ML.LessorID = MLGST.LessorID"
    ''-----------------
    tmp = tmp & " Where GR.BatchID = " & Val(CmbBatch.Text) & " and  GR.Flag in ('S', 'A')"
    tmp = tmp & " group by GR.ApprovalDate,MLGST.GSTNavCode ,T.NavCode  ,GR.ApprovalRemarks ,GR.ClaimID, L.NAVCode,ML.NAVCode,GR.BatchID,GD.ClaimID"
    tmp = tmp & " Order By GD.ClaimID,GR.ApprovalDate"

End If
Call Tmp1Table
Dim mDat_ As Variant
'If TmpRs.RecordCount <> TmpRs1.RecordCount Then
'   MsgBox "Data Mismatch Contact Technology!!!"
'   Exit Sub
'End If
If TmpRs.RecordCount > 0 And TmpRs1.RecordCount > 0 Then
    Dim x As Variant
    x = 0
    ExpFlag = "Y"
    mDat_ = IIf(IsNull(TmpRs.Fields(0)), "", TmpRs.Fields(0))
    For x = 1 To TmpRs.RecordCount
        If TmpRs!ExpenseTypeNAVCode = "" Or IsNull(TmpRs!ExpenseTypeNAVCode) Or TmpRs!LocationNAVCode = "" Or IsNull(TmpRs!LocationNAVCode) Then
           ExpFlag = "N"
        End If
        
''---- Ullhas Patil 12 Oct 2017
''        i = MSHFlexGrid1.Rows - 1
''        For c = 0 To MSHFlexGrid1.Cols - 1
''            MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
''        Next
''        MSHFlexGrid1.TextMatrix(i, 7) = MSHFlexGrid1.TextMatrix(i, 7) & " Claim ID : " & TmpRs!ClaimID
''        'MSHFlexGrid1.TextMatrix(i, 0) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 0))
''        MSHFlexGrid1.TextMatrix(i, 10) = "EMP-1569"
''        'MSHFlexGrid1.TextMatrix(i, 12) = ConvertToyyyymmdd(MSHFlexGrid1.TextMatrix(i, 12))
''        'CmbBatch.Text = ""
''        If Option1.Value = True Then
''           MSHFlexGrid1.TextMatrix(i, 11) = "SMG" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(i, 0), "yyyymmdd")
''        Else
''           MSHFlexGrid1.TextMatrix(i, 11) = "SMG" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(i, 0), "yyyymmdd")
''        End If
''        If Option1.Value = True Then
''            mDocumentNo = "DocNo-" & TmpRs!ClaimID
''            MSHFlexGrid1.TextMatrix(i, 13) = mDocumentNo
''        End If
        
        I = MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!ApprovalDate), "", TmpRs!ApprovalDate)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Inv), "", TmpRs!Inv)
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!GLAccount), "", TmpRs!GLAccount)
        If LCase(TmpRs!Description) = LCase("GST with Tax") Then
            MSHFlexGrid1.TextMatrix(I, 3) = GetGLAccNAVCodeFromExpenseType(mWarehouseRentExpenseID, "NAGRI", "G")
        ElseIf LCase(TmpRs!Description) = LCase("GST without Tax") Then
            MSHFlexGrid1.TextMatrix(I, 3) = GetGLAccNAVCodeFromExpenseType(mWarehouseRentExpenseID, "AGRI", "G")
        End If
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount)
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!TDS), "", TmpRs!TDS)
        MSHFlexGrid1.TextMatrix(I, 6) = "" 'IIf(IsNull(TmpRs!ServiceTaxNAVCode), "", TmpRs!ServiceTaxNAVCode)
        MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!ApprovalRemarks), "", TmpRs!ApprovalRemarks) & " Claim ID : " & TmpRs!ClaimID
        MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!LocationNAVCode), "", TmpRs!LocationNAVCode)
        MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DeptDim), "", TmpRs!DeptDim)
        MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
        If Option1.Value = True Then
           MSHFlexGrid1.TextMatrix(I, 11) = "SMG" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd") & "-" & TmpRs!ClaimID
        Else
           MSHFlexGrid1.TextMatrix(I, 11) = "SMG" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd") & "-" & TmpRs!ClaimID
        End If
        MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!ApprovalDate), "", TmpRs!ApprovalDate)
        MSHFlexGrid1.TextMatrix(I, 13) = ""
        mDocumentNo = "DocNo-" & TmpRs!ClaimID
        MSHFlexGrid1.TextMatrix(I, 17) = mDocumentNo
        MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!ServiceTaxNAVCode), "", TmpRs!ServiceTaxNAVCode)
        If LCase(TmpRs!Description) = LCase("GST with Tax") Then
            MSHFlexGrid1.TextMatrix(I, 15) = GetGLAccNAVCodeFromExpenseType(mWarehouseRentExpenseID, "NAGRI", "H")
        ElseIf LCase(TmpRs!Description) = LCase("GST without Tax") Then
            MSHFlexGrid1.TextMatrix(I, 15) = GetGLAccNAVCodeFromExpenseType(mWarehouseRentExpenseID, "AGRI", "H")
        End If
        MSHFlexGrid1.TextMatrix(I, 16) = ""
        
        If MSHFlexGrid1.TextMatrix(I, 14) = "" Or MSHFlexGrid1.TextMatrix(I, 15) = "" Then
            ExpFlag = "N"
        End If
        
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'        Call CrEntry
        
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
''If TmpRs1!ExpenseTypeNAVCode = "" Or IsNull(TmpRs1!ExpenseTypeNAVCode) Or TmpRs1!LocationNAVCode = "" Or IsNull(TmpRs1!LocationNAVCode) Then
''   ExpFlag = "N"
''End If

TmpRs1.MoveFirst
TmpRs1.Find "ClaimID = " & mPDat_ 'TmpRs!BillTxnID
If TmpRs1.EOF = False Then
   If TmpRs1!ExpenseTypeNAVCode = "" Or IsNull(TmpRs1!ExpenseTypeNAVCode) Or TmpRs1!LocationNAVCode = "" Or IsNull(TmpRs1!LocationNAVCode) Or TmpRs1!LessorNAVCode = "" Or IsNull(TmpRs1!LessorNAVCode) Then
      ExpFlag = "N"
   End If

   I = MSHFlexGrid1.Rows - 1
   MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!ApprovalDate), "", TmpRs1!ApprovalDate)
   MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!Inv), "", TmpRs1!Inv)
   MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!GLAccount), "", TmpRs1!GLAccount)
   MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!ExpenseTypeNAVCode), "", TmpRs1!ExpenseTypeNAVCode)
   MSHFlexGrid1.TextMatrix(I, 4) = Val(IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount)) * -1
   MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!TDS), "", TmpRs1!TDS)
   MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!ServiceTaxNAVCode), "", TmpRs1!ServiceTaxNAVCode)
   MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!ApprovalRemarks), "", TmpRs1!ApprovalRemarks)
   MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!LocationNAVCode), "", TmpRs1!LocationNAVCode)
   MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!DeptDim), "", TmpRs1!DeptDim)
   MSHFlexGrid1.TextMatrix(I, 10) = "EMP-1569"
   If Option1.Value = True Then
      MSHFlexGrid1.TextMatrix(I, 11) = "SMG" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd") & "-" & TmpRs1!ClaimID
   Else
      MSHFlexGrid1.TextMatrix(I, 11) = "SMG" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd") & "-" & TmpRs1!ClaimID
   End If
   MSHFlexGrid1.TextMatrix(I, 12) = TmpRs1!ApprovalDate
   MSHFlexGrid1.TextMatrix(I, 13) = ""
   MSHFlexGrid1.TextMatrix(I, 14) = ""
   MSHFlexGrid1.TextMatrix(I, 15) = ""
   MSHFlexGrid1.TextMatrix(I, 16) = ""
   MSHFlexGrid1.TextMatrix(I, 17) = mDocumentNo
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'TmpRs1.MoveNext
End If

End Sub

Sub CrEntryOld(mPDat_ As Variant)

For x = 1 To TmpRs1.RecordCount
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
    
    If CmbBatch.Text = "" Then
       MSHFlexGrid1.TextMatrix(I, 11) = "SMG" & Trim(MaxID) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
    Else
       MSHFlexGrid1.TextMatrix(I, 11) = "SMG" & Trim(CmbBatch.Text) & "-" & Format(MSHFlexGrid1.TextMatrix(I, 0), "yyyymmdd")
    End If
    
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    TmpRs1.MoveNext
    If TmpRs1.EOF Then Exit For
    If mPDat_ <> IIf(IsNull(TmpRs1.Fields(0)), "", TmpRs1.Fields(0)) Then Exit For
Next
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
mWarehouseRentExpenseID = 1
Call Grid_Head
tmp = "Select Distinct BatchID from Txn_GodownRent Where BatchID > 0  order by BatchID"
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
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
Else
   CmbBatch.Enabled = False
   CmbApprovedBy.Text = ""
   CmbApprovedBy.Enabled = True
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
End If
End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Public Sub Grid_Head()
CmdExport.Enabled = False
With MSHFlexGrid1
    .Clear
    .Rows = 2
    ''---- Ullhas Patil 12 Oct 2017
    ''If Option1.Value = True Then .Cols = 14 Else .Cols = 13
    .Cols = 18 '17
    
    .TextMatrix(0, 0) = "Posting Date":                 .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "Document Type":                .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "Account Type":                 .ColWidth(2) = 1200
    .TextMatrix(0, 3) = "Account No.":                  .ColWidth(3) = 1500
    .TextMatrix(0, 4) = "Amount":                       .ColWidth(4) = 1200
    .TextMatrix(0, 5) = "TDS Nature of Deduction":      .ColWidth(5) = 800
    .TextMatrix(0, 6) = "Service Tax Group Code":       .ColWidth(6) = 1800
    .TextMatrix(0, 7) = "Narration":                    .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "RegionDim":                    .ColWidth(8) = 1500
    .TextMatrix(0, 9) = "DeptDim":                      .ColWidth(9) = 1000
    .TextMatrix(0, 10) = "LessorDim":                   .ColWidth(10) = 1000
    .TextMatrix(0, 11) = "External Document No.":       .ColWidth(11) = 1200
    .TextMatrix(0, 12) = "Invoice Date":                .ColWidth(12) = 1000
    ''---- Ullhas Patil 12 Oct 2017
    ''If Option1.Value = True Then .TextMatrix(0, 13) = "Document No": .ColWidth(13) = 1000
    .TextMatrix(0, 13) = "Project Dim":                 .ColWidth(13) = 1000
    .TextMatrix(0, 14) = "GSTGroup Code":               .ColWidth(14) = 1000
    .TextMatrix(0, 15) = "HSN Code":                    .ColWidth(15) = 1000
    .TextMatrix(0, 16) = "Invoice Type":                .ColWidth(16) = 1000
    .TextMatrix(0, 17) = "Document No":                     .ColWidth(16) = 1800
    .RowHeight(0) = 1000
End With
End Sub


Private Sub CmbApprovedBy_GotFocus()

tmp1 = CmbApprovedBy.Text
'Call TableMst_Employee("Where Flag = 'Y' And DesignationID=")
tmp = "Select ME.EmployeeID,ME.EmployeeNo,Me.EmployeeName From Mst_Employee ME"
tmp = tmp & " Inner join Mst_Users MU On MU.EmployeeNo=ME.EmployeeNo Where MU.UserRoleID in (7,33) And ME.Flag = 'Y'"

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



