VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ExportReservationInvoice 
   Caption         =   "Export Reservation Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8025
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7680
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13547
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10440
         TabIndex        =   21
         Top             =   720
         Width           =   1215
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
         Left            =   6720
         TabIndex        =   7
         Top             =   720
         Width           =   3600
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Invoice Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   315
         Width           =   1455
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
         Height          =   350
         Left            =   11880
         TabIndex        =   11
         Top             =   600
         Width           =   1140
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
         Height          =   350
         Left            =   11520
         TabIndex        =   8
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   13020
         TabIndex        =   12
         Top             =   600
         Width           =   1140
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Left            =   13800
         TabIndex        =   10
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
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
         Left            =   12660
         TabIndex        =   9
         Top             =   240
         Width           =   1140
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
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
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
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   675
         Width           =   855
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
         Left            =   6720
         TabIndex        =   5
         Top             =   247
         Width           =   1320
      End
      Begin VB.ComboBox CmbInvoiceType 
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
         ItemData        =   "Frm_ExportReservationInvoice.frx":0000
         Left            =   9480
         List            =   "Frm_ExportReservationInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   247
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   3945
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Format          =   113573889
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   3945
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Format          =   113573889
         CurrentDate     =   39884
      End
      Begin VB.Label Label4 
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
         Left            =   5400
         TabIndex        =   20
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label23 
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
         Height          =   225
         Left            =   2880
         TabIndex        =   19
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label14 
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
         Height          =   225
         Left            =   2880
         TabIndex        =   18
         Top             =   780
         Width           =   810
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
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Invoice Type"
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
         Left            =   8160
         TabIndex        =   15
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Frm_ExportReservationInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCode, mSTCode As Variant
Dim ExpFlag, MaxID As Variant
Dim mSMonth, mSYear, mSTotalDays, mForwardToID As Variant
Dim mEmployeeName, mEmployeeNo As Variant
Dim mJournalBatchName, mJournalType, mInvoiceTypeID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub CmbBatch_GotFocus()
CmbInvoiceType.Text = ""
'Call TableTxn_ReservationInvoice("select distinct(BatchID) As BatchID from Txn_ReservationInvoice Where BatchID > 0  order by BatchID")
tmp = "select distinct BatchID As BatchID from Txn_ReservationInvoice Where BatchID > 0  order by BatchID"
Call TmpTable
CmbBatch.Clear
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       CmbBatch.AddItem TmpRs!BatchID
       MaxID = Val(TmpRs!BatchID) + 1
       TmpRs.MoveNext
   Next
Else
   MaxID = 1
End If
End Sub

Private Sub CmbBatch_LostFocus()
If CmbBatch.Text = "" Then
   Exit Sub
End If
 If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
   
   TmpRs.Find "BatchID = " & Val(CmbBatch.Text) & ""
   If TmpRs.EOF Then
      MsgBox "Invalid selection "
      CmbBatch.SetFocus
      Exit Sub
   End If
 

End Sub

Private Sub CmbInvoiceType_LostFocus()
If CmbInvoiceType.Text = "" Then Exit Sub
If CmbInvoiceType.Text <> "Non Taxable" And CmbInvoiceType.Text <> "Taxable" Then
   MsgBox "Invalid Selection"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbJournalCode_GotFocus()
Dim mR As Variant
tmp1 = CmbJournalCode
CmbJournalCode.Clear
'tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
'tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
'tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '" & Gen_UserLoginID & "' Order by JCM.JournalTypeID"

tmp = " SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]" '''JT.JournalTypeName,JCM.JournalCode
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where (ME.EmployeeNo = '" & Gen_UserLoginID & "' Or ME.EmployeeNo = '4474') And JCM.JournalTypeID = 2  Order by JCM.JournalTypeID"



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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Export_ReservationInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Export_ReservationInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid3.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid3.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant

''-- On 05 Aug 2016
Dim mUserBatchID, mInvoiceType As Variant
'mJournalType = "SALES"
'mUserBatchID = GetGUID
'mJournalBatchName = GetJournalDetailsFromHRPro(Gen_UserLoginID, mJournalType)
'If Trim(mJournalBatchName) = "" Then
'   MsgBox "No Journal Batch Name mapped to " & Gen_UserName & " - " & Gen_UserLoginID & "  in HR Pro . Please map the same"
'   Exit Sub
'End If
''-- On 05 Aug 2016
'Commented on 14 Aug 2017 during GST Effect implementaion
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

Path = App.Path
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_ReservationInvoice" & "_" & CmbInvoiceType.Text & "_" & MaxID & " .csv"
Else
   strFilName = "\Navision\NavUpload_ReservationInvoice" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
End If

''-- On 05 Aug 2016
If Option2.Value = True Then
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
Else
   'Commented on 14 Aug 2017 during GST Effect implementaion
    Call TableTxn_InvoicesForNAVISION(" Where G_UID Is Null ")
    strFilLocation = Path + strFilName
    If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
    Open strFilLocation For Output As #1
End If

''-- Ullhas Patil 09 Nov 2017
If Option1.Value = True Then
   If CmbInvoiceType.Text = "Non Taxable" Then
      tmp = "Select * from Txn_ReservationInvoice Where Isnull(BatchID,0) = 0 And ServiceTaxNAVCode='GSTS00' And Flag='A' And BillingDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And BillingDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
      Call Tmp1Table
   Else
      tmp = "Select * from Txn_ReservationInvoice Where Isnull(BatchID,0) = 0 And ServiceTaxNAVCode='GSTS18' And Flag='A' And BillingDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And BillingDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  "
      Call Tmp1Table
   End If
End If
''--------------------------

With MSHFlexGrid3
For I = 2 To .Rows - 1
    tmp = ""
    If Option1.Value = True Then
    
'Commented on 14 Aug 2017 during GST Effect implementaion
       If .TextMatrix(I, 25) = strChecked Then
            RsTxn_InvoicesForNAVISION.AddNew
            RsTxn_InvoicesForNAVISION!Date = Trim(.TextMatrix(I, 0))
            RsTxn_InvoicesForNAVISION!Invoice = Trim(.TextMatrix(I, 1))
            RsTxn_InvoicesForNAVISION!Customer = Trim(.TextMatrix(I, 2))
            RsTxn_InvoicesForNAVISION!NavCode = Trim(.TextMatrix(I, 3))
            RsTxn_InvoicesForNAVISION!Amount = Trim(.TextMatrix(I, 4))
            RsTxn_InvoicesForNAVISION!GLAccount = Trim(.TextMatrix(I, 5))
            RsTxn_InvoicesForNAVISION!GLACNAVCode = Trim(.TextMatrix(I, 6))
            RsTxn_InvoicesForNAVISION!Narration = Trim(.TextMatrix(I, 12))
            RsTxn_InvoicesForNAVISION!LocationNAVCode = Trim(.TextMatrix(I, 13))
            RsTxn_InvoicesForNAVISION!DeptDim = Trim(.TextMatrix(I, 14))
            RsTxn_InvoicesForNAVISION!PINo = Trim(.TextMatrix(I, 18))
            RsTxn_InvoicesForNAVISION!STCode = Trim(.TextMatrix(I, 19))
            RsTxn_InvoicesForNAVISION!AgreeNonAgree = Trim(.TextMatrix(I, 20))
            
            RsTxn_InvoicesForNAVISION!GSTGroupCode = Trim(.TextMatrix(I, 21))
            RsTxn_InvoicesForNAVISION!HSNNo = Trim(.TextMatrix(I, 22))
            RsTxn_InvoicesForNAVISION!AccountType = Trim(.TextMatrix(I, 23))
            RsTxn_InvoicesForNAVISION!NumberSeriesFlag = Trim(.TextMatrix(I, 24))
            
            RsTxn_InvoicesForNAVISION!CreatedBy = Gen_UserLoginID
            RsTxn_InvoicesForNAVISION!G_UID = GetGUID
            RsTxn_InvoicesForNAVISION!BatchID = IIf(Option1.Value = True, MaxID, CmbBatch.Text)
            'RsTxn_InvoicesForNAVISION!DataTransfer = Null
            RsTxn_InvoicesForNAVISION!InvoiceType = "RESERVATION"
            RsTxn_InvoicesForNAVISION!JournalBatchType = Trim(mJournalType)
            RsTxn_InvoicesForNAVISION!JournalBatchName = Trim(mJournalBatchName)
            'RsTxn_InvoicesForNAVISION!ReverseInvoice = Null
            'RsTxn_InvoicesForNAVISION!EntryNoTransfer = Null
            'RsTxn_InvoicesForNAVISION!EntryNoTransferRemark = Null
            RsTxn_InvoicesForNAVISION!UserBatchID = mUserBatchID
            
            ' For integration with Business Central - Start - 10-10-2023
            RsTxn_InvoicesForNAVISION!DataTransfer = "T"
            ' For integration with Business Central - End - 10-10-2023
            
            RsTxn_InvoicesForNAVISION.Update
        
        
    '        For c = 0 To .Cols - 1
    '            If tmp = "" Then
    '               tmp = tmp & .TextMatrix(i, c)
    '            Else
    '               tmp = tmp & "," & .TextMatrix(i, c)
    '            End If
    '        Next
    '        Print #1, tmp
                
              ''-- Ullhas Patil 09 Nov 2017 -- Save the GSTGroupCode, HSN Code, Invoice Type
              If Option1.Value = True Then
                If TmpRs1.RecordCount > 0 Then
                   TmpRs1.MoveFirst
                   TmpRs1.Find "InvoiceNo = " & Val(Replace(MSHFlexGrid3.TextMatrix(I, 18), "RI", "")) & " "
                   If TmpRs1.EOF = False Then
                      TmpRs1!BatchID = MaxID
                      TmpRs1!HSNNo = MSHFlexGrid3.TextMatrix(I, 22)
                      TmpRs1!AccountType = MSHFlexGrid3.TextMatrix(I, 23)
                      TmpRs1.Update
                   End If
                End If
              End If
       End If   ''----------------------------
    Else
        For C = 0 To .Cols - 1
            If tmp = "" Then
               tmp = tmp & .TextMatrix(I, C)
            Else
               tmp = tmp & "," & .TextMatrix(I, C)
            End If
        Next
        Print #1, tmp
    End If
Next
End With

''''-- Ullhas Patil 09 Nov 2017
''If Option1.Value = True Then
''   If CmbInvoiceType.Text = "Non Taxable" Then
''      tmp = "Update Txn_ReservationInvoice set BatchID = " & MaxID & " Where BatchID = 0  And ServiceTaxNAVCode='GSTS00' And Flag='A' And BillingDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And BillingDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
''      Call TmpTable
''   Else
''      tmp = "Update Txn_ReservationInvoice set BatchID = " & MaxID & " Where BatchID = 0  And ServiceTaxNAVCode='GSTS18' And Flag='A' And BillingDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And BillingDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'  "
''      Call TmpTable
''   End If
''End If

'MsgBox "Done !!!"
MsgBox IIf(Option1.Value = True, "New", "Old") & " Batch No - " & IIf(Option1.Value = True, MaxID, CmbBatch.Text) & " is Exported."
Call Grid_Head
Close #1
End Sub
Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_ReservationInvoice"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function
Private Sub CmdGetReport_Click()
Dim wc As Variant
ExpFlag = ""
If Option1.Value = True Then
   
   If Check1.Value = 0 Then
      MsgBox "Select Date filter."
      Check1.SetFocus
      Exit Sub
   End If
  
   
   If CmbInvoiceType.Text = "" Then
      MsgBox "Select Invoice Type."
      CmbInvoiceType.SetFocus
      Exit Sub
   End If
   
   If CmbInvoiceType.Text = "Non Taxable" Then
      wc = " where  TPI.ActualAmount > 0 And TPi.ServiceTaxNAVCode ='GSTS00'  And Isnull(TPI.BatchID,0) = 0"
   ElseIf CmbInvoiceType.Text = "Taxable" Then
      wc = " where TPI.ActualAmount > 0 And  TPi.ServiceTaxNAVCode = 'GSTS18'  And Isnull(TPI.BatchID,0) = 0"
   End If
   
 
'   If CmbApprovedBy.Text <> "" Then
'      Dim Str_() As String
'      Str_ = Split(CmbApprovedBy.Text, "~")
'      If UBound(Str_) = 1 Then
'         mEmployeeNo = Str_(1)
'      End If
'
'      wc = wc & " And ME.EmployeeNo= " & mEmployeeNo & " "
'   End If
   
   
   If Check1.Value = 1 Then
      If Gen_DBType = "MDB" Then
         wc = wc & "And TPI.BillingDate >= #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And TPI.BillingDate <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
      Else
         wc = wc & "And TPI.BillingDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TPI.BillingDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
      End If
   End If
   
   wc = wc & " And TPI.Flag='A'"
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      CmbBatch.SetFocus
      Exit Sub
   End If
    wc = " where   TPI.BatchID=" & Val(CmbBatch.Text) & ""
End If

Call Grid_Head

' mGLAccNAVCode = GetGLAccNAVCode("ADHOC")
'mGLAccNAVCode = GetGLAccNAVCode("Reservation")

'tmp = " Select TPI.InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MCL.ClientName,MC.NAVCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,MS.NAVCode, TPI.PIDate"
'tmp = tmp & " From Txn_ReservationInvoice TPI"
'tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
'
'tmp = tmp & " Inner Join Mst_ServiceTax MS on MS.ServiceTaxID=TPI.ServiceTaxID"
'tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MS.CommodityID"
'tmp = tmp & " inner join Mst_Client MCL on MCL.ClientIDRow=MS.ClientIDRow"
'''tmp = tmp & " Where  MS.NAVCode <> '' And TPI.BatchID = 0 And TPI.PIDate >= '2010/07/31' And TPI.PIDate <= '2011/07/31'"
'tmp = tmp & wc & " Order By TPI.PIDate"

'commented on 27072011 12.30PM
'tmp = " Select TPI.InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',"
'tmp = tmp & " MCL.ClientName , MCL.NAVCode 'ClientNAVCode',TPI.BillingAmount,TPI.BatchID,MS.NAVCode,"
'tmp = tmp & " TPI.BillingDate From Txn_ReservationInvoice TPI"
'tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
'tmp = tmp & " inner join Mst_StorageContract MSC on TPI.ContractID = MSc.ContractID"
'tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
'tmp = tmp & " Inner Join Mst_ServiceTax MS on MS.ServiceTaxID=MC.ServiceTaxID"
'tmp = tmp & " inner join Mst_Client MCL on MCL.ClientIDRow=MSC.ClientIDRow"
''tmp = tmp & " Where  MS.NAVCode <> '' And "
''tmp = tmp & " TPI.Flag='A' "
'tmp = tmp & wc & " Order By TPI.BillingDate"

tmp = " Select TPI.BillingDate,TPI.InvoiceNo,ML.LocationName,ML.NAVCode 'LocatioNAVCode',"
tmp = tmp & "  MCL.ClientName , MCSG.RGSTNAVCode 'ClientNAVCode',TPI.ActualAmount,"
tmp = tmp & " TPI.BatchID,TPI.ServiceTaxNAVCode 'NAVCode', TPI.BillingDate,TPI.Remarks,MC.HSNNo,MC.AccountType,MST.InvoiceType,isnull(MSC.StockManagmentContract,0) 'StockManagmentContract'  From Txn_ReservationInvoice TPI"
tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_StorageContract MSC on TPI.ContractID = MSc.ContractID"
tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
tmp = tmp & " left Join Mst_ServiceTax MST on MST.NAVCode=TPI.ServiceTaxNAVCode"
tmp = tmp & " inner join Mst_Client MCL on MCL.ClientID=MSC.ClientID"
'tmp = tmp & " inner join Mst_Employee ME on ME.EmployeeID=TPI.ForwardToID "
'tmp = tmp & " where  (MS.NAVCode <> '' or TPi.ServiceTaxNAVCode<>'') "
'tmp = tmp & " And TPI.BatchID = 0 And TPI.Flag='A' Order By TPI.BillingDate,invoiceno"
tmp = tmp & " inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " inner Join Mst_StateIndia MSI on MS.StateIndiaID = MSI.StateIndiaID"
tmp = tmp & " Left Join Mst_ClientStateWiseGST MCSG on MCL.ClientID = MCSG.ClientID And MSI.StateIndiaID = MCSG.StateIndiaID"

tmp = tmp & wc & " Order By TPI.BillingDate,InvoiceNo "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
'   If TmpRs!NAVCode <> "" Then
'      CmbInvoiceType.Text = "Taxable"
'   Else
'      CmbInvoiceType.Text = "Non Taxable"
'   End If
'
'   Frame5.Caption = CmbInvoiceType.Text
'
   For I = 1 To TmpRs.RecordCount
       If TmpRs!NavCode = "GSTS18" Then
          CmbInvoiceType.Text = "Taxable"
       Else
          CmbInvoiceType.Text = "Non Taxable"
       End If
   
       Frame5.Caption = CmbInvoiceType.Text
       
       If Option1.Value = True Then
          MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!BillingDate, "mm/dd/yyyy") 'MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!BillingDate + 1, "mm/dd/yyyy") 'Format(TmpRs!BillingDate, "dd/mm/yyyy")
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!BillingDate, "mm/dd/yyyy") 'Format(TmpRs!BillingDate, "dd/mm/yyyy")
       End If
       
       MSHFlexGrid3.TextMatrix(I + 1, 1) = "Invoice"
       MSHFlexGrid3.TextMatrix(I + 1, 2) = "Customer" 'TmpRs!ClientName
       MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
       
       MSHFlexGrid3.TextMatrix(I + 1, 4) = TmpRs!ActualAmount
       MSHFlexGrid3.TextMatrix(I + 1, 5) = "G/L Account"
       
       'MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCode = GetGLAccNAVCodeFromInvoiceType(mIncomeTypeID, "AGRI") 'mGLAccNAVCode
       
       If CmbInvoiceType.Text = "Non Taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 12) = TmpRs!Remarks '"Being Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!BillingDate)) & " " & Year(TmpRs!BillingDate) & " (Without ST) having Invoice No :- " & TmpRs!InvoiceNo & " "
          MSHFlexGrid3.TextMatrix(I + 1, 6) = GetGLAccNAVCodeFromInvoiceType(mInvoiceTypeID, "AGRI", "G")
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 12) = TmpRs!Remarks '"Being Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!BillingDate)) & " " & Year(TmpRs!BillingDate) & " (With ST) having Invoice No :- " & TmpRs!InvoiceNo & " "
          MSHFlexGrid3.TextMatrix(I + 1, 6) = GetGLAccNAVCodeFromInvoiceType(mInvoiceTypeID, "NAGRI", "G")
       End If
       MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!InvoiceNo), "", "RI" & TmpRs!InvoiceNo)
       MSHFlexGrid3.TextMatrix(I + 1, 13) = IIf(IsNull(TmpRs!LocatioNAVCode), "", TmpRs!LocatioNAVCode)
       MSHFlexGrid3.TextMatrix(I + 1, 14) = "CMP"
       MSHFlexGrid3.TextMatrix(I + 1, 21) = IIf(IsNull(TmpRs!NavCode), "", (TmpRs!NavCode))
       MSHFlexGrid3.TextMatrix(I + 1, 22) = IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo)
       MSHFlexGrid3.TextMatrix(I + 1, 23) = IIf(IsNull(TmpRs!InvoiceType), "", TmpRs!InvoiceType)
       MSHFlexGrid3.TextMatrix(I + 1, 24) = "M"
       
       If TmpRs!StockManagmentContract = False Then
            If MSHFlexGrid3.TextMatrix(I + 1, 22) = "" Then
               If CmbInvoiceType.Text = "Non Taxable" Then
                  MSHFlexGrid3.TextMatrix(I + 1, 22) = GetGLAccNAVCodeFromInvoiceType(mInvoiceTypeID, "AGRI", "H")
               Else
                  MSHFlexGrid3.TextMatrix(I + 1, 22) = GetGLAccNAVCodeFromInvoiceType(mInvoiceTypeID, "NAGRI", "H")
               End If
            End If
       Else
            MSHFlexGrid3.TextMatrix(I + 1, 22) = GetGLAccNAVCodeFromInvoiceType(mInvoiceTypeID, "NAGRI", "H")
       End If
       If LCase(CmbInvoiceType) = "non taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 20) = "AGRI"
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 20) = "NAGRI"
       End If
       ' following line of code is for showing sum of that column
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I + 1, 4))
       
       If IsNull(TmpRs!ClientNAVCode) = True Or TmpRs!ClientNAVCode = "" Then
          ExpFlag = "N"
       End If
       If IsNull(TmpRs!LocatioNAVCode) = True Or TmpRs!LocatioNAVCode = "" Then
          ExpFlag = "N"
       End If
       If MSHFlexGrid3.TextMatrix(I + 1, 6) = "" Then
          ExpFlag = "N"
       End If
       
        If Option1.Value = True Then
           MSHFlexGrid3.row = I + 1
           MSHFlexGrid3.Col = 25
           MSHFlexGrid3.CellFontName = "Wingdings"
           MSHFlexGrid3.CellFontSize = 12
           MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
           MSHFlexGrid3.Text = strUnChecked
        End If
       
       TmpRs.MoveNext
   
   
   
   Next
   

   
'   If ExpFlag = "N" Then
'      CmdExport.Enabled = False
'   Else
'      CmdExport.Enabled = True
'   End If
Else
   'CmdExport.Enabled = False
End If

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmbInvoiceType.Text = ""
'CmdExport.Enabled = False
Option1.Value = True
Call Option1_Click
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame5.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame5.Enabled = True
Call Grid_Head
Call CmdReset_Click
mInvoiceTypeID = 4


tmp = "SELECT JCM.EMployeeID,ME.EmployeeNo,ME.FullName,JCM.JournalTypeID,JT.JournalTypeName+'~'+JCM.JournalCode[JournalCode]"
tmp = tmp & " FROM Mst_EmpJournalCodeMapping JCM Inner Join Mst_JournalType JT On JCM.JournalTypeID = JT.JournalTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On JCM.EMployeeID = ME.EmployeeID Where ME.EmployeeNo = '4474' And JCM.JournalTypeID = 2 "
tmp = tmp & " Order by JCM.JournalTypeID"

Call TmpTableHRPro

If TmpRsHRPro.RecordCount = 1 Then
   CmbJournalCode.Text = TmpRsHRPro!JournalCode
'   mSForwardToID = TmpRs!EmployeeID
Else
   CmbJournalCode.Text = ""
'   mSForwardToID = Null
End If


mSplit = Split(CmbJournalCode.Text, "~")
mJournalType = UCase(mSplit(0))
mJournalBatchName = UCase(mSplit(1))

End Sub
Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    '.Cols = 25 '24 '21
    If Option1.Value = True Then
       .Cols = 26 '25
    Else
       .Cols = 25 '24
    End If
    
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Date"
    .TextMatrix(1, 1) = "Invoice"
    .TextMatrix(1, 2) = "Customer"
    .TextMatrix(1, 3) = "DPNavCode"
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "G/L Account"
    .TextMatrix(1, 6) = "GL A/C NAVCode "
    .TextMatrix(1, 12) = "Narration"
    .TextMatrix(1, 13) = "Location NAV Code"
    .TextMatrix(1, 14) = "CMP"
    .TextMatrix(1, 18) = "Invoice No"
    .TextMatrix(1, 19) = "ST Code"
    .TextMatrix(1, 20) = "Agri/Non Agri"
    .TextMatrix(1, 21) = "GST Code"
    .TextMatrix(1, 22) = "HSN No"
    .TextMatrix(1, 23) = "Invoice Type"
    .TextMatrix(1, 24) = "Number Series Flag"
    If Option1.Value = True Then
       .TextMatrix(1, 25) = "Selection"
    End If
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 900
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 900
    .ColWidth(5) = 1200
    .ColWidth(6) = 1600
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 2000
    .ColWidth(13) = 1600
    .ColWidth(14) = 550
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 900
    .ColWidth(19) = 0
    .ColWidth(20) = 1200
    .ColWidth(21) = 1200
    .ColWidth(22) = 1200
    .ColWidth(23) = 1200
    If Option1.Value = True Then
       .ColWidth(25) = 1000
    End If
End With

End Sub
Private Sub MSHFlexGrid3_Click()
With MSHFlexGrid3
     If .Col = 25 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 25
        If .TextMatrix(.row, .Col) = strUnChecked Then
           If .TextMatrix(.row, 3) <> "" And .TextMatrix(.row, 6) <> "" And .TextMatrix(.row, 13) <> "" And Val(.TextMatrix(.row, 4)) > 0 Then
              .TextMatrix(.row, .Col) = strChecked
           End If
        Else
           .TextMatrix(.row, .Col) = strUnChecked
        End If
     End If
End With
End Sub
Private Sub ChkCheckAll_Click()
If MSHFlexGrid3.Rows < 3 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.row = I
       If MSHFlexGrid3.TextMatrix(I, 3) <> "" And MSHFlexGrid3.TextMatrix(I, 6) <> "" And MSHFlexGrid3.TextMatrix(I, 13) <> "" And Val(MSHFlexGrid3.TextMatrix(I, 4)) > 0 Then
          MSHFlexGrid3.row = I
          MSHFlexGrid3.TextMatrix(I, 25) = strChecked
       End If
   Next
Else
   For I = 2 To MSHFlexGrid3.Rows - 1
       MSHFlexGrid3.row = I
       MSHFlexGrid3.TextMatrix(I, 25) = strUnChecked
   Next
End If
End Sub

Private Sub Option1_Click()
Call Option2_Click
End Sub

Private Sub Option2_Click()
'CmdExport.Enabled = False
CmbInvoiceType.Text = ""
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbInvoiceType.Enabled = False
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
   Check1.Enabled = False
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = False
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbInvoiceType.Enabled = True
   TxtFromDate.Enabled = True
   TxtToDate.Enabled = True
   TxtFromDate.Value = GetFirstDateOfMonth(Date)
   'TxtFromDate.Value = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")
   TxtToDate.Value = GetLastDateOfMonth(Date)
   Check1.Enabled = True
   Check1.Value = 1
   CmbJournalCode.Text = ""
   CmbJournalCode.Enabled = True
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = True
End If
End Sub
Private Function GetFirstDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_) - 1
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))

If mSMonth > 0 Then
   GetFirstDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & "01")
Else
   GetFirstDateOfMonth = CDate(mSYear - 1 & "-12-" & "01")
End If
End Function

Private Sub TxtFromDate_LostFocus()
If Check1.Value = 1 Then
   TxtFromDate.Value = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")
End If
'TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
End Sub

Private Sub TxtToDate_LostFocus()
If Check1.Value = 1 Then
   TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End If
End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)

mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function
