VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_GodownRentClaim 
   Caption         =   "Processed Godown Rent Status Report"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2145
      Left            =   120
      TabIndex        =   14
      Top             =   45
      Width           =   15045
      Begin VB.TextBox TxtSClaimID 
         Appearance      =   0  'Flat
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
         Left            =   105
         TabIndex        =   7
         Top             =   1432
         Width           =   1800
      End
      Begin VB.ComboBox CmbSPaymentStatus 
         Height          =   315
         ItemData        =   "FrmRpt_GodownRentClaim.frx":0000
         Left            =   11235
         List            =   "FrmRpt_GodownRentClaim.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   743
         Width           =   3555
      End
      Begin VB.ComboBox CmbSFlag 
         Height          =   315
         ItemData        =   "FrmRpt_GodownRentClaim.frx":001D
         Left            =   7645
         List            =   "FrmRpt_GodownRentClaim.frx":002D
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1455
         Width           =   3480
      End
      Begin VB.ComboBox CmbSBatchID 
         Height          =   315
         Left            =   5510
         TabIndex        =   9
         Top             =   1455
         Width           =   2025
      End
      Begin VB.ComboBox CmbSLessor 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1455
         Width           =   3375
      End
      Begin VB.ComboBox CmbSLocationID 
         Height          =   315
         Left            =   7645
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   743
         Width           =   3480
      End
      Begin VB.ComboBox CmbStateID 
         Height          =   315
         Left            =   3935
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   743
         Width           =   3600
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
         Left            =   1620
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   135
         Width           =   1575
      End
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
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   120
         Value           =   -1  'True
         Width           =   1575
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
         Height          =   360
         Left            =   12450
         TabIndex        =   12
         Top             =   1432
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Ger Report"
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
         Left            =   11235
         TabIndex        =   11
         Top             =   1432
         Width           =   1215
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
         Height          =   360
         Left            =   13665
         TabIndex        =   13
         Top             =   1432
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   735
         Width           =   1800
         _ExtentX        =   3175
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
         Format          =   112328705
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   2040
         TabIndex        =   3
         Top             =   735
         Width           =   1785
         _ExtentX        =   3149
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
         Format          =   112328705
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label LblTo 
         Caption         =   "Claim Entered Upto"
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
         TabIndex        =   26
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label Label7 
         Caption         =   "Claim ID"
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
         Left            =   615
         TabIndex        =   25
         Top             =   1155
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "Payment Status"
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
         Left            =   12307
         TabIndex        =   24
         Top             =   465
         Width           =   1410
      End
      Begin VB.Label Label5 
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
         Left            =   8808
         TabIndex        =   23
         Top             =   1155
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Navision Batch ID"
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
         Left            =   5715
         TabIndex        =   22
         Top             =   1155
         Width           =   1620
      End
      Begin VB.Label Label2 
         Caption         =   "Lessor Name"
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
         Left            =   3120
         TabIndex        =   21
         Top             =   1155
         Width           =   1215
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
         Height          =   210
         Left            =   8988
         TabIndex        =   20
         Top             =   480
         Width           =   795
      End
      Begin VB.Label LblStateID 
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
         Left            =   5203
         TabIndex        =   19
         Top             =   465
         Width           =   1065
      End
      Begin VB.Label LblFrom 
         Caption         =   "Claim Entered From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   465
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record Found :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11220
         TabIndex        =   16
         Top             =   1830
         Width           =   2220
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   13545
         TabIndex        =   15
         Top             =   1845
         Width           =   90
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7395
      Left            =   120
      TabIndex        =   17
      Top             =   2250
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13044
      _Version        =   393216
      MergeCells      =   1
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
Attribute VB_Name = "FrmRpt_GodownRentClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLessorID, mSStatus  As Variant


Private Sub CmbSFlag_LostFocus()
If CmbSFlag.Text = "" Then Exit Sub

If LCase(CmbSFlag.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSFlag.Text) = "pending" Then
   mSStatus = "P"
ElseIf LCase(CmbSFlag.Text) = "rejected" Then
   mSStatus = "R"
ElseIf LCase(CmbSFlag.Text) = "settled" Then
   mSStatus = "S"
Else
   MsgBox "Invalid selection!!!"
   CmbSFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSLessor_GotFocus()
tmp = CmbSLessor.Text

Call TableMst_Lessor

CmbSLessor.Clear
If RsMst_Lessor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Lessor.RecordCount
    CmbSLessor.AddItem RsMst_Lessor!LessorName
    RsMst_Lessor.MoveNext
Next
CmbSLessor.Text = tmp
End Sub

Private Sub CmbSLessor_LostFocus()
If CmbSLessor.Text = "" Then
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorName = '" & CmbSLessor.Text & "'"
If RsMst_Lessor.EOF Then
   MsgBox "Invalid selection "
   CmbSLessor.SetFocus
   Exit Sub
End If
mLessorID = RsMst_Lessor!LessorID
End Sub


Private Sub CmbSPaymentStatus_LostFocus()
If CmbSPaymentStatus.Text = "" Then Exit Sub

If LCase(CmbSPaymentStatus.Text) <> "paid" And LCase(CmbSPaymentStatus.Text) <> "un-paid" Then
   MsgBox "Invalid selection!!!"
   CmbSPaymentStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetReport_Click()
Dim mClaimID, mClaimIDNew   As Variant

LblTotalRecords.Caption = ""
Call GridHead


If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "'From Date' must be less than 'To Date'"
   TxtFromDate.SetFocus
   Exit Sub
End If
Screen.MousePointer = vbHourglass

tmp = "select ms.StateName, l.LocationName, a.ClaimID, a.ClaimDate 'ClaimDate', a.ExpenseFrom 'From' ,"
tmp = tmp & " a.ExpenseTo 'To' , c.LessorName, a.Remarks, a.ApprovalRemarks,a.ApprovalDate,"
tmp = tmp & " a.BatchID, e.EmployeeName 'ForwardTo', a.Flag, a.UpdatedByFAG, a.UpdatedDateFAG, d.GodownNo,d.Address, b.Amount,a.BankName,a.ChequeNo, a.Amount 'TotalAmt',a.TDS, a.PaidAmount,a.PaidDate,d.Area,mc.CMPName,mc.CMPID "
tmp = tmp & " from Txn_GodownRent a"
tmp = tmp & " inner join Txn_GodownRentDetail b on a.ClaimID = b.ClaimID"
tmp = tmp & " inner join mst_lessor c on a.lessorid = c.lessorid"
tmp = tmp & " inner join mst_godown d on b.godownid = d.godownid"
tmp = tmp & " inner join Mst_Employee e on a.ForwardToID = e.EmployeeID"
tmp = tmp & " inner Join Mst_CMP mc on d.CMPID = mc.cmpid"
tmp = tmp & " inner join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner join Mst_State ms on l.StateID = ms.StateID"

If Option1.Value = True Then
   If Gen_DBType = "MDB" Then
      tmp = tmp & " Where (a.ClaimDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " Where (a.ClaimDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
   End If
Else
   If Gen_DBType = "MDB" Then
      tmp = tmp & " Where (a.ExpenseFrom Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
   Else
      tmp = tmp & " Where (a.ExpenseFrom Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
   End If
End If

If Trim(TxtSClaimID.Text) = "" Then
   If CmbStateID.Text <> "" Then
      tmp = tmp & " And ms.StateName = '" & CmbStateID.Text & "' "
   Else
      If Gen_WcState <> "" Then
         tmp = tmp & " And ms." & Gen_WcState
      End If
   End If
   
   If CmbSLocationID.Text <> "" Then
      tmp = tmp & " And l.LocationName = '" & CmbSLocationID.Text & "' "
   Else
      If Gen_WcLocation <> "" Then
         tmp = tmp & " And l." & Gen_WcLocation
      End If
   End If
   
   If Trim(LCase(CmbSPaymentStatus.Text)) = "paid" Then
      tmp = tmp & " AND a.BankName <> ''"
   ElseIf Trim(LCase(CmbSPaymentStatus.Text)) = "un-paid" Then
      tmp = tmp & " AND (a.BankName = '' OR a.BankName IS NULL)"
   End If

   If CmbSLessor.Text <> "" Then
      tmp = tmp & " And a.lessorid = " & mLessorID & ""
   End If
   
   If CmbSBatchID.Text <> "" Then
      tmp = tmp & " And a.BatchID = " & CmbSBatchID.Text & ""
   End If
   
   If CmbSFlag.Text <> "" Then
      tmp = tmp & " And a.Flag = '" & mSStatus & "'"
   End If
Else
   CmbStateID.Text = ""
   CmbSLocationID.Text = ""
   CmbSPaymentStatus.Text = ""
   CmbSLessor.Text = ""
   CmbSBatchID.Text = ""
   CmbSFlag.Text = ""
   tmp = tmp & " And a.ClaimID=" & TxtSClaimID.Text
End If

tmp = tmp & " Order By ms.StateName,l.LocationName,a.ClaimID"

Call TmpTable
mClaimID = 0
mClaimIDNew = 0

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       DoEvents
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           If C = 2 Then
              mClaimIDNew = MSHFlexGrid1.TextMatrix(I, C)
           End If
       Next
       
       If MSHFlexGrid1.TextMatrix(I, 12) = "P" Then
          MSHFlexGrid1.TextMatrix(I, 12) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(I, 12) = "R" Then
          MSHFlexGrid1.TextMatrix(I, 12) = "Rejected"
       ElseIf MSHFlexGrid1.TextMatrix(I, 12) = "A" Then
          MSHFlexGrid1.TextMatrix(I, 12) = "Approved"
       ElseIf MSHFlexGrid1.TextMatrix(I, 12) = "S" Then
          MSHFlexGrid1.TextMatrix(I, 12) = "Settled"
       End If
       
       MSHFlexGrid1.TextMatrix(I, 3) = Format(TmpRs!ClaimDate, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs!From, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 5) = Format(TmpRs!To, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 9) = Format(TmpRs!ApprovalDate, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 14) = Format(TmpRs!UpdatedDateFAG, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 23) = Format(TmpRs!PaidDate, "dd-MMM-yyyy")
       
       MSHFlexGrid1.TextMatrix(I, 20) = Format(TmpRs!TotalAmt, "#####0.00")
       MSHFlexGrid1.TextMatrix(I, 17) = Format(TmpRs!Amount, "#####0.00")
       MSHFlexGrid1.TextMatrix(I, 22) = Format(TmpRs!PaidAmount, "#####0.00")
       
       'MSHFlexGrid1.TextMatrix(i, 24) = 'Format(TmpRs!PaidDate, "dd-MMM-yyyy")
       
       If mClaimID = mClaimIDNew Then
          
          
          
          MSHFlexGrid1.TextMatrix(I, 3) = ""
          MSHFlexGrid1.TextMatrix(I, 4) = ""
          MSHFlexGrid1.TextMatrix(I, 5) = ""
          MSHFlexGrid1.TextMatrix(I, 6) = ""
          MSHFlexGrid1.TextMatrix(I, 7) = ""
          MSHFlexGrid1.TextMatrix(I, 8) = ""
          MSHFlexGrid1.TextMatrix(I, 9) = ""
          MSHFlexGrid1.TextMatrix(I, 10) = ""
          MSHFlexGrid1.TextMatrix(I, 11) = ""
          MSHFlexGrid1.TextMatrix(I, 12) = ""
          MSHFlexGrid1.TextMatrix(I, 13) = ""
          MSHFlexGrid1.TextMatrix(I, 14) = ""
          MSHFlexGrid1.TextMatrix(I, 18) = ""
          MSHFlexGrid1.TextMatrix(I, 20) = ""
          MSHFlexGrid1.TextMatrix(I, 22) = ""
          MSHFlexGrid1.TextMatrix(I, 23) = ""
       End If
       mClaimID = mClaimIDNew
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
   LblTotalRecords.Caption = MSHFlexGrid1.Rows - 2
   CmdExcel.Enabled = True
Else
   MsgBox "No record Found."
End If
Screen.MousePointer = vbNormal
End Sub



Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
Call GridHead
Frame1.Enabled = True
CmdExcel.Enabled = False
TxtFromDate.Value = Date - 365
TxtToDate.Value = Date
CmbStateID.Text = ""
CmbSLocationID.Text = ""

tmp = "Select Distinct BatchID from Txn_GodownRent Where BatchID > 0  order by BatchID"
Call Tmp4Table
CmbSBatchID.Clear
If TmpRs4.RecordCount > 0 Then
   For I = 1 To TmpRs4.RecordCount
       CmbSBatchID.AddItem TmpRs4.Fields(0)
       TmpRs4.MoveNext
   Next
End If

End Sub
Private Function GridHead()
With MSHFlexGrid1
    .Clear
    .Cols = 27 ' 25
    .Rows = 2
    .WordWrap = True
    
    
    .TextMatrix(0, 0) = "State"
    .TextMatrix(0, 1) = "Location"
    .TextMatrix(0, 2) = "Claim ID"
    .TextMatrix(0, 3) = "Claim Date"
    .TextMatrix(0, 4) = "Rent From"
    .TextMatrix(0, 5) = "Rent To"
'    .TextMatrix(0, 4) = "Lessor ID"
    .TextMatrix(0, 6) = "Lessor Name"
    .TextMatrix(0, 7) = "Remark : CMPM"
    .TextMatrix(0, 8) = "Remark : Approval"
    .TextMatrix(0, 9) = "Approval Date"
    .TextMatrix(0, 10) = "Navision Batch ID"
'    .TextMatrix(0, 11) = "ForwardTo ID"
    .TextMatrix(0, 11) = "Handling Officer"
    .TextMatrix(0, 12) = "Claim Status"
    .TextMatrix(0, 13) = "UpdatedBy FAG"
    .TextMatrix(0, 14) = "FAG Posted Date"
'    .TextMatrix(0, 16) = "Godown ID"
    .TextMatrix(0, 15) = "Godown No"
    .TextMatrix(0, 16) = "Address"
    .TextMatrix(0, 17) = "Rent for Godown"
'    .TextMatrix(0, 20) = "G_UID"
    .TextMatrix(0, 18) = "Bank Name"
    .TextMatrix(0, 19) = "Cheque No"
    .TextMatrix(0, 20) = "Total Rent for Lessor"
    .TextMatrix(0, 21) = "TDS "
    .TextMatrix(0, 22) = "Cheque/RTGS Amount"
    .TextMatrix(0, 23) = "Cheque/RTGS Date"
    .TextMatrix(0, 24) = "Area in Sq. Ft."
    .TextMatrix(0, 25) = "CMP Name"
    .TextMatrix(0, 26) = "CMP ID"

    .ColWidth(0) = 1200
    .ColWidth(1) = 1400
    .ColWidth(2) = 1000
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 2400
    .ColWidth(7) = 2600
    .ColWidth(8) = 2600
    .ColWidth(9) = 1400
    .ColWidth(10) = 1000
    .ColWidth(11) = 1400
    .ColWidth(12) = 1000
    .ColWidth(13) = 0
    .ColWidth(14) = 1400
    .ColWidth(15) = 1000
    .ColWidth(16) = 3000
    .ColWidth(17) = 1400
    .ColWidth(18) = 2600
    .ColWidth(19) = 1000
    .ColWidth(20) = 1400
    .ColWidth(21) = 1400
    .ColWidth(22) = 1400
    .ColWidth(23) = 1500

    
    .RowHeight(0) = 600
End With
End Function
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownRentClaim.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownRentClaim.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text

If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
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
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
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

Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbSLocationID.Clear
   CmbStateID.SetFocus
   Exit Sub
End If

If Gen_WcLocation = "" Then
   If CmbStateID.Text <> "" Then
      Call TableMst_Location(" Where StateID = " & mStateID & " ")
'   Else
'      Call TableMst_Location
   End If
Else
   If CmbStateID.Text <> "" Then
      Call TableMst_Location(" Where StateID = " & mStateID & " And " & Gen_WcLocation)
'   Else
'      Call TableMst_Location(" Where " & Gen_WcLocation)
   End If
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbSLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
    
End Sub

'----
Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID

End Sub

Private Sub Option1_Click()
    LblFrom.Caption = "Claim Entered From"
    LblTo.Caption = "Claim Entered Upto"
End Sub

Private Sub Option2_Click()
    LblFrom.Caption = "Expense From"
    LblTo.Caption = "Expense To"
End Sub
