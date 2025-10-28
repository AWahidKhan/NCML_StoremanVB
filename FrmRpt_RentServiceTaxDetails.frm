VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_RentServiceTaxDetails 
   Caption         =   "Rent Service Tax Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.ComboBox CmbSTaxDesc 
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
         Left            =   11400
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   375
         Width           =   3105
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
         Left            =   3675
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   375
         Width           =   4095
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
         Left            =   7860
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   375
         Width           =   3495
      End
      Begin VB.ComboBox CmbMonth 
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
         ItemData        =   "FrmRpt_RentServiceTaxDetails.frx":0000
         Left            =   120
         List            =   "FrmRpt_RentServiceTaxDetails.frx":0002
         TabIndex        =   1
         Top             =   375
         Width           =   1935
      End
      Begin VB.ComboBox CmbYear 
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
         ItemData        =   "FrmRpt_RentServiceTaxDetails.frx":0004
         Left            =   2130
         List            =   "FrmRpt_RentServiceTaxDetails.frx":0006
         TabIndex        =   2
         Top             =   375
         Width           =   1455
      End
      Begin VB.ComboBox CmbYN 
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
         ItemData        =   "FrmRpt_RentServiceTaxDetails.frx":0008
         Left            =   7515
         List            =   "FrmRpt_RentServiceTaxDetails.frx":0012
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   345
         Left            =   12600
         TabIndex        =   11
         Top             =   1215
         Width           =   1005
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   11595
         TabIndex        =   10
         Top             =   1215
         Width           =   1005
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   345
         Left            =   13605
         TabIndex        =   12
         Top             =   1215
         Width           =   1005
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
         Left            =   3390
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox TxtAgreementID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1725
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtClaimID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   90
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton OptExpenseDate 
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
         Height          =   375
         Left            =   1515
         TabIndex        =   21
         Top             =   555
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptClaimDate 
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
         Height          =   375
         Left            =   195
         TabIndex        =   20
         Top             =   555
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   345
         Left            =   6330
         TabIndex        =   13
         Top             =   1665
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "MMM-YY"
         Format          =   50003969
         CurrentDate     =   40012
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   345
         Left            =   8610
         TabIndex        =   14
         Top             =   1665
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50003969
         CurrentDate     =   40012
      End
      Begin VB.Label Label9 
         Caption         =   "Tax Desc"
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
         Left            =   12457
         TabIndex        =   32
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label7 
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
         Left            =   5115
         TabIndex        =   31
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Left            =   9180
         TabIndex        =   30
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Month"
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
         Left            =   780
         TabIndex        =   29
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Year"
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
         Left            =   2610
         TabIndex        =   28
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Rent Service Tax Transaction generated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7650
         TabIndex        =   26
         Top             =   705
         Width           =   2715
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
         Left            =   10935
         TabIndex        =   25
         Top             =   855
         Width           =   2415
      End
      Begin VB.Label LblGodown 
         AutoSize        =   -1  'True
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
         Left            =   4830
         TabIndex        =   24
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Agreement ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1905
         TabIndex        =   23
         Top             =   930
         Width           =   1260
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Left            =   480
         TabIndex        =   22
         Top             =   900
         Width           =   900
      End
      Begin VB.Label LblWHRToDate 
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
         Height          =   240
         Left            =   7770
         TabIndex        =   18
         Top             =   1695
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblWHRDateFrom 
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
         Left            =   5280
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
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
         TabIndex        =   16
         Top             =   240
         Width           =   75
      End
      Begin VB.Label LblRecordCount 
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
         Left            =   13335
         TabIndex        =   15
         Top             =   855
         Width           =   135
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7545
      Left            =   120
      TabIndex        =   19
      Top             =   1875
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   13309
      _Version        =   393216
      WordWrap        =   -1  'True
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   9480
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmRpt_RentServiceTaxDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLessorID, mFromDate, mToDate, mLocationID, mStateID As Variant

Public Sub Grid_Head()

With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 24 '20 '18
     .WordWrap = True
     
     .TextMatrix(0, 0) = "Rent Claim ID"
     .TextMatrix(0, 1) = "Rent Claim Date"
     .TextMatrix(0, 2) = "Expense From"
     .TextMatrix(0, 3) = "State"
     .TextMatrix(0, 4) = "Location"
     .TextMatrix(0, 5) = "AgreementID"
     .TextMatrix(0, 6) = "Agreement Start Date"
     .TextMatrix(0, 7) = "Agreement End Date"
     .TextMatrix(0, 8) = "Lessor Name"
     .TextMatrix(0, 9) = "Rent Amount"
     .TextMatrix(0, 10) = "Status"
     .TextMatrix(0, 11) = "BR Agreement Flag"
     .TextMatrix(0, 12) = "Tax Code"
     .TextMatrix(0, 13) = "Key"
     .TextMatrix(0, 14) = "Godown Rent Service Tax ID"
     .TextMatrix(0, 15) = "Godown Rent Service Tax Date"
     .TextMatrix(0, 16) = "Bill No"
     .TextMatrix(0, 17) = "Bill Date"
     .TextMatrix(0, 18) = "Bill Amount"
     .TextMatrix(0, 19) = "Service Tax Rate"
     .TextMatrix(0, 20) = "Service Tax Amount"
     .TextMatrix(0, 21) = "Flag"
     .TextMatrix(0, 22) = "Batch ID"
     .TextMatrix(0, 23) = "Remarks"
     
     .ColWidth(0) = 1500
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     .ColWidth(3) = 1500
     .ColWidth(4) = 1500
     .ColWidth(5) = 1500
     .ColWidth(6) = 2500
     .ColWidth(7) = 1500
     .ColWidth(8) = 1500
     .ColWidth(9) = 1500
     .ColWidth(10) = 1500
     .ColWidth(11) = 1500
     .ColWidth(12) = 1500
     .ColWidth(13) = 1500
     .ColWidth(14) = 1500
     .ColWidth(15) = 1500
     .ColWidth(16) = 1500
     .ColWidth(17) = 1500
     .ColWidth(18) = 1500
     .ColWidth(19) = 1500
     
     .RowHeight(0) = 800
End With
End Sub
Private Sub CmbLessor_GotFocus()
tmp = CmbLessor.Text
Call TableMst_Lessor(" Order by LessorName")

CmbLessor.Clear
If RsMst_Lessor.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Lessor.RecordCount
    CmbLessor.AddItem RsMst_Lessor!LessorName
    RsMst_Lessor.MoveNext
Next
CmbLessor.Text = tmp
End Sub

Private Sub CmbLessor_LostFocus()
If CmbLessor.Text = "" Then
   Exit Sub
End If
RsMst_Lessor.MoveFirst
RsMst_Lessor.Find "LessorName = '" & CmbLessor.Text & "'"
If RsMst_Lessor.EOF Then
   MsgBox "Invalid selection "
   CmbLessor.SetFocus
   Exit Sub
End If
mLessorID = RsMst_Lessor!LessorID
End Sub

Private Sub CmbSTaxDesc_GotFocus()
tmp = CmbSTaxDesc.Text
Call TableMst_ServiceTax
CmbSTaxDesc.Clear
If RsMst_ServiceTax.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ServiceTax.RecordCount
    CmbSTaxDesc.AddItem RsMst_ServiceTax!Description
    RsMst_ServiceTax.MoveNext
Next
CmbSTaxDesc.Text = tmp
End Sub

Private Sub CmbSTaxDesc_LostFocus()
If CmbSTaxDesc.Text = "" Then
   mServiceTaxID = Null
   TxtServiceTaxNAVCode = ""
   Exit Sub
End If
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "Description = '" & CmbSTaxDesc.Text & "'"
If RsMst_ServiceTax.EOF Then
   MsgBox "Invalid selection "
   CmbSTaxDesc.SetFocus
   Exit Sub
End If
mServiceTaxID = RsMst_ServiceTax!ServiceTaxID
End Sub

Private Sub CmbYN_LostFocus()
If CmbYN.Text = "" Then Exit Sub

If LCase(CmbYN.Text) <> "yes" And LCase(CmbYN.Text) <> "no" Then
   MsgBox "Invalid Selection !!!"
   CmbYN.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
OptClaimDate.Value = True
TxtFromDate = Date - 30
TxtToDate = Date
CmbMonth.Text = ""
CmbYear.Text = ""
TxtClaimID = ""
TxtAgreementID = ""
CmbStateID.Text = ""
CmbLocation.Text = ""
CmbSTaxDesc.Text = ""
CmbLessor = ""
CmbYN = ""
Call Grid_Head
CmdExcel.Enabled = False
ProgressBar1.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_RentServiceTaxDetails.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_RentServiceTaxDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

Dim wc As Variant
wc = ""
Call Grid_Head
MsgBox "Wait till next message !!!"
CmdExcel.Enabled = False
ProgressBar1.Value = 0


If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

mMonth = CmbMonth.ListIndex + 1
vMonth = Month(Now)


If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

'If OptClaimDate.Value = True Then
'   If Gen_DBType = "MDB" Then
'      wc = " And (b.ClaimDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
'   Else
'      wc = " And (b.ClaimDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
'   End If
'Else
'   If Gen_DBType = "MDB" Then
'      wc = " And (b.ExpenseFrom Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
'   Else
'      wc = " And (b.ExpenseFrom Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
'   End If
'End If

If TxtClaimID.Text <> "" Then
   wc = wc & " And TGR.ClaimID = " & Val(TxtClaimID) & " "
End If

If TxtAgreementID <> "" Then
   wc = wc & " And TL.LAID = " & Val(TxtAgreementID) & " "
End If

If CmbLessor.Text <> "" Then
   wc = wc & " And ML.LessorID=" & Val(mLessorID) & ""
End If

If CmbStateID.Text <> "" Then
   wc = wc & " And MState.StateName='" & CmbStateID.Text & "'"
Else
   If Gen_WcState <> "" Then
      wc = wc & " And MState." & Gen_WcState & ""
   End If
End If

If CmbLocation <> "" Then
   wc = wc & " And MLoc.LocationName='" & CmbLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And MLoc." & Gen_WcLocation & ""
   End If
End If

If CmbSTaxDesc.Text <> "" Then
   wc = wc & " And MSTax.Description='" & CmbSTaxDesc.Text & "'"
End If


'tmp = " Select Distinct b.ClaimID 'RentClaimID',b.ClaimDate 'RentCLaimDate',b.ExpenseFrom,TLA.LAID 'AgreementID',"
'tmp = tmp & " TLA.StartDate 'AgreeStartDate',TLA.ExpiryDate 'AgreeEndDate',MLes.LessorName,b.Amount 'RentAmount',"
'tmp = tmp & " d.GodownRentSTID , d.GodownRentSTDate, d.BillNo, d.BillDate, d.BillAmount, d.STRate, d.STAmount, d.Flag, d.BatchID, d.Remarks"
'tmp = tmp & " From Txn_GodownRentDetail a"
'tmp = tmp & " inner Join Txn_GodownRent b on a.ClaimID = b.ClaimID"
'tmp = tmp & " left Join Txn_GodownRentSTClaimDetail c on a.claimid = c.claimid"
'tmp = tmp & " left join Txn_GodownRentST d on c.GodownRentSTID = d.GodownRentSTID"
'tmp = tmp & " inner Join Mst_Lessor MLes on b.LessorID = MLes.LessorID"
'tmp = tmp & " inner join Txn_LessorAgreement TLA On a.LAID = TLA.LAID"
'tmp = tmp & " Where TLA.ServiceTaxID in (3,5) and b.Flag in ('S','A') And b.BatchID > 0"
'tmp = tmp & wc
'tmp = tmp & " Order By b.ClaimID,b.ExpenseFrom"

tmp = "Select ClaiMID,ClaimDate,ExpenseFrom,StateName,LocationName, LAID,StartDate,ExpiryDate,LessorName,Amount,[Status],BRAgreement,Description,[Key],GodownRentSTID , GodownRentSTDate, BillNo, BillDate, BillAmount, STRate, STAmount, Flag, BatchID, Remarks"
tmp = tmp & " From ("
tmp = tmp & " Select TGR.ClaiMID,TGR.ClaimDate,TGR.ExpenseFrom,MState.StateName,MLoc.LocationName, TL.LAID,TL.StartDate,TL.ExpiryDate,ML.LessorName,sum(TG.Amount)Amount,'Not Processed' 'Status',TL.BRAgreement,MSTax.Description,Convert(Varchar,TL.LAID) +'-'+ Convert(Varchar,ML.LessorID) 'Key',0 'GodownRentSTID' , '' 'GodownRentSTDate', '' 'BillNo', '' 'BillDate', 0 'BillAmount', 0 'STRate', 0 'STAmount', '' 'Flag', 0 BatchID, '' 'Remarks'"
tmp = tmp & " From Txn_GodownRent TGR"
tmp = tmp & " Inner join Txn_GodownRentDetail TG on TG.ClaimID=TGR.ClaimID"
tmp = tmp & " Inner join Txn_LessorAgreement TL on TG.LAID=TL.LAID"
tmp = tmp & " Inner join Mst_Godown MG on TG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID"
tmp = tmp & " Inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Location MLoc on MC.LocationID=MLoc.LocationID"
tmp = tmp & " Inner join Mst_State MState on MLoc.StateID=MState.StateID"
tmp = tmp & " Inner join Mst_ServiceTax MSTax on TL.ServiceTaxID=MSTax.ServiceTaxID"
tmp = tmp & " Where TGR.Flag in ('S','A') And TGR.BatchID > 0 and TGR.ExpenseFrom='" & mFromDate & "'"
tmp = tmp & " And Convert(Varchar,TL.LAID) +'-'+ Convert(Varchar,ML.LessorID) Not In (Select Convert(Varchar,TGRSD.AgreementID) +'-'+ Convert(Varchar,TGRS.LessorID)  From Txn_GodownRentST TGRS Inner Join Txn_GodownRentSTAgreeDetail TGRSD On TGRS.GodownRentSTID=TGRSD.GodownRentSTID Where TGRS.Flag not in ('R') And TGRS.MonthYear='" & mFromDate & "')"
tmp = tmp & wc
tmp = tmp & " Group by  TGR.ClaimID,TGR.ClaimDate,TGR.ExpenseFrom,MState.StateName,MLoc.LocationName, TL.LAID,TL.StartDate,TL.ExpiryDate,ML.LessorName,TL.EntryDate,ML.LessorID,TL.BRAgreement,MSTax.Description"
tmp = tmp & " Union All"
tmp = tmp & " Select TGR.ClaiMID,TGR.ClaimDate,TGR.ExpenseFrom,MState.StateName,MLoc.LocationName, TL.LAID,TL.StartDate,TL.ExpiryDate,ML.LessorName,sum(TG.Amount)Amount,'Processed' 'Status',TL.BRAgreement,MSTax.Description,Convert(Varchar,TL.LAID) +'-'+ Convert(Varchar,ML.LessorID) 'Key',0 'GodownRentSTID' , '' 'GodownRentSTDate', '' 'BillNo', '' 'BillDate', 0 'BillAmount', 0 'STRate', 0 'STAmount', '' 'Flag', 0 BatchID, '' 'Remarks'"
tmp = tmp & " From Txn_GodownRent TGR"
tmp = tmp & " Inner join Txn_GodownRentDetail TG on TG.ClaimID=TGR.ClaimID"
tmp = tmp & " Inner join Txn_LessorAgreement TL on TG.LAID=TL.LAID"
tmp = tmp & " Inner join Mst_Godown MG on TG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID"
tmp = tmp & " Inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Location MLoc on MC.LocationID=MLoc.LocationID"
tmp = tmp & " Inner join Mst_State MState on MLoc.StateID=MState.StateID"
tmp = tmp & " Inner join Mst_ServiceTax MSTax on TL.ServiceTaxID=MSTax.ServiceTaxID"
tmp = tmp & " Where TGR.Flag in ('S','A') And TGR.BatchID > 0 and TGR.ExpenseFrom='" & mFromDate & "'"
tmp = tmp & " And Convert(Varchar,TL.LAID) +'-'+ Convert(Varchar,ML.LessorID) In (Select Convert(Varchar,TGRSD.AgreementID) +'-'+ Convert(Varchar,TGRS.LessorID)  From Txn_GodownRentST TGRS Inner Join Txn_GodownRentSTAgreeDetail TGRSD On TGRS.GodownRentSTID=TGRSD.GodownRentSTID Where TGRS.Flag not in ('R') And TGRS.MonthYear='" & mFromDate & "')"
tmp = tmp & wc
tmp = tmp & " Group by  TGR.ClaiMID,TGR.ClaimDate,TGR.ExpenseFrom,MState.StateName,MLoc.LocationName, TL.LAID,TL.StartDate,TL.ExpiryDate,ML.LessorName,ML.LessorID,TL.BRAgreement,MSTax.Description"
tmp = tmp & " ) A"
If CmbYN.Text <> "" Then
   If LCase(CmbYN.Text) = "yes" Then
     tmp = tmp & " Where [Status] = 'Processed'"
   ElseIf LCase(CmbYN.Text) = "no" Then
      tmp = tmp & " Where [Status] = 'Not Processed'"
   End If
End If
tmp = tmp & " Order by ClaiMID"

Call Tmp1Table


tmp = " Select Distinct b.GodownRentSTID , b.GodownRentSTDate, b.BillNo, b.BillDate, b.BillAmount, b.STRate,b.STAmount, b.Flag, b.BatchID, b.Remarks,Convert(Varchar,a.AgreementID) +'-'+ Convert(Varchar,b.LessorID) 'Key'"
tmp = tmp & " From Txn_GodownRentSTAgreeDetail a"
tmp = tmp & " Inner Join Txn_GodownRentST b On a.GodownRentSTID = b.GodownRentSTID"
tmp = tmp & " Where b.Flag not in ('R') And b.MonthYear='" & mFromDate & "' " 'And Convert(Varchar,a.AgreementID) +'-'+ Convert(Varchar,b.LessorID) = '3580-7'"
If CmbLessor.Text <> "" Then
   tmp = tmp & " And b.LessorID=" & Val(mLessorID) & ""
End If

Call Tmp2Table

If TmpRs1.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs1.RecordCount + 2
   CmdExcel.Enabled = True
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   TmpRs1.MoveFirst
   For i = 1 To TmpRs1.RecordCount
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs1.Fields(c)), "", TmpRs1.Fields(c))
           If c > 13 Then
              If TmpRs2.RecordCount > 0 Then
                 TmpRs2.MoveFirst
                 TmpRs2.Find "Key = '" & TmpRs1!Key & "'"
                 If TmpRs2.EOF = False Then
                    MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs2.Fields(c - 14)), "", TmpRs2.Fields(c - 14))
                 End If
              End If
           End If
           If c = 1 Or c = 2 Or c = 6 Or c = 7 Or c = 15 Or c = 17 Then
              MSHFlexGrid1.TextMatrix(i, c) = Format(MSHFlexGrid1.TextMatrix(i, c), "dd-MMM-yyyy")
           End If
       Next
       If MSHFlexGrid1.TextMatrix(i, 21) <> "" Then
          MSHFlexGrid1.TextMatrix(i, 21) = IIf(MSHFlexGrid1.TextMatrix(i, 21) = "A", "Approved", "Pending")
       End If
       TmpRs1.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
Else
   MsgBox "No Record Found !!!"
   Exit Sub
End If
LblRecordCount.Caption = TmpRs1.RecordCount
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
For i = 1 To 12
    CmbMonth.AddItem MonthName(i), i - 1
Next
Call CmdClear_Click
End Sub

Private Sub CmbMonth_LostFocus()

If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For i = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(i) Then
       CmbMonth.Text = CmbMonth.List(i)
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For i = 2015 To Year(Now)
    CmbYear.AddItem i
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For i = 2015 To Year(Now)
    If CmbYear.Text = i Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

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
For i = 1 To RsMst_State.RecordCount
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

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where LocationID <> 0 And StateID= " & mStateID & " And " & Gen_WcLocation & " ")
Else
   Call TableMst_Location(" Where LocationID <> 0 And StateID= " & mStateID & "")
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
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

