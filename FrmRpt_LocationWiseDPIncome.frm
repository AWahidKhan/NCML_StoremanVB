VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_LocationWiseDPIncome 
   Caption         =   "Location Wise DP Income"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   9435
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   9735
      Left            =   120
      TabIndex        =   0
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
         Height          =   360
         Left            =   12915
         TabIndex        =   5
         Top             =   540
         Width           =   975
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
         Height          =   360
         Left            =   11700
         TabIndex        =   4
         Top             =   540
         Width           =   1215
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
         ItemData        =   "FrmRpt_LocationWiseDPIncome.frx":0000
         Left            =   2055
         List            =   "FrmRpt_LocationWiseDPIncome.frx":0002
         TabIndex        =   2
         Top             =   540
         Width           =   1455
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
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   1890
      End
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   3540
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   3660
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
         Height          =   360
         Left            =   13890
         TabIndex        =   6
         Top             =   540
         Width           =   975
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   8175
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   14420
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Location Wise DP Income"
         TabPicture(0)   =   "FrmRpt_LocationWiseDPIncome.frx":0004
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "DP Income as per Store Man"
         TabPicture(1)   =   "FrmRpt_LocationWiseDPIncome.frx":0020
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid2"
         Tab(1).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7575
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13361
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   7335
            Left            =   -74880
            TabIndex        =   15
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
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
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
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
         Left            =   2535
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Month"
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
         Left            =   735
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LblLocationID 
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
         Height          =   240
         Left            =   4950
         TabIndex        =   10
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   8100
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label LblTotalRecord 
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
         Left            =   10500
         TabIndex        =   8
         Top             =   600
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmRpt_LocationWiseDPIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mYear, mTotalDays, mLocationID, mSM_Prefix As Variant

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("Where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If

For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I)
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
   If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
      CmbMonth.Text = CmbMonth.List(I)
      mStatus = True
      Exit For
   End If
Next
If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If
mMonth = CmbMonth.ListIndex + 1

End Sub

Private Sub CmbYear_GotFocus()
If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If
tmp = CmbYear.Text
CmbYear.Clear
For I = 2012 To Year(Now)
    CmbYear.AddItem I
Next

If Month(Now) = 12 Then
   CmbYear.AddItem Year(Now) + 1
End If

CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
Dim myr As Variant

flg = "N"

If Month(Now) = 12 Then
   myr = Year(Now) + 1
Else
   myr = Year(Now)
End If


For I = 2010 To myr 'Year(Now)
    If CmbYear.Text = I Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

mYear = CmbYear.Text
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

End Sub

Private Sub CmdClear_Click()
CmbMonth.Text = ""
CmbYear.Text = ""
CmbLocationID.Text = ""
CmdExcel.Enabled = False
LblTotalRecord.Caption = ""
Call Grid_Head
Call Grid_Head1
End Sub

Private Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 3
    .Cols = 19
    .FixedRows = 2
    .WordWrap = True
    
    .TextMatrix(1, 0) = "Location Name"
    .TextMatrix(1, 1) = "CMP Name"
    .TextMatrix(1, 2) = "GSLID"
    .TextMatrix(1, 3) = "Godown No"
    .TextMatrix(1, 4) = "Stack No"
    .TextMatrix(1, 5) = "Lot No"
    .TextMatrix(1, 6) = "Commodity"
    .TextMatrix(1, 7) = "Client Name"
    .TextMatrix(1, 8) = "Exchange Type"
    .TextMatrix(1, 9) = "Balance Bags"
    .TextMatrix(1, 10) = "Balance Weight"
    .TextMatrix(1, 11) = "Bag Size"
    .TextMatrix(1, 12) = "Billing Cycle"
    .TextMatrix(1, 13) = "Billing Unit"
    .TextMatrix(1, 14) = "Invoice Amt"
    .TextMatrix(1, 15) = "Rate"
    .TextMatrix(1, 16) = "No Of Days"
    .TextMatrix(1, 17) = "TFlag"
    .TextMatrix(1, 18) = "Invoice No"
    
    
    .ColWidth(0) = 1700
    .ColWidth(1) = 1800
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1800
    .ColWidth(7) = 2500
    .ColWidth(8) = 1200
    .ColWidth(9) = 1500
    .ColWidth(10) = 1500
    .ColWidth(11) = 1200
    .ColWidth(12) = 1500
    .ColWidth(13) = 1200
    .ColWidth(14) = 1500
    .ColWidth(15) = 1200
    .ColWidth(16) = 1500
    .ColWidth(17) = 1500
    .ColWidth(18) = 1500
    
    .RowHeight(1) = 700
End With
End Sub


Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .Cols = 4
    .FixedRows = 2
    
    .TextMatrix(1, 0) = "LocationName"
    .TextMatrix(1, 1) = "DP Amount as per Store Man"
    .TextMatrix(1, 2) = "DP Amount as per NCDEX"
    .TextMatrix(1, 3) = "Difference"
    
    .ColWidth(0) = 2500
    .ColWidth(1) = 2800
    .ColWidth(2) = 2500
    .ColWidth(3) = 2000
    
End With
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_LocationWiseDPIncome.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Variant
Dim mCol As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_LocationWiseDPIncome.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Location Wise DP Income"
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 25
oWS.Range("b:b").ColumnWidth = 17
oWS.Range("c:c").ColumnWidth = 17
oWS.Range("d:d").ColumnWidth = 17

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, C + 1) = "Location Wise DP Income"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, C + 1).WrapText = True

ProgressBar1.Max = MSHFlexGrid1.Rows
ProgressBar1.Value = 0


For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        If I = 1 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, C + 1).Font.Bold = True
           oWS.Cells(mRow, C + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, C + 1).WrapText = True
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "DP Income as per Store Man"
oWS.ClearArrows
mRow = 1
C = 0
oWS.Range("a:a").ColumnWidth = 15
oWS.Range("b:b").ColumnWidth = 18
oWS.Range("c:c").ColumnWidth = 7.57
oWS.Range("d:d").ColumnWidth = 7.57
oWS.Range("e:e").ColumnWidth = 7.57
oWS.Range("f:f").ColumnWidth = 11
oWS.Range("g:g").ColumnWidth = 18
oWS.Range("h:h").ColumnWidth = 25
oWS.Range("i:i").ColumnWidth = 13.57
oWS.Range("j:j").ColumnWidth = 8
oWS.Range("k:k").ColumnWidth = 8
oWS.Range("l:l").ColumnWidth = 8
oWS.Range("m:m").ColumnWidth = 9
oWS.Range("n:n").ColumnWidth = 9
oWS.Range("o:o").ColumnWidth = 10
oWS.Range("p:p").ColumnWidth = 10
oWS.Range("q:q").ColumnWidth = 10
oWS.Range("r:r").ColumnWidth = 10
oWS.Range("s:s").ColumnWidth = 10

tmp = "a" & mRow & ":" & "s" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, C + 1) = "Details of DP Income"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, C + 1).WrapText = True


ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid2.Rows

For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        
        If I = 1 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, C + 1).Font.Bold = True
           oWS.Cells(mRow, C + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, C + 1).WrapText = True
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()
Dim wc As Variant
Call Grid_Head
Call Grid_Head1

If CmbMonth.Text = "" Then
    MsgBox "Select Month!!!"
    CmbMonth.SetFocus
    Exit Sub
End If
If CmbYear.Text = "" Then
    MsgBox "Select Year!!!"
    CmbYear.SetFocus
    Exit Sub
End If

wc = ""


mStartDate = mYear & "-" & mMonth & "-1"
mEndDate = mYear & "-" & mMonth & "-" & mTotalDays

tmp = " select a.LocationID, a.LocationName, sum(a.StoreManAmount) StoreManAmount, sum(a.NCDEXAmount) NCDEXAmount"
tmp = tmp & " from("
tmp = tmp & " Select ML.LocationID, ML.LocationName, isnull(sum(TIMG.InvoiceAmount),0) StoreManAmount, 0 'NCDEXAmount'"
tmp = tmp & " From Mst_Location ML"
'tmp = tmp & " Left Join Mst_CMP MCMP on MCMP.LocatioNID = ML.LocationID"
'tmp = tmp & " Left Join Txn_InvMonGSLDetail TIMG on TIMG.CMPID = MCMP.CMPID And TIMG.YearMonth between '2012-02-01' and '2012-02-29'"
tmp = tmp & " Left Join Txn_InvMonGSLDetail TIMG on TIMG.SM_Prefix  = ML.SM_Prefix And TIMG.YearMonth = '" & mStartDate & "' and TIMG.ExchangeTypeID = 2"
If CmbLocationID.Text <> "" Then
    tmp = tmp & " Where TIMG.SM_Prefix = '" & mSM_Prefix & "' "
End If

'tmp = tmp & wc
tmp = tmp & " group by ML.LocationID, ML.LocationName"
tmp = tmp & " Union "
tmp = tmp & " Select ML.LocationID, ML.LocationName, 0 'StoreManAmount', isnull(sum(MDIU.TotalAmt),0) NCDEXAmount From Mst_Location ML"
tmp = tmp & " Left Join Mst_DPInvoiceUpload MDIU on MDIU.LocationID = ML.LocationID And MDIU.MonthYear between '" & mStartDate & "' and '" & mEndDate & "'"
If CmbLocationID.Text <> "" Then
    tmp = tmp & " Where MDIU.LocationID = " & mLocationID
End If
'tmp = tmp & wc
tmp = tmp & " group by ML.LocationID, ML.LocationName"
tmp = tmp & " ) a"
tmp = tmp & " group by a.LocationID, a.LocationName"
tmp = tmp & " Order by a.LocationName"
Call TmpTable

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 3

If TmpRs.RecordCount > 0 Then
    LblTotalRecord.Caption = TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
    For I = 2 To TmpRs.RecordCount + 1
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!StoreManAmount), "", TmpRs!StoreManAmount)
        MSHFlexGrid1.TextMatrix(0, 1) = Val(MSHFlexGrid1.TextMatrix(I, 1)) + Val(MSHFlexGrid1.TextMatrix(0, 1))
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!NCDEXAmount), "", TmpRs!NCDEXAmount)
        MSHFlexGrid1.TextMatrix(0, 2) = Val(MSHFlexGrid1.TextMatrix(I, 2)) + Val(MSHFlexGrid1.TextMatrix(0, 2))
        MSHFlexGrid1.TextMatrix(I, 3) = Val(MSHFlexGrid1.TextMatrix(I, 1)) - Val(MSHFlexGrid1.TextMatrix(I, 2))
        MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(I, 3)) + Val(MSHFlexGrid1.TextMatrix(0, 3))
        ProgressBar1.Value = ProgressBar1.Value + 1
        TmpRs.MoveNext
    Next
    ProgressBar1.Value = ProgressBar1.Max
    Call GetDetail
    ProgressBar1.Value = ProgressBar1.Max
    MsgBox "Done!!!"
    CmdExcel.Enabled = True
Else
    MsgBox "No Record Found!!!"
    LblTotalRecord.Caption = 0
    CmdExcel.Enabled = False
End If
End Sub

Private Sub GetDetail()

mStartDate = mYear & "-" & mMonth & "-1"
mEndDate = mYear & "-" & mMonth & "-" & mTotalDays

tmp = " Select"
tmp = tmp & " ML.LocationName, MCMP.CMPName, TIMG.Mst_GSL_ID, MG.GodownNo, TIMG.StackNo, TIMG.LotNo,"
tmp = tmp & " MCm.Commodity, MCL.CLientName, MET.ExchangeType, TIMG.BalanceBags, TIMG.BalanceWeight,"
tmp = tmp & " TIMG.BagSize , TIMG.InvoiceAmount, TIMG.Rate, TIMG.NoOfDays, TIMG.TFlag, TIMG.PINo"
'MBC.BillingCycle, TIMG.BillingUnit,
tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TIMG.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = TIMG.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_Commodity MCm on MCm.CommodityID = TIMG.CommodityID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.CLientID = TIMG.CLientID"
tmp = tmp & " Inner Join Mst_ExchangeType MET on MET.ExchangeTypeID = TIMG.ExchangeTypeID"
'tmp = tmp & " Left Join Mst_BillingCycle MBC on MBC.BillingCycleID = TIMG.BillingCycleID"
'tmp = tmp & " Where TIMG.YearMonth between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Where  TIMG.ExchangeTypeID = 2 And TIMG.YearMonth = '" & mStartDate & "'"
If CmbLocationID.Text <> "" Then
    tmp = tmp & " and TIMG.SM_Prefix = '" & mSM_Prefix & "'"
End If
tmp = tmp & " Order by ML.LocationName "
Call TmpTable

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 3

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    For I = 2 To TmpRs.RecordCount + 1
        MSHFlexGrid2.TextMatrix(I, 0) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
        MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs!Mst_GSL_ID), "", TmpRs!Mst_GSL_ID)
        MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
        MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
        MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo)
        MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
        MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
        MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
        MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(I, 9)) + Val(MSHFlexGrid2.TextMatrix(0, 9))
        MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs!BalanceWeight), 0, TmpRs!BalanceWeight)
        MSHFlexGrid2.TextMatrix(0, 10) = Val(MSHFlexGrid2.TextMatrix(I, 10)) + Val(MSHFlexGrid2.TextMatrix(0, 10))
        MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs!BagSize), "", TmpRs!BagSize)
        MSHFlexGrid2.TextMatrix(I, 12) = "Daily" 'IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)
        MSHFlexGrid2.TextMatrix(I, 13) = "Quantity" 'IIf(IsNull(TmpRs!BillingUnit), "", TmpRs!BillingUnit)
'        If LCase(MSHFlexGrid2.TextMatrix(I, 13)) = "b" Then
'            MSHFlexGrid2.TextMatrix(I, 13) = "Bags"
'        ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 13)) = "q" Then
'            MSHFlexGrid2.TextMatrix(I, 13) = "Quantity"
'        End If
        MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs!InvoiceAmount), "", TmpRs!InvoiceAmount)
        MSHFlexGrid2.TextMatrix(0, 14) = Val(MSHFlexGrid2.TextMatrix(I, 14)) + Val(MSHFlexGrid2.TextMatrix(0, 14))
        MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(TmpRs!Rate), "", TmpRs!Rate)
        MSHFlexGrid2.TextMatrix(I, 16) = IIf(IsNull(TmpRs!NoOfDays), "", TmpRs!NoOfDays)
        MSHFlexGrid2.TextMatrix(I, 17) = IIf(IsNull(TmpRs!tFlag), "", TmpRs!tFlag)
        If LCase(MSHFlexGrid2.TextMatrix(I, 17)) = "d" Then
            MSHFlexGrid2.TextMatrix(I, 17) = "Deposite"
        ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 17)) = "w" Then
            MSHFlexGrid2.TextMatrix(I, 17) = "Withdraw"
        ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 17)) = "o" Then
            MSHFlexGrid2.TextMatrix(I, 17) = "Opening Balance"
        End If
        
        MSHFlexGrid2.TextMatrix(I, 18) = IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)
        TmpRs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Next
End If

End Sub


Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Call CmdClear_Click
End Sub

