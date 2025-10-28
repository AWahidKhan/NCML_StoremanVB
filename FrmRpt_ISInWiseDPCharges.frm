VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_ISInWiseDPCharges 
   Caption         =   "ISIN wise DP Charges"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   9495
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   14910
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   9045
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   14715
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
            ItemData        =   "FrmRpt_ISInWiseDPCharges.frx":0000
            Left            =   120
            List            =   "FrmRpt_ISInWiseDPCharges.frx":0002
            TabIndex        =   12
            Top             =   420
            Width           =   1680
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
            ItemData        =   "FrmRpt_ISInWiseDPCharges.frx":0004
            Left            =   1800
            List            =   "FrmRpt_ISInWiseDPCharges.frx":0006
            TabIndex        =   11
            Top             =   420
            Width           =   1560
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
            Height          =   420
            Left            =   13020
            TabIndex        =   7
            Top             =   345
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
            Height          =   420
            Left            =   10110
            TabIndex        =   6
            Top             =   345
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
            Height          =   420
            Left            =   11565
            TabIndex        =   5
            Top             =   345
            Width           =   1455
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
            Left            =   3375
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   420
            Width           =   3480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   690
            TabIndex        =   14
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   2363
            TabIndex        =   13
            Top             =   120
            Width           =   435
         End
         Begin VB.Label lblRec 
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
            Left            =   6960
            TabIndex        =   10
            Top             =   450
            Width           =   2325
         End
         Begin VB.Label lblRecordCount 
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
            Left            =   9360
            TabIndex        =   9
            Top             =   450
            Width           =   210
         End
         Begin VB.Label LblLocationID 
            AutoSize        =   -1  'True
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
            Left            =   4733
            TabIndex        =   8
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   14715
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7575
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   14505
            _ExtentX        =   25585
            _ExtentY        =   13361
            _Version        =   393216
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
End
Attribute VB_Name = "FrmRpt_ISInWiseDPCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSM_Prefix As Variant
Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I)
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then
   Exit Sub
End If
End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If CmbYear.Text = "" Then
   Exit Sub
End If
End Sub
Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("") 'Where StateID=" & mStateID & "
Else
   Call TableMst_Location("Where " & Gen_WcLocation) 'StateID=" & mStateID & " And
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   mSM_Prefix = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmdClear_Click()
ProgressBar1.Value = 0
CmbMonth.Text = ""
CmbYear.Text = ""
CmbLocationID.Text = ""
Call Grid_Head
lblRecordCount.Caption = ""

End Sub

Private Sub CmdGetReport_Click()
Dim mTotDays As Double
Dim mEndDate As Date
Dim mStartDate As Date
Dim mTotQty, mTotAmt, mRate As Variant


ProgressBar1.Value = 0
Call Grid_Head

MsgBox "Wait till next message!!!"

If CmbMonth.Text = "" Then
   MsgBox "Please Select Month."
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Please Select Year."
   Exit Sub
End If

''If CmbLocationID.Text = "" Then
''   MsgBox "Please select Location!!"
''   CmbLocationID.SetFocus
''   Exit Sub
''End If

mYear = CmbYear.Text
mMonthNo = Month(mYear & "-" & CmbMonth.Text & "-" & 1)
mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
mStartDate = CDate(mYear & "-" & CmbMonth.Text & "-" & 1)
mEndDate = CDate(mYear & "-" & CmbMonth.Text & "-" & mTotDays)

''tmp = " Select Distinct ISINID,SM_Prefix From Txn_InvMonGSLDetail"
''tmp = tmp & " Where  ExchangeTypeID = 2  And KRNo is not null And  YearMonth = '" & mStartDate & "' And CRDate <= '" & mEndDate & "'" 'DRDate <= '" & mEndDate & "'"
''''tmp = tmp & " And ISINID = 'INC000675982'"

tmp = " Select Distinct TG.ISINID,ML.LocationName,MC.Commodity "
tmp = tmp & " From Txn_InvMonGSLDetail TG"
tmp = tmp & " Inner Join Mst_Location ML On TG.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_Commodity MC On TG.CommodityID = MC.CommodityID"
tmp = tmp & " Where  TG.ExchangeTypeID = 2  And TG.KRNo is not null And  TG.YearMonth = '" & mStartDate & "' And TG.CRDate <= '" & mEndDate & "'" 'DRDate <= '" & mEndDate & "'"

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And TG.SM_Prefix='" & mSM_Prefix & "'"
End If


Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 5

For m = 1 To TmpRs.RecordCount
    mdate = mStartDate
    
    tmp = "Select TG.ISINID,TG.CRDate,Sum(TG.BalanceWeight) as 'Qty',Max(TG.NoOFDays) as 'Days',Sum(TG.InvoiceAmount) as 'InvoiceAmount',"
    tmp = tmp & " TG.Rate , MI.FED,TG.TFlag,ML.LocationName,MC.Commodity "
    tmp = tmp & " from Txn_InvMonGSLDetail TG"
    tmp = tmp & " Inner Join Mst_ISIN MI on TG.ISINID = MI.ISINID"
    tmp = tmp & " Inner Join Mst_Commodity MC on MI.CommodityID = MC.CommodityID "
    tmp = tmp & " Inner Join Mst_Location ML on TG.SM_prefix = ML.SM_prefix "
    tmp = tmp & " Where  TG.ExchangeTypeID = 2  And TG.KRNo is not null And  TG.YearMonth = '" & mStartDate & "' And TG.CRDate < '" & mStartDate & "'" 'TG.DRDate < '" & mStartDate & "'"
    tmp = tmp & " And TG.ISINID = '" & TmpRs!ISINID & "'"
    tmp = tmp & " Group by TG.ISINID,TG.CRDate,TG.Rate,MI.FED,TG.TFlag,ML.LocationName,MC.Commodity"
    
    Call Tmp1Table
    
    tmp = " Select TG.ISINID,TG.CRDate,Sum(TG.BalanceWeight) as 'Qty',Max(TG.NoOFDays) as 'Days',Sum(TG.InvoiceAmount) as 'InvoiceAmount',"
    tmp = tmp & " TG.Rate , MI.FED,TG.TFlag,TG.CxTFDate,ML.LocationName,MC.Commodity "
    tmp = tmp & " from Txn_InvMonGSLDetail TG"
    tmp = tmp & " Inner Join Mst_ISIN MI on TG.ISINID = MI.ISINID"
    tmp = tmp & " Inner Join Mst_Commodity MC on MI.CommodityID = MC.CommodityID "
    tmp = tmp & " Inner Join Mst_Location ML on TG.SM_prefix = ML.SM_prefix "
    tmp = tmp & " Where  TG.ExchangeTypeID = 2  And TG.KRNo is not null And  TG.YearMonth = '" & mStartDate & "' And TG.CRDate < '" & mEndDate & "'" ' and TG.CRDate >= '" & mStartDate & "'"
    tmp = tmp & " And TG.ISINID = '" & TmpRs!ISINID & "'"
    tmp = tmp & " Group by TG.ISINID,TG.CRDate,TG.Rate,MI.FED,TG.TFlag,TG.CxTFDate,ML.LocationName,MC.Commodity"
    
    
    Call Tmp2Table
    
''    tmp = " Select TG.ISINID,TG.DRDate,Sum(TG.BalanceWeight) as 'Qty',Max(TG.NoOFDays) as 'Days',Sum(TG.InvoiceAmount) as 'InvoiceAmount',"
''    tmp = tmp & " TG.Rate , MI.FED,TG.TFlag"
''    tmp = tmp & " from Txn_InvMonGSLDetail TG"
''    tmp = tmp & " Inner Join Mst_ISIN MI on TG.ISINID = MI.ISINID"
''    tmp = tmp & " Where  TG.ExchangeTypeID = 2  And TG.KRNo is not null And  TG.YearMonth = '" & mStartDate & "' And TG.DRDate < '" & mEndDate & "'" ' and TG.DRDate >= '" & mStartDate & "'"
''    tmp = tmp & " And TG.ISINID = '" & TmpRs!ISINID & "'"
''    tmp = tmp & " Group by TG.ISINID,TG.DRDate,TG.Rate,MI.FED,TG.TFlag"
''
    
    tmp = " Select TG.ISINID,TG.CRDate,Sum(TG.BalanceWeight) as 'Qty',Max(TG.NoOFDays) as 'Days',Sum(TG.InvoiceAmount) as 'InvoiceAmount',"
    tmp = tmp & " TG.Rate , MI.FED,TG.TFlag,ML.LocationName,MC.Commodity "
    tmp = tmp & " from Txn_InvMonGSLDetail TG"
    tmp = tmp & " Inner Join Mst_ISIN MI on TG.ISINID = MI.ISINID"
    tmp = tmp & " Inner Join Mst_Commodity MC on MI.CommodityID = MC.CommodityID "
    tmp = tmp & " Inner Join Mst_Location ML on TG.SM_prefix = ML.SM_prefix "
    tmp = tmp & " Where  TG.ExchangeTypeID = 2  And TG.KRNo is not null And  TG.YearMonth = '" & mStartDate & "' And TG.CRDate < '" & mEndDate & "'" ' and TG.CRDate >= '" & mStartDate & "'"
    tmp = tmp & " And TG.ISINID = '" & TmpRs!ISINID & "'"
    tmp = tmp & " Group by TG.ISINID,TG.CRDate,TG.Rate,MI.FED,TG.TFlag,ML.LocationName,MC.Commodity"
    
    
    Call Tmp3Table
    
    mTotQty = 0
    mTotAmt = 0
    mRate = 0
    For I = 1 To mTotDays
        r = MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.TextMatrix(r, 0) = Format(mdate, "dd-MMM-yyyy")
        MSHFlexGrid1.TextMatrix(r, 1) = I
        MSHFlexGrid1.TextMatrix(r, 2) = TmpRs!LocationName 'TmpRs!SM_Prefix
        MSHFlexGrid1.TextMatrix(r, 3) = TmpRs!ISINID
        MSHFlexGrid1.TextMatrix(r, 4) = TmpRs!Commodity
        If I = 1 And TmpRs1.RecordCount > 0 Then
           For rw = 1 To TmpRs1.RecordCount
               If TmpRs1!FED < mStartDate Then
                  mTotAmt = mTotAmt + Format(TmpRs1!Qty * (TmpRs1!Rate * 2), "#####0.00")
                  mRate = TmpRs1!Rate * 2
               Else
                  mTotAmt = mTotAmt + Format(TmpRs1!Qty * TmpRs1!Rate, "#####0.00")
                  mRate = TmpRs1!Rate
               End If
               MSHFlexGrid1.TextMatrix(r, 6) = TmpRs1!Days
               mTotQty = mTotQty + TmpRs1!Qty
               'mRate = TmpRs1!Rate
               TmpRs1.MoveNext
           Next
        End If
        If I > 1 Then
           TmpRs2.Filter = "CxTFDate = '" & mdate - 1 & "'"
           If TmpRs2.RecordCount > 0 Then
              For rw = 1 To TmpRs2.RecordCount
                  If TmpRs2!tFlag = "W" Then
                     mTotQty = mTotQty - TmpRs2!Qty
                     MSHFlexGrid1.TextMatrix(r, 6) = TmpRs2!Days
                     If TmpRs2!FED < mStartDate Then
                        mTotAmt = mTotAmt - Format(TmpRs2!Qty * (TmpRs2!Rate * 2), "#####0.00")
                        mRate = TmpRs2!Rate * 2
                     Else
                        mTotAmt = mTotAmt - Format(TmpRs2!Qty * TmpRs2!Rate, "#####0.00")
                        mRate = TmpRs2!Rate
                     End If
                  End If
                  TmpRs2.MoveNext
              Next
           End If

            
           TmpRs2.Filter = ""
            
           TmpRs3.Filter = "CRDate = '" & mdate & "'"
           If TmpRs3.RecordCount > 0 Then
              For rw = 1 To TmpRs3.RecordCount
                  If TmpRs3!FED < mStartDate Then
                     mTotAmt = mTotAmt + Format(TmpRs3!Qty * (TmpRs3!Rate * 2), "#####0.00")
                     mRate = TmpRs3!Rate * 2
                  Else
                     mTotAmt = mTotAmt + Format(TmpRs3!Qty * TmpRs3!Rate, "#####0.00")
                     mRate = TmpRs3!Rate
                  End If
                  MSHFlexGrid1.TextMatrix(r, 6) = TmpRs3!Days
                  If TmpRs2!tFlag <> "W" Then
                     mTotQty = mTotQty + TmpRs3!Qty
                  End If
                  TmpRs3.MoveNext
              Next
           End If
           TmpRs3.Filter = ""
            
        End If
        
        MSHFlexGrid1.TextMatrix(r, 5) = mTotQty
        MSHFlexGrid1.TextMatrix(r, 8) = Format(mTotAmt, "#####0.00")
        MSHFlexGrid1.TextMatrix(r, 7) = mRate
        
        MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(r, 5))
        MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(r, 8))
        MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(r, 7))
        mdate = mdate + 1
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
    TmpRs.MoveNext
Next

ProgressBar1.Value = ProgressBar1.Max
CmdExcel.Enabled = True
MsgBox "Done!!"

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 9
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.MergeCells = flexMergeRestrictRows

MSHFlexGrid1.TextMatrix(1, 0) = "Date"
MSHFlexGrid1.TextMatrix(1, 1) = "Cycle"
MSHFlexGrid1.TextMatrix(1, 2) = "Location"
MSHFlexGrid1.TextMatrix(1, 3) = "ISIN"
MSHFlexGrid1.TextMatrix(1, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(1, 5) = "Qty"
MSHFlexGrid1.TextMatrix(1, 6) = "Days"
MSHFlexGrid1.TextMatrix(1, 7) = "Rate"
MSHFlexGrid1.TextMatrix(1, 8) = "Invoice Amt"


MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 2200
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200


MSHFlexGrid1.RowHeight(1) = 800

End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ISINwiseDPCharges.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ISINwiseDPCharges.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
''FrmLoadAccess = GetFrmLoadAccess(Me.Name)
''If FrmLoadAccess(0) = False Then
''   Frame4.Enabled = False
''   MsgBox "Access denied !!!!!!!!!"
''   Exit Sub
''End If
Frame4.Enabled = True
Call CmdClear_Click

End Sub


