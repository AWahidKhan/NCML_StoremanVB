VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_ShortageMIS 
   Caption         =   "Shortage MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   10560
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15000
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9495
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   14535
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   9000
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   14280
         Begin VB.OptionButton OptSummary 
            Caption         =   "Summary"
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
            Left            =   1560
            TabIndex        =   22
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton OptDetail 
            Caption         =   "Detail"
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
            Left            =   240
            TabIndex        =   21
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox CmbGodown 
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
            Left            =   10245
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   420
            Width           =   3480
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
            Left            =   12330
            TabIndex        =   8
            Top             =   885
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
            Left            =   9420
            TabIndex        =   7
            Top             =   885
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
            Left            =   10875
            TabIndex        =   6
            Top             =   885
            Width           =   1455
         End
         Begin VB.ComboBox CmbCMPName 
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
            Left            =   6750
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   420
            Width           =   3375
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
            Left            =   3150
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   420
            Width           =   3480
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   420
            Width           =   1365
            _ExtentX        =   2408
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
            CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
            Format          =   61210627
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   375
            Left            =   1620
            TabIndex        =   11
            Top             =   420
            Width           =   1380
            _ExtentX        =   2434
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
            CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
            Format          =   61210627
            CurrentDate     =   39850
            MinDate         =   38353
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
            Left            =   5715
            TabIndex        =   18
            Top             =   975
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
            Left            =   8115
            TabIndex        =   17
            Top             =   975
            Width           =   210
         End
         Begin VB.Label LblCommodityID 
            Caption         =   "Godown"
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
            Left            =   11280
            TabIndex        =   16
            Top             =   120
            Width           =   1320
         End
         Begin VB.Label Label2 
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
            Left            =   1920
            TabIndex        =   15
            Top             =   165
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "From  Date"
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
            Left            =   300
            TabIndex        =   14
            Top             =   165
            Width           =   1005
         End
         Begin VB.Label LblCMPName 
            Caption         =   "CMP"
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
            Left            =   8055
            TabIndex        =   13
            Top             =   180
            Width           =   600
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
            Left            =   4440
            TabIndex        =   12
            Top             =   180
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   14295
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6975
            Left            =   60
            TabIndex        =   2
            Top             =   150
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   12303
            _Version        =   393216
            Rows            =   3
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
            _Band(0).ColHeader=   1
         End
      End
   End
End
Attribute VB_Name = "FrmRpt_ShortageMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxRow, mCMPID, mLocationID, loc_name, mGodownID, mSM_Prefix As Variant
Dim I As Integer
Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub


Private Sub CmbGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP"
   CmbCMPName.SetFocus
   Exit Sub
End If
 
tmp = CmbGodown.Text
Call TableMst_Godown(" Where CMPID = " & mCMPID & "")
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID

End Sub

'----
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation <> "" Then
   qry = " Where  " & Gen_WcLocation & " " 'And StateID= " & mStateID & ""
Else
   qry = " Where LocationID <> 0 " 'And StateID= " & mStateID & ""
End If

Call TableMst_Location(qry)
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "First Select State  !!!! "
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix


End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmdExcel.Enabled = False
lblRecordCount.Caption = ""
OptDetail.Value = True
ProgressBar1.Value = 0

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Call CmdClear_Click
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.Font.Size = 10

If OptDetail.Value = True Then

   MSHFlexGrid1.Cols = 15
   MSHFlexGrid1.TextMatrix(0, 0) = "Gain / Shortage"
   MSHFlexGrid1.TextMatrix(0, 1) = "GSL ID"
   MSHFlexGrid1.TextMatrix(0, 2) = "Location"
   MSHFlexGrid1.TextMatrix(0, 3) = "CMP Name"
   MSHFlexGrid1.TextMatrix(0, 4) = "Commodity"
   MSHFlexGrid1.TextMatrix(0, 5) = "Godown No"
   MSHFlexGrid1.TextMatrix(0, 6) = "Stack No"
   MSHFlexGrid1.TextMatrix(0, 7) = "Lot No" '9'
   MSHFlexGrid1.TextMatrix(0, 8) = "Bags" '10'
   MSHFlexGrid1.TextMatrix(0, 9) = "Deposit Wt" '11'
   MSHFlexGrid1.TextMatrix(0, 10) = "Withdrawan Wt"
   MSHFlexGrid1.TextMatrix(0, 11) = "Gain / Loss"
   MSHFlexGrid1.TextMatrix(0, 12) = "Flag"
   MSHFlexGrid1.TextMatrix(0, 13) = "CTDF No"
   MSHFlexGrid1.TextMatrix(0, 14) = "Date"
   
    
    
   MSHFlexGrid1.ColWidth(0) = 1200
   MSHFlexGrid1.ColWidth(1) = 1000
   MSHFlexGrid1.ColWidth(2) = 1500
   MSHFlexGrid1.ColWidth(3) = 1400
   MSHFlexGrid1.ColWidth(4) = 1600
   MSHFlexGrid1.ColWidth(5) = 1000
   MSHFlexGrid1.ColWidth(6) = 1000
   MSHFlexGrid1.ColWidth(7) = 1000
   MSHFlexGrid1.ColWidth(8) = 800
   MSHFlexGrid1.ColWidth(9) = 800
   MSHFlexGrid1.ColWidth(10) = 1000
   MSHFlexGrid1.ColWidth(11) = 1000
   MSHFlexGrid1.ColWidth(12) = 1000
   MSHFlexGrid1.ColWidth(13) = 1000
   MSHFlexGrid1.ColWidth(14) = 1200
   
ElseIf OptSummary.Value = True Then

   MSHFlexGrid1.Cols = 6
   MSHFlexGrid1.TextMatrix(0, 0) = "Month Of Deposit"
   MSHFlexGrid1.TextMatrix(0, 1) = "Deposited Weight Of Withdrawn Qty"
   MSHFlexGrid1.TextMatrix(0, 2) = "Actual Withdrawn Qty"
   MSHFlexGrid1.TextMatrix(0, 3) = "Shortage"
   MSHFlexGrid1.TextMatrix(0, 4) = "Excess"
   MSHFlexGrid1.TextMatrix(0, 5) = "Net Shortage/ Excess"
   MSHFlexGrid1.ColWidth(0) = 2000
   MSHFlexGrid1.ColWidth(1) = 2500
   MSHFlexGrid1.ColWidth(2) = 1500
   MSHFlexGrid1.ColWidth(3) = 1400
   MSHFlexGrid1.ColWidth(4) = 1600
   MSHFlexGrid1.ColWidth(5) = 2000

End If
MSHFlexGrid1.RowHeight(0) = 700



End Sub

Private Sub CmdGetReport_Click()

Dim tmp1 As String
Dim wc As String

Call Grid_Head

wc = ""
ProgressBar1.Value = 0

MsgBox " Wait till next message !!!"

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date Must be less than To Date !!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If Gen_DBType <> "MDB" Then
   wc = " and gsl.cdtfdate between '" & TxtFromDate.Value & "' and '" & TxtToDate.Value & "' "
Else
   wc = " and gsl.cdtfdate between #" & TxtFromDate.Value & "# and #" & TxtToDate.Value & "# "
End If



If OptDetail.Value = True Then

    
   If CmbLocationID.Text <> "" Then
      wc = wc & " And l.LocationName = '" & CmbLocationID.Text & "' "
   Else
      If Gen_WcLocation <> "" Then
         wc = wc & " And c." & Gen_WcLocation
      End If
   End If
   
  
   If CmbCMPName.Text <> "" Then
      wc = wc & " And c.CMPName = '" & CmbCMPName.Text & "'  "
   End If
   
   If CmbGodown.Text <> "" Then
      wc = wc & " And G.GodownNo = '" & CmbGodown.Text & "'"
   End If
   
   wc1 = " Order By l.locationName"

   tmp = "select 'Gain' As 'Type',"
   tmp = tmp & " gsl.GSLID,l.locationName,c.CMPName,co.Commodity,g.GodownNo,"
   tmp = tmp & " gsl.StackNo , gsl.LotNo, gsl.CDTFBags, gsl.DepositeWeight, gsl.CDTFWeight, gsl.BalanceWeight, gsl.Flag, gsl.CdtfNo, gsl.CDTFDate from mst_gsl gsl "
   tmp = tmp & " inner join mst_Cmp c on gsl.cmpid = c.cmpid"
   tmp = tmp & " inner join mst_Godown g on gsl.GodownID = g.GodownID"
   tmp = tmp & " inner join mst_location l on c.locationID = l.locationID "
   tmp = tmp & " inner join mst_Commodity co on gsl.CommodityID = co.CommodityID "
   tmp = tmp & " where  gsl.flag = 'E' And gsl.BalanceWeight < 0 " & wc '& Wc1
   tmp = tmp & " Union All"
   tmp = tmp & " select 'Shortage' As 'Type',"
   tmp = tmp & " gsl.GSLID,l.locationName,c.CMPName,co.Commodity,g.GodownNo,"
   tmp = tmp & " gsl.StackNo , gsl.LotNo, gsl.CDTFBags, gsl.DepositeWeight, gsl.CDTFWeight, gsl.BalanceWeight, gsl.Flag, gsl.CdtfNo, gsl.CDTFDate"
   tmp = tmp & " from mst_gsl gsl"
   tmp = tmp & " inner join mst_Cmp c on gsl.cmpid = c.cmpid"
   tmp = tmp & " inner join mst_Godown g on gsl.GodownID = g.GodownID"
   tmp = tmp & " inner join mst_location l on c.locationID = l.locationID "
   tmp = tmp & " inner join mst_Commodity co on gsl.CommodityID = co.CommodityID"
   tmp = tmp & " where gsl.flag = 'E' And gsl.BalanceWeight > 0 " & wc '& Wc1
   tmp = tmp & " Union All"
   tmp = tmp & " select 'NIL' As 'Type',"
   tmp = tmp & " gsl.GSLID,l.locationName,c.CMPName,co.Commodity,g.GodownNo,"
   tmp = tmp & " gsl.StackNo , gsl.LotNo, gsl.CDTFBags, gsl.DepositeWeight, gsl.CDTFWeight, gsl.BalanceWeight, gsl.Flag, gsl.CdtfNo, gsl.CDTFDate"
   tmp = tmp & " from mst_gsl gsl"
   tmp = tmp & " inner join mst_Cmp c on gsl.cmpid = c.cmpid"
   tmp = tmp & " inner join mst_Godown g on gsl.GodownID = g.GodownID"
   tmp = tmp & " inner join mst_location l on c.locationID = l.locationID "
   tmp = tmp & " inner join mst_Commodity co on gsl.CommodityID = co.CommodityID"
   tmp = tmp & " where gsl.flag = 'E' And gsl.BalanceWeight = 0 " & wc '& Wc1
   
   Call TmpTable
     
   If TmpRs.RecordCount > 0 Then
      ProgressBar1.Max = TmpRs.RecordCount + 1
      MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
      For I = 1 To TmpRs.RecordCount
          For c = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
          Next
          MSHFlexGrid1.TextMatrix(I, 14) = Format(MSHFlexGrid1.TextMatrix(I, 14), "dd-mmm-yyyy")
          TmpRs.MoveNext
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
      ProgressBar1.Value = ProgressBar1.Max
      CmdExcel.Enabled = True
      MsgBox ("Done!!!")
   Else
      MsgBox "No record found for given Date!!!"
      Call Grid_Head
      Exit Sub
   End If
ElseIf OptSummary.Value = True Then

   mFromDate = TxtFromDate.Value
   mToDate = TxtToDate.Value
   
   mRows = DateDiff("M", mFromDate, mToDate)
   MSHFlexGrid1.Rows = mRows + 3
   mFromMonth = Month(TxtFromDate.Value)
   mFromYear = Year(TxtFromDate.Value)
   ProgressBar1.Max = mRows + 1
   For I = 1 To mRows + 1
        
       MSHFlexGrid1.TextMatrix(I, 0) = MonthName(mFromMonth) & "-" & mFromYear
       If mFromMonth = 12 Then
          mFromYear = mFromYear + 1
          mFromMonth = 1
       Else
          mFromMonth = mFromMonth + 1
       End If
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = 0
   If CmbLocationID.Text <> "" Then
      wc = wc & " And gsl.SM_Prefix = '" & mSM_Prefix & "' "
   Else
'      If Gen_WcLocation <> "" Then
'         wc = wc & " And c." & Gen_WcLocation
'      End If
   End If
   
   
   If CmbCMPName.Text <> "" Then
      wc = wc & " And gsl.CMPID = " & mCMPID & "  "
   Else
      If Gen_WcCMP <> "" Then
         wc = wc & " And gsl." & Gen_WcCMP
      End If
   End If
   
   If CmbGodown.Text <> "" Then
      wc = wc & " And GSL.GodownID = " & mGodownID & ""
   End If

   tmp = "Select Month(CDTFDate) As 'Month',"
   tmp = tmp & " Year(CDTFDate) As 'Year',Sum(DepositeWeight) as 'Dep Wt',"
   tmp = tmp & " Sum(CDTFWeight) as 'Withd Wt',"
   tmp = tmp & " Sum(BalanceWeight) as 'Bal Wt' From Mst_GSL gsl"
   tmp = tmp & " Where BalanceWeight > 0 and gsl.flag = 'E' "
   tmp = tmp & wc
   tmp = tmp & " Group By Month(CDTFDate),Year(CDTFDate)"
   tmp = tmp & " Order By Year(CDTFDate),Month(CDTFDate)"
    
   Call TmpTable

   If TmpRs.RecordCount > 0 Then
      ProgressBar1.Max = TmpRs.RecordCount + 1
      For I = 1 To TmpRs.RecordCount
          For c = 1 To MSHFlexGrid1.Cols - 3
              MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c + 1)), "", TmpRs.Fields(c + 1))
          Next
          'MSHFlexGrid1.TextMatrix(I, 4) = MSHFlexGrid1.TextMatrix(I, 3)
          'MSHFlexGrid1.TextMatrix(I, 3) = ""
          TmpRs.MoveNext
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
      ProgressBar1.Value = ProgressBar1.Max
   End If

   ProgressBar1.Value = 0

   tmp = "Select Month(CDTFDate) As 'Month',"
   tmp = tmp & " Year(CDTFDate) As 'Year',Sum(DepositeWeight) as 'DepWt',"
   tmp = tmp & " Sum(CDTFWeight) as 'WithdWt',"
   tmp = tmp & " Sum(BalanceWeight) as 'BalWt' From Mst_GSL gsl"
   tmp = tmp & " Where BalanceWeight < 0 and gsl.flag = 'E' "
   tmp = tmp & wc
   tmp = tmp & " Group By Month(CDTFDate),Year(CDTFDate)"
   tmp = tmp & " Order By Year(CDTFDate),Month(CDTFDate)"
    
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      ProgressBar1.Max = TmpRs.RecordCount + 1
      For I = 1 To TmpRs.RecordCount
          For x = 1 To MSHFlexGrid1.Rows - 1
              If MSHFlexGrid1.TextMatrix(x, 0) = MonthName(TmpRs!Month) & "-" & TmpRs!Year Then
                 MSHFlexGrid1.TextMatrix(x, 1) = Val(MSHFlexGrid1.TextMatrix(x, 1)) + IIf(IsNull(TmpRs!DepWt), 0, TmpRs!DepWt)
                 MSHFlexGrid1.TextMatrix(x, 2) = Val(MSHFlexGrid1.TextMatrix(x, 2)) + IIf(IsNull(TmpRs!WithdWt), 0, TmpRs!WithdWt)
                 MSHFlexGrid1.TextMatrix(x, 4) = IIf(IsNull(TmpRs!BalWt), 0, TmpRs!BalWt)
                 MSHFlexGrid1.TextMatrix(x, 4) = Abs(MSHFlexGrid1.TextMatrix(x, 4))
                 MSHFlexGrid1.TextMatrix(x, 5) = Val(MSHFlexGrid1.TextMatrix(x, 3)) - Val(MSHFlexGrid1.TextMatrix(x, 4))
                 TmpRs.MoveNext
                 Exit For
              End If
          Next
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End If
   ProgressBar1.Value = ProgressBar1.Max
   CmdExcel.Enabled = True


End If



lblRecordCount.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp

Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ShortageMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ShortageMIS.xls")
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
           If c = 24 Then
              oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid1.TextMatrix(I, c)
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbCMPName_GotFocus()


If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text

If Gen_WcCMP <> "" Then
   qry = " Where " & Gen_WcCMP & " and LocationID =" & mLocationID & " "
Else
   qry = " Where LocationID =" & mLocationID & ""
End If
'qry = " Where LocationID =" & mLocationID & " "
Call TableMst_CMP(qry)
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
'----
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
    mCMPID = Null
   Exit Sub
End If
RsMst_CMP.MoveFirst

RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
'dtCurrentDate = RsMst_CMP!CurrentDate 'CDate(Format(RsMst_CMP!CurrentDate, "mm/dd/yyyy"))
'RptDate.Value = RsMst_CMP!CurrentDate
End Sub

Private Sub SetScrollBars()
'    Frame1.Width = 10000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub OptSelected_Click()
If OptSelected.Value = True Then
   Call Grid_Head
    Else
End If
End Sub

Private Sub OptSummarized_Click()
If OptSummarized.Value = True Then
   Call Grid_Head
   Else
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub OptDetail_Click()

Call Grid_Head
lblRecordCount.Caption = ""
End Sub

Private Sub OptSummary_Click()
Call OptDetail_Click
End Sub

Private Sub VScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l
'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub

Private Sub HScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub


