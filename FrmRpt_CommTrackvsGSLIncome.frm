VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_CommTrackvsGSLIncome 
   Caption         =   "CommTrack VS GSL Income"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   7
      Top             =   45
      Width           =   15015
      Begin VB.TextBox TxtLotNo 
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
         Height          =   375
         Left            =   6480
         TabIndex        =   3
         Top             =   270
         Width           =   4095
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
         Left            =   13710
         TabIndex        =   6
         Top             =   750
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
         Left            =   11280
         TabIndex        =   4
         Top             =   750
         Width           =   1215
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
         Left            =   12480
         TabIndex        =   5
         Top             =   750
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   111607811
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   270
         Width           =   1260
         _ExtentX        =   2223
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   111607811
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label7 
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
         Left            =   8280
         TabIndex        =   18
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label6 
         Caption         =   " GSL Detail Record Found :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label5 
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
         Left            =   4680
         TabIndex        =   13
         Top             =   810
         Width           =   90
      End
      Begin VB.Label Label4 
         Caption         =   "Commtrack lots not in Store-Man Records Found :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CMS E Lot No "
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
         Left            =   5160
         TabIndex        =   11
         Top             =   330
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   2520
         TabIndex        =   10
         Top             =   337
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   337
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8535
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   8220
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   7935
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   13996
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         TabCaption(0)   =   "Commtrack lots not in Store-Man"
         TabPicture(0)   =   "FrmRpt_CommTrackvsGSLIncome.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Store Man GSL Detail"
         TabPicture(1)   =   "FrmRpt_CommTrackvsGSLIncome.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid2"
         Tab(1).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   7215
            Left            =   -74880
            TabIndex        =   15
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12726
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7335
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12938
            _Version        =   393216
            FixedCols       =   0
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
      End
   End
End
Attribute VB_Name = "FrmRpt_CommTrackvsGSLIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grid_Head1()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 11
     .Font.Size = 10
     .WordWrap = True
     .TextMatrix(0, 0) = "Month Year"
     .TextMatrix(0, 1) = "Lot No"
     .TextMatrix(0, 2) = "Exchange Amount"
     .TextMatrix(0, 3) = "No of Days as per Exchange"
     .TextMatrix(0, 4) = "ISIN ID"
     .TextMatrix(0, 5) = "GSL ID"
     .TextMatrix(0, 6) = "Location"
     .TextMatrix(0, 7) = "Rate"
     .TextMatrix(0, 8) = "No Of Days as per GSL "
     .TextMatrix(0, 9) = "GSL Income"
     .TextMatrix(0, 10) = "Difference [Exchange-GSL] "
     .ColWidth(0) = 2000
     .ColWidth(1) = 3000
     .ColWidth(2) = 3000
     .ColWidth(3) = 3000
     .ColWidth(4) = 0
     .ColWidth(5) = 0
     .ColWidth(6) = 0
     .ColWidth(7) = 0
     .ColWidth(8) = 0
     .ColWidth(9) = 0
     .ColWidth(10) = 2000
     .RowHeight(0) = 800
End With
End Sub
Private Sub Grid_Head2()
With MSHFlexGrid2
     .Clear
     .Rows = 2
     .Cols = 18
     .Font.Size = 10
     .WordWrap = True
     .TextMatrix(0, 0) = "Month Year"
     .TextMatrix(0, 1) = "GSLID"
     .TextMatrix(0, 2) = "Location"
     .TextMatrix(0, 3) = "CMP"
     .TextMatrix(0, 4) = "GSL"
     .TextMatrix(0, 5) = "Commodity"
     .TextMatrix(0, 6) = "Client"
     .TextMatrix(0, 7) = "KR NO"
     .TextMatrix(0, 8) = "KR Date"
     .TextMatrix(0, 9) = "ISIN No"
     .TextMatrix(0, 10) = "DR Date"
     .TextMatrix(0, 11) = "Transaction Flag" '"Balance Bag"
     .TextMatrix(0, 12) = "GSL Flag"
     .TextMatrix(0, 13) = "Balance Qty"
     .TextMatrix(0, 14) = "No of Days"
     .TextMatrix(0, 15) = "Rate"
     .TextMatrix(0, 16) = "Invoice Amt"
     .TextMatrix(0, 17) = "Commtract Invoice amount"
     
     
     .ColWidth(0) = 1300
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     .ColWidth(3) = 1500
     .ColWidth(4) = 1000
     .ColWidth(5) = 2000
     .ColWidth(6) = 1000
     .ColWidth(7) = 1000
     .ColWidth(8) = 1500
     .ColWidth(9) = 1800
     .ColWidth(10) = 1800
     .ColWidth(11) = 1800
     .ColWidth(12) = 1800
     .ColWidth(13) = 1800
     .ColWidth(14) = 1800
     .ColWidth(15) = 1800
     .ColWidth(16) = 1800
     .RowHeight(0) = 600
End With
End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date - 1
TxtToDate.Value = Date
TxtLotNo.Text = ""
Call Grid_Head1
Call Grid_Head2
End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head1

MsgBox "Wait till next message!!!"
wc = ""
Label5.Caption = "-"
wc = " Where TMCF.MonthYear Between '" & Format(TxtFromDate.Value, "yyyy-MM") & "-1" & "'" & " And " & "'" & Format(TxtToDate.Value, "yyyy-MM") & "-1" & "' And  TINV.ISINID is null "
If Trim(TxtLotNo.Text) <> "" Then
   wc = wc & " and TMCF.LotID like '%" & TxtLotNo.Text & "%'"
End If
ProgressBar1.Min = 0

'tmp = "Select TMCF.MonthYear, TMCF.LotID, sum(TMCF.Amount) 'ExchangeAmount', TINV.ISINID, TINV.GSLID,TINV.SM_Prefix, Mst_L.LocationName, TINV.Rate, Sum(TINV.NoOfDays) NoOfDays, Sum(TINV.InvoiceAmount) 'GSLIncome' "
'tmp = tmp & " From Txn_MonthlyCommtrackFileDetail TMCF"
'tmp = tmp & " Left Join Txn_InvMonGSLDetail TINV On TMCF.LotID = TINV.ISINID And TMCF.MonthYear = TINV.YearMonth And TINV.ExchangeTypeID = 2 And TINV.CommtrackFlag = '1'"
'tmp = tmp & " Left Join Mst_Location as Mst_L on Mst_L.SM_Prefix = TINV.SM_Prefix"
'tmp = tmp & wc
'tmp = tmp & " Group By TMCF.MonthYear, TMCF.LotID,TINV.ISINID, TINV.GSLID, TINV.SM_Prefix, TINV.Rate, Mst_L.LocationName"
'tmp = tmp & " Order By TMCF.MonthYear,TINV.SM_Prefix"

tmp = " Select TMCF.MonthYear, TMCF.LotID, sum(TMCF.Amount) 'ExchangeAmount', sum(TMCF.HoldDuration) 'HoldDuration' ,"
tmp = tmp & " TINV.ISINID, TINV.GSLID,TINV.SM_Prefix, Mst_L.LocationName, TINV.Rate, TINV.NoOfDays ,"
tmp = tmp & " Sum (TInv.InvoiceAmount) 'GSLIncome'"
tmp = tmp & " From Txn_MonthlyCommtrackFileDetail TMCF"
tmp = tmp & " Left Join Txn_InvMonGSLDetail TINV On TMCF.LotID = TINV.ISINID And TMCF.MonthYear = TINV.YearMonth And TINV.ExchangeTypeID = 2 And TINV.CommtrackFlag = '1'"
tmp = tmp & " Left Join Mst_Location as Mst_L on Mst_L.SM_Prefix = TINV.SM_Prefix"
tmp = tmp & wc
tmp = tmp & " Group By TMCF.MonthYear, TMCF.LotID,TINV.ISINID, TINV.GSLID, TINV.SM_Prefix, TINV.Rate, Mst_L.LocationName,"
tmp = tmp & " TInv.NoOfDays"
tmp = tmp & " Order By TMCF.MonthYear,TINV.SM_Prefix"


Call Tmp2Table
'tmp = "Select LocationID, LocationName, SM_Prefix from Mst_Location"
'Call Tmp1Table
'Label5.Caption = TmpRs2.RecordCount
If TmpRs2.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs2.RecordCount
   Label5.Caption = TmpRs2.RecordCount
   TmpRs2.MoveFirst
   For I = 1 To TmpRs2.RecordCount
       ProgressBar1.Value = I
       'TmpRs1.MoveFirst
       'TmpRs1.Find "SM_Prefix = '" & IIf(IsNull(TmpRs2!SM_PREFIX), "", TmpRs2!SM_PREFIX) & "'"
       MSHFlexGrid1.TextMatrix(I, 0) = Format(TmpRs2!MonthYear, "MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs2!LotId
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs2!ExchangeAmount
       MSHFlexGrid1.TextMatrix(I, 3) = TmpRs2!HoldDuration
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!ISINID), "", TmpRs2!ISINID)
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!GSLID), "", TmpRs2!GSLID)
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!Locationname), "", TmpRs2!Locationname)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs2!Rate), "", TmpRs2!Rate)
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs2!NoOfDays), "", TmpRs2!NoOfDays)
       MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs2!GSLIncome), "", TmpRs2!GSLIncome)
       MSHFlexGrid1.TextMatrix(I, 10) = TmpRs2!ExchangeAmount - IIf(IsNull(TmpRs2!GSLIncome), 0, TmpRs2!GSLIncome)
       TmpRs2.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Next
End If
Call Grid_Head2

wc = ""
wc = " And TIMGD.YearMonth Between '" & Format(TxtFromDate.Value, "yyyy-MM") & "-1" & "'" & " And " & "'" & Format(TxtToDate.Value, "yyyy-MM") & "-1" & "'"

If Trim(TxtLotNo.Text) <> "" Then
   wc = wc & "And TIMGD.ISINID like '%" & TxtLotNo.Text & "%'"
End If

tmp = "Select TIMGD.YearMonth, TIMGD.Mst_GSL_ID, Mst_L.LocationName, MC.CMPName, MG.GodownNo, TIMGD.StackNo, TIMGD.LotNo, MG.GodownNo + '~'+ TIMGD.StackNo + '~'+ TIMGD.LotNo as GSL, "
tmp = tmp & " MComm.Commodity, Mst_C.ClientName, TIMGD.KRNO, TIMGD.KRDate, TIMGD.ISINID, TIMGD.DRDate,TIMGD.TFlag,TIMGD.Status, " 'TIMGD.BalanceBags,"
tmp = tmp & " TIMGD.BalanceWeight , TIMGD.NoofDays, TIMGD.Rate, TIMGD.InvoiceAmount, TIMGD.CommtrackIncome "
tmp = tmp & " From Txn_InvMonGSLDetail as TIMGD "
tmp = tmp & " Inner Join Mst_Client as Mst_C on Mst_C.ClientID  = TIMGD.ClientID "
tmp = tmp & " Inner Join Mst_CMP as MC On TIMGD.CMPID = MC.CMPID "
tmp = tmp & " Inner Join Mst_Godown as MG On TIMGD.GodownID = MG.GodownID "
tmp = tmp & " Inner Join Mst_Commodity as MComm On TIMGD.CommodityID = MComm.CommodityID "
tmp = tmp & " Inner Join Mst_Location as Mst_L on Mst_L.SM_Prefix = TIMGD.SM_Prefix "
tmp = tmp & " Where TIMGD.ExchangeTypeID = 2 And TIMGD.CommtrackFlag = 1 " & wc
tmp = tmp & " Order By TIMGD.YearMonth, Mst_L.LocationName, TIMGD.Mst_GSL_ID "

Call Tmp1Table

If Not TmpRs1.EOF Then
   ProgressBar1.Max = TmpRs1.RecordCount
   Label7.Caption = TmpRs1.RecordCount
   TmpRs1.MoveFirst
   For I = 1 To TmpRs1.RecordCount
       ProgressBar1.Value = I
       MSHFlexGrid2.TextMatrix(I, 0) = Format(TmpRs1!YearMonth, "MMM-yyyy")
       MSHFlexGrid2.TextMatrix(I, 1) = TmpRs1!Mst_GSL_ID
       MSHFlexGrid2.TextMatrix(I, 2) = TmpRs1!Locationname
       MSHFlexGrid2.TextMatrix(I, 3) = TmpRs1!CMPName
       MSHFlexGrid2.TextMatrix(I, 4) = TmpRs1!GSL
       MSHFlexGrid2.TextMatrix(I, 5) = TmpRs1!Commodity
       MSHFlexGrid2.TextMatrix(I, 6) = TmpRs1!ClientName
       MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!KRNo), "", TmpRs1!KRNo)
       MSHFlexGrid2.TextMatrix(I, 8) = Format(TmpRs1!KRDate, "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!ISINID), "", TmpRs1!ISINID)
       MSHFlexGrid2.TextMatrix(I, 10) = Format(TmpRs1!DRDate, "dd-MMM-yyyy")
       MSHFlexGrid2.TextMatrix(I, 11) = TmpRs1!tFlag 'TmpRs1!BalanceBags
       MSHFlexGrid2.TextMatrix(I, 12) = TmpRs1!Status 'TmpRs1!BalanceBags
       MSHFlexGrid2.TextMatrix(I, 13) = TmpRs1!BalanceWeight
       MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs1!NoOfDays), "", TmpRs1!NoOfDays)
       MSHFlexGrid2.TextMatrix(I, 15) = IIf(IsNull(TmpRs1!Rate), "", TmpRs1!Rate)
       MSHFlexGrid2.TextMatrix(I, 16) = IIf(IsNull(TmpRs1!InvoiceAmount), "", TmpRs1!InvoiceAmount)
       MSHFlexGrid2.TextMatrix(I, 17) = IIf(IsNull(TmpRs1!CommtrackIncome), "", TmpRs1!CommtrackIncome)
       
       If LCase(MSHFlexGrid2.TextMatrix(I, 11)) = "o" Then
          MSHFlexGrid2.TextMatrix(I, 11) = "Op. Balance"
       ElseIf LCase(MSHFlexGrid2.TextMatrix(I, 11)) = "d" Then
          MSHFlexGrid2.TextMatrix(I, 11) = "Deposit"
       Else
          MSHFlexGrid2.TextMatrix(I, 11) = "Withdrawal"
       End If
       
       TmpRs1.MoveNext
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   Next
End If

MsgBox "Done!!!"
ProgressBar1.Value = 0
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame2.Enabled = False
   Exit Sub
End If
Call CmdClear_Click
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CommTrackvsGSLIncome.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CommTrackvsGSLIncome.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Commtrack Lots" '"CompTrack VS GSL Income"
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Store Man GSL Detail"
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid2.Rows
For I = 0 To MSHFlexGrid2.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid2.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next

oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

