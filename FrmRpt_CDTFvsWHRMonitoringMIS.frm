VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_CDTFvsWHRMonitoringMIS 
   Caption         =   "CDTF vs WHR Monitoring MIS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   9840
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   15015
      TabIndex        =   0
      Top             =   315
      Width           =   255
   End
   Begin VB.PictureBox Frame4 
      Height          =   9540
      Left            =   45
      ScaleHeight     =   9480
      ScaleWidth      =   15135
      TabIndex        =   7
      Top             =   60
      Width           =   15195
      Begin VB.Frame Frame11 
         Height          =   825
         Left            =   45
         TabIndex        =   8
         Top             =   -30
         Width           =   14880
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
            Left            =   13335
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox ChkWHR 
            Caption         =   "Show WHR Delivered details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4185
            TabIndex        =   2
            Top             =   285
            Width           =   3015
         End
         Begin VB.CheckBox ChkMinus 
            Caption         =   "Show All (with and without Discrepancy)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            TabIndex        =   1
            Top             =   278
            Value           =   1  'Checked
            Width           =   3855
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
            Left            =   10905
            TabIndex        =   3
            Top             =   240
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
            Height          =   360
            Left            =   12120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
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
            Left            =   7560
            TabIndex        =   10
            Top             =   293
            Width           =   2280
         End
         Begin VB.Label Label22 
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
            Left            =   10005
            TabIndex        =   9
            Top             =   300
            Width           =   90
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8265
         Left            =   45
         TabIndex        =   11
         Top             =   810
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   14579
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   45
         TabIndex        =   12
         Top             =   9150
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_CDTFvsWHRMonitoringMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClear_Click()
ChkMinus.Value = vbUnchecked
ChkWHR.Value = vbUnchecked
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CDTFvsWHRMonitoringMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CDTFvsWHRMonitoringMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   For c = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
   Next
   mRow = mRow + 1
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Call Grid_Head
MsgBox ("Wait till Next Message!!!")

tmp = " select"
tmp = tmp & " ml.locationname, mcom.commodity, mcl.clientname, mcl.clientidrow, a.sm_prefix, mcom.commodityid, sum(b.NoOfBags) CxTFBags, sum(b.GrossWeight) CxTFQty,"
tmp = tmp & " (select sum(WHRNoOfBags) from txn_whr_detail where clientidrow = mcl.clientidrow and a.sm_prefix = sm_prefix and commodityid = mcom.commodityid ) WHRBags,"
tmp = tmp & " (select sum(WHRQuantity) from txn_whr_detail where clientidrow = mcl.clientidrow and a.sm_prefix = sm_prefix and commodityid = mcom.commodityid ) WHRQty,"
If ChkWHR.Value = vbChecked Then
   tmp = tmp & " (select sum(DeliveredNoOfBags) from txn_whr_detail where clientidrow = mcl.clientidrow and a.sm_prefix = sm_prefix and commodityid = mcom.commodityid ) WHRDeliveredBags,"
   tmp = tmp & " (select sum(DeliveredQuantity) from txn_whr_detail where clientidrow = mcl.clientidrow and a.sm_prefix = sm_prefix and commodityid = mcom.commodityid ) WHRDeliveredQty,"
End If
tmp = tmp & " (select Isnull(sum(g.DepositeBags)-sum(g.WHRBags),0) from mst_gsl g where g.clientidrow = mcl.clientidrow and a.sm_prefix = g.sm_prefix and g.commodityid = mcom.commodityid And g.exchangetypeid = 1 and g.flag <> 'Z' and g.DepositeBags<>g.WHRBags) 'YetToBags',"
tmp = tmp & " (select Isnull(sum(g.DepositeWeight)-sum(g.WHRWeight),0) from mst_gsl g where g.clientidrow = mcl.clientidrow and a.sm_prefix = g.sm_prefix and g.commodityid = mcom.commodityid And g.exchangetypeid = 1 and g.flag <> 'Z' and g.DepositeBags<>g.WHRBags) 'YetToWt'"
tmp = tmp & " from txn_cxtf_gsl b"
tmp = tmp & " inner join txn_cxtf_details a on a.sm_prefix = b.SM_Prefix and a.cxtfno = b.cxtfno and a.txntype = b.txntype"
tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid  =  a.commodityid"
tmp = tmp & " inner join mst_cmp mcmp on mcmp.cmpid = a.cmpid"
tmp = tmp & " inner join mst_location ml on ml.locationid  = mcmp.locationid"
tmp = tmp & " inner join mst_client mcl on mcl.clientidrow = b.clientidrow"
tmp = tmp & " where a.txntype = 'D'"
tmp = tmp & " group by ml.locationname, mcom.commodity, mcl.clientname, mcl.clientidrow, a.sm_prefix, mcom.commodityid"
'If ChkMinus.Value = vbChecked Then
'   tmp = tmp & " Having (select sum(WHRNoOfBags) from txn_whr_detail where clientidrow = mcl.clientidrow and a.sm_prefix = sm_prefix and commodityid = mcom.commodityid ) > sum(b.NoOfBags)"
'End If
tmp = tmp & " order by ml.locationname, mcl.clientname, mcom.commodity"
Call TmpTable
Dim x As Double
If TmpRs.RecordCount > 0 Then
   CmdExcel.Enabled = True
   ProgressBar1.Max = TmpRs.RecordCount + 2
   x = 1
   For I = 1 To TmpRs.RecordCount
      If ChkMinus.Value = vbUnchecked Then
         t = IIf(IsNull(TmpRs!CxTFBags), 0, TmpRs!CxTFBags) - (IIf(IsNull(TmpRs!WHRBags), 0, TmpRs!WHRBags) + IIf(IsNull(TmpRs!YetToBags), 0, TmpRs!YetToBags))
         If t >= 0 Then
            GoTo nextrow
         End If
      End If
      MSHFlexGrid1.TextMatrix(x, 0) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
      MSHFlexGrid1.TextMatrix(x, 1) = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
      MSHFlexGrid1.TextMatrix(x, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
      MSHFlexGrid1.TextMatrix(x, 3) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
      MSHFlexGrid1.TextMatrix(x, 4) = IIf(IsNull(TmpRs!CxTFBags), 0, TmpRs!CxTFBags)
      MSHFlexGrid1.TextMatrix(x, 5) = IIf(IsNull(TmpRs!CxTFQty), 0, TmpRs!CxTFQty)
      MSHFlexGrid1.TextMatrix(x, 6) = IIf(IsNull(TmpRs!WHRBags), 0, TmpRs!WHRBags)
      MSHFlexGrid1.TextMatrix(x, 7) = IIf(IsNull(TmpRs!WHRQty), 0, TmpRs!WHRQty)
      MSHFlexGrid1.TextMatrix(x, 8) = IIf(IsNull(TmpRs!CxTFBags), 0, TmpRs!CxTFBags) - IIf(IsNull(TmpRs!WHRBags), 0, TmpRs!WHRBags)
      MSHFlexGrid1.TextMatrix(x, 9) = IIf(IsNull(TmpRs!CxTFQty), 0, TmpRs!CxTFQty) - IIf(IsNull(TmpRs!WHRQty), 0, TmpRs!WHRQty)
      If ChkWHR.Value = vbChecked Then
         MSHFlexGrid1.TextMatrix(x, 10) = IIf(IsNull(TmpRs!WHRDeliveredBags), 0, TmpRs!WHRDeliveredBags)
         MSHFlexGrid1.TextMatrix(x, 11) = IIf(IsNull(TmpRs!WHRDeliveredQty), 0, TmpRs!WHRDeliveredQty)
         MSHFlexGrid1.TextMatrix(x, 12) = IIf(IsNull(TmpRs!YetToBags), 0, TmpRs!YetToBags)
         MSHFlexGrid1.TextMatrix(x, 13) = IIf(IsNull(TmpRs!YetToWt), 0, TmpRs!YetToWt)
         MSHFlexGrid1.TextMatrix(x, 14) = Val(MSHFlexGrid1.TextMatrix(x, 8)) - Val(MSHFlexGrid1.TextMatrix(x, 10)) - Val(MSHFlexGrid1.TextMatrix(x, 12))
         MSHFlexGrid1.TextMatrix(x, 15) = Val(MSHFlexGrid1.TextMatrix(x, 9)) - Val(MSHFlexGrid1.TextMatrix(x, 11)) - Val(MSHFlexGrid1.TextMatrix(x, 13))
         MSHFlexGrid1.TextMatrix(x, 14) = Format(MSHFlexGrid1.TextMatrix(x, 14), "########0")
         MSHFlexGrid1.TextMatrix(x, 15) = Format(MSHFlexGrid1.TextMatrix(x, 15), "########0.000")
      Else
         MSHFlexGrid1.TextMatrix(x, 10) = IIf(IsNull(TmpRs!YetToBags), 0, TmpRs!YetToBags)
         MSHFlexGrid1.TextMatrix(x, 11) = IIf(IsNull(TmpRs!YetToWt), 0, TmpRs!YetToWt)
         MSHFlexGrid1.TextMatrix(x, 12) = Val(MSHFlexGrid1.TextMatrix(x, 8)) - Val(MSHFlexGrid1.TextMatrix(x, 10))
         MSHFlexGrid1.TextMatrix(x, 13) = Val(MSHFlexGrid1.TextMatrix(x, 9)) - Val(MSHFlexGrid1.TextMatrix(x, 11))
         MSHFlexGrid1.TextMatrix(x, 12) = Format(MSHFlexGrid1.TextMatrix(x, 12), "########0")
         MSHFlexGrid1.TextMatrix(x, 13) = Format(MSHFlexGrid1.TextMatrix(x, 13), "########0.000")
      End If
      x = x + 1
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
nextrow:
      ProgressBar1.Value = ProgressBar1.Value + 1
      TmpRs.MoveNext
   Next
   Label22.Caption = MSHFlexGrid1.Rows - 2
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
   MsgBox "No Record found!!!"
End If
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame11.Enabled = False
   Exit Sub
End If
Call CmdClear_Click
Frame11.Enabled = True
End Sub
Private Sub Grid_Head()
Label22.Caption = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .WordWrap = True
   .Rows = 2
   .Cols = 14
   
   .TextMatrix(0, 0) = "Location"
   .TextMatrix(0, 1) = "Client Row ID"
   .TextMatrix(0, 2) = "Client Name"
   .TextMatrix(0, 3) = "Commodity"
   .TextMatrix(0, 4) = "CDTF Bags"
   .TextMatrix(0, 5) = "CDTF Qty"
   .TextMatrix(0, 6) = "WHR Bags"
   .TextMatrix(0, 7) = "WHR Qty"
   .TextMatrix(0, 8) = "Excess/short WHr -   Bags"
   .TextMatrix(0, 9) = "Excess/short WHr -   Qty"
   If ChkWHR.Value = vbChecked Then
      .Cols = 16
      .TextMatrix(0, 10) = "Delivered Bags Against WHR"
      .TextMatrix(0, 11) = "Delivered Qty Against WHR"
      .TextMatrix(0, 12) = "Yet to be generated Bags"
      .TextMatrix(0, 13) = "Yet to be generated Qty"
      .TextMatrix(0, 14) = "Difference in Bags"
      .TextMatrix(0, 15) = "Difference in  Qty"
      
      .ColWidth(9) = 1200
      .ColWidth(10) = 1200
   Else
      .TextMatrix(0, 10) = "Yet to be generated Bags"
      .TextMatrix(0, 11) = "Yet to be generated Qty"
      .TextMatrix(0, 12) = "Difference in Bags"
      .TextMatrix(0, 13) = "Difference in  Qty"
   End If
   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1200
   .ColWidth(2) = 3500
   .ColWidth(3) = 1800
   .ColWidth(4) = 1200
   .ColWidth(5) = 1200
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 1200
   .ColWidth(9) = 1200
   
   .RowHeight(0) = 800
End With

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

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub VScroll1_Change()

   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * x) / 1000
   Frame4.Top = -L

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()

   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * x) / 1000
   Frame4.Top = -L

'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * x) / 1000
   Frame4.Left = -L
End Sub

Private Sub HScroll1_Scroll()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * x) / 1000
   Frame4.Left = -L
End Sub
