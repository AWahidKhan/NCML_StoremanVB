VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CITFBOE 
   Caption         =   "CITF BOE"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   15015
      TabIndex        =   22
      Top             =   330
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   180
      TabIndex        =   21
      Top             =   9855
      Width           =   10935
   End
   Begin VB.PictureBox Frame4 
      Height          =   9780
      Left            =   0
      ScaleHeight     =   9720
      ScaleWidth      =   15255
      TabIndex        =   0
      Top             =   15
      Width           =   15315
      Begin VB.Frame Frame11 
         Height          =   1380
         Left            =   45
         TabIndex        =   1
         Top             =   -30
         Width           =   15000
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
            Left            =   3150
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   390
            Width           =   3750
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
            Left            =   13680
            TabIndex        =   7
            Top             =   930
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
            Left            =   11250
            TabIndex        =   6
            Top             =   930
            Width           =   1215
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
            Left            =   12465
            TabIndex        =   5
            Top             =   930
            Width           =   1215
         End
         Begin VB.TextBox TxtCITFNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   10680
            MaxLength       =   10
            TabIndex        =   4
            Top             =   390
            Width           =   2085
         End
         Begin VB.TextBox TxtCRNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   12825
            MaxLength       =   10
            TabIndex        =   3
            Top             =   390
            Width           =   2085
         End
         Begin VB.ComboBox CmbExchangeTypeID 
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
            Left            =   6960
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   390
            Width           =   3630
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   330
            Left            =   120
            TabIndex        =   9
            Top             =   390
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   61079553
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   330
            Left            =   1635
            TabIndex        =   10
            Top             =   390
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   61079553
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin VB.Label Label2 
            Caption         =   "To CITF Date"
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
            Left            =   1695
            TabIndex        =   18
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "From CITF  Date"
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
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1575
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
            Height          =   255
            Left            =   4598
            TabIndex        =   16
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "CITF No"
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
            Left            =   11415
            TabIndex        =   15
            Top             =   120
            Width           =   750
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
            Left            =   9660
            TabIndex        =   14
            Top             =   1020
            Width           =   90
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
            Left            =   7335
            TabIndex        =   13
            Top             =   1020
            Width           =   2280
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cash Receipt No"
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
            Left            =   13095
            TabIndex        =   12
            Top             =   120
            Width           =   1530
         End
         Begin VB.Label Label6 
            Caption         =   "ExchangeType"
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
            Left            =   8040
            TabIndex        =   11
            Top             =   120
            Width           =   1470
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7545
         Left            =   45
         TabIndex        =   19
         Top             =   1410
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   13309
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   9150
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_CITFBOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mExchangeTypeID As Variant
Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location("")
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocation")
End If
End Sub

Private Sub CmbLocation_LostFocus()

If CmbLocation.Text = "" Then
   mLocationID = 0
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

Private Sub CmdClear_Click()
Call Grid_Head
Label22.Caption = 0
TxtToDate.Value = Date
TxtFromDate.Value = Date - 30
CmbLocation.Text = ""
TxtCITFNo.Text = ""
TxtCRNo = ""
CmbExchangeTypeID.Text = ""
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CITFBOE.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CITFBOE.xls")
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

Private Sub CmdGetReport_Click()

ProgressBar1.Value = 0
If TxtFromDate > TxtToDate Then
   MsgBox "FromDate Should be Less Than Todate."
   TxtFromDate.SetFocus
   Exit Sub
End If

Label22.Caption = ""


'tmp = "Select LocationName,CxTFNo,CxTFDate,ClientName,Commodity,WithdrawalBags,WithdrawalQty, BOEAmount"
'tmp = tmp & " TCRD.CRID, CRDate, InstrumentAmount, TCRD.Remarks"
'tmp = tmp & " From Txn_CxTF_BOE  TC"
'tmp = tmp & " Inner Join Mst_Commodity MC on TC.CommodityID=MC.CommodityID"
'tmp = tmp & " Inner Join Mst_Client MCl on TC.ClientIDRow=MCl.ClientIDRow"
'tmp = tmp & " Inner Join Mst_Location ML on TC.LocationID=ML.LocationID"

tmp = " Select ML.LocationName,TC.CxTFNo,TC.CxTFDate,MCl.ClientName,MC.Commodity, ME.ExchangeType, TC.WithdrawalBags,TC.WithdrawalQty, TC.BOEAmount, TCRD.CRID,"
tmp = tmp & " TCR.CRDate, TCR.InstrumentAmount,TCRD.Remarks"
tmp = tmp & " From Txn_CxTF_BOE  TC"
tmp = tmp & " LEFT Join Txn_CITFCRDetail TCRD on TCRD.CITFNo=TC.CxTFNo AND TCRD.SM_Prefix=TC.SM_Prefix"
tmp = tmp & " LEFT Join Txn_CashReceipt TCR on TCR.CRID=TCRD.CRID AND TCR.SM_Prefix=TCRD.SM_Prefix"
tmp = tmp & " Inner Join Mst_Commodity MC on TC.CommodityID=MC.CommodityID"
tmp = tmp & " Inner Join Mst_Client MCl on TC.ClientIDRow=MCl.ClientIDRow"
tmp = tmp & " Inner Join Mst_Location ML on TC.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME on ME.ExchangeTypeID=TC.ExchangeTypeID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TC.CxTFDate >= #" & Format(TxtFromDate, Gen_DatFormat) & "# And TC.CxTFDate <= #" & Format(TxtToDate, Gen_DatFormat) & "#"
Else
   tmp = tmp & " Where TC.CxTFDate >= '" & Format(TxtFromDate, Gen_DatFormat) & "' And TC.CxTFDate <= '" & Format(TxtToDate, Gen_DatFormat) & "'"
End If

If CmbLocation.Text <> "" Then
   tmp = tmp & " And ML.LocationName='" & CmbLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

If CmbExchangeTypeID.Text <> "" Then
   tmp = tmp & " And ME.ExchangeType ='" & CmbExchangeTypeID.Text & "'"
End If


If TxtCITFNo.Text <> "" Then
   tmp = tmp & " And TC.CxTFNo=" & Val(TxtCITFNo.Text) & ""
End If

If TxtCRNo.Text <> "" Then
   tmp = tmp & " And TCRD.CRID=" & Val(TxtCRNo.Text) & ""
End If

tmp = tmp & "  order by ML.LocationName,TC.CxTFNo,TC.CxTFDate"

Call TmpTable



Call dispdata

MsgBox ("Done!!!")

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
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
'   CmdGetReport.Enabled = False
'   CmdExcel.Enabled = False
   Frame11.Enabled = False
   Exit Sub
End If
Call Grid_Head
Frame11.Enabled = True
TxtFromDate.Value = Date - 365
TxtToDate.Value = Date
CmdExcel.Enabled = False
ProgressBar1.Value = 0
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13

MSHFlexGrid1.TextMatrix(0, 0) = "Location"
MSHFlexGrid1.TextMatrix(0, 1) = "CITF No"
MSHFlexGrid1.TextMatrix(0, 2) = "CITF Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 5) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 6) = "Withdrawal Bags"
MSHFlexGrid1.TextMatrix(0, 7) = "Withdrawal Qty"
MSHFlexGrid1.TextMatrix(0, 8) = "BOE Amount"
MSHFlexGrid1.TextMatrix(0, 9) = "Cash Receipt No"
MSHFlexGrid1.TextMatrix(0, 10) = "Cash Receipt Date"
MSHFlexGrid1.TextMatrix(0, 11) = "Cash Receipt Amount"
MSHFlexGrid1.TextMatrix(0, 12) = "Remarks"


MSHFlexGrid1.ColWidth(0) = 2200
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 3600
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 2500
 
MSHFlexGrid1.RowHeight(0) = 450

End Sub

Public Sub dispdata()
Dim x As Double
Call Grid_Head
MsgBox ("Wait till Next Message")
x = 1
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(I, 2) = Format(MSHFlexGrid1.TextMatrix(I, 2), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 10) = Format(MSHFlexGrid1.TextMatrix(I, 10), "dd-mmm-yyyy")
       ProgressBar1.Value = ProgressBar1.Value + 1
       TmpRs.MoveNext
   Next
   Label22.Refresh
   Label22.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If
ProgressBar1.Value = ProgressBar1.Max

End Sub

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
    
End Sub

Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
 Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
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

