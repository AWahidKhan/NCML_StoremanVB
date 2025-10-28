VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_UnApprovedCRorTDS 
   Caption         =   "Unapproved CR or TDS"
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
      Left            =   0
      TabIndex        =   19
      Top             =   10680
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.Frame Frame1 
         Height          =   9615
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   14655
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
            Left            =   6480
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   487
            Width           =   3495
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
            Left            =   3000
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   487
            Width           =   3375
         End
         Begin VB.CommandButton CmdGo 
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
            Left            =   10800
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   480
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
            Left            =   12015
            TabIndex        =   6
            Top             =   480
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
            Height          =   360
            Left            =   13230
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   480
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
            Format          =   111542273
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   8145
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   900
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   14367
            _Version        =   393216
            FixedCols       =   0
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
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   375
            Left            =   1530
            TabIndex        =   2
            Top             =   480
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
            Format          =   111542273
            CurrentDate     =   36494
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   9120
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
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
            Left            =   7800
            TabIndex        =   21
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label6 
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
            Left            =   4155
            TabIndex        =   20
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label24 
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
            Left            =   13080
            TabIndex        =   17
            Top             =   5760
            Width           =   90
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   10680
            TabIndex        =   16
            Top             =   5760
            Width           =   2325
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Following Guards are not linked to any godowns and hence attendance can not be marked."
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
            Left            =   120
            TabIndex        =   15
            Top             =   5760
            Width           =   9360
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
            Left            =   1830
            TabIndex        =   14
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   10815
            TabIndex        =   13
            Top             =   187
            Width           =   2325
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
            Left            =   13215
            TabIndex        =   12
            Top             =   187
            Width           =   90
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   315
            TabIndex        =   11
            Top             =   180
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "FrmRpt_UnApprovedCRorTDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID As Variant

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" where " & Gen_WcState & "")
Else
   Call TableMst_State
End If

CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
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
If tmp <> CmbStateID.Text Then
   CmbLocation.Text = ""
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection!!!"
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
   Call TableMst_Location(" Where StateID= " & mStateID & " And " & Gen_WcLocation)
Else
   Call TableMst_Location(" Where StateID= " & mStateID)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocation.AddItem RsMst_Location!Locationname
   RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()

If CmbLocation.Text = "" Then
   Exit Sub
End If

' Code for clearing Godown combox in case of change

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Selection!!! "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub


Private Sub CmdClear_Click()
Call Grid_Head
CmbStateID.Text = ""
CmbLocation.Text = ""
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_UnApprovedCRorTDS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_UnApprovedCRorTDS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
If TxtFromDate.Value > Date Then
   MsgBox "Future Date not Allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
If TxtToDate.Value > Date Then
   MsgBox "Future Date not Allowed!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date should not greater than To Date!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
Call Grid_Head
MsgBox "Wait till Next Message!!!"

If CmbLocation.Text <> "" Then
   wc = " and ML.LocationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = " and ML. " & Gen_WcLocation
End If

If CmbStateID.Text <> "" Then
   wc = wc & " and MS.StateID = " & mStateID
ElseIf Gen_WcState <> "" Then
   wc = wc & " and MS. " & Gen_WcState
End If

tmp = "select mcl.navcode, mcl.rnavcode, mcl.clientname, crc.NewCRID crtdsid, crc.sm_prefix, cr.crdate , sum(crc.recamount) amount, sum(crc.rectds) rectds, 0 as recotherded from txn_newcashreceipt cr"
tmp = tmp & " inner join Txn_NewCashReceiptTIDetail crc on crc.NewCRID = cr.NewCRID and crc.sm_prefix = cr.sm_prefix"
tmp = tmp & " inner join mst_client mcl on mcl.ClientID = crc.ClientID"
tmp = tmp & " left join mst_location ml on ml.locationid = cr.locationid"
tmp = tmp & " left join mst_state ms on ms.stateid = ml.stateid"
tmp = tmp & " Where crc.recamount+crc.rectds > 0 And isnull(cr.batchid, 0) = 0 and cr.crdate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
tmp = tmp & " and isnull(cr.rejectflag, 'N') = 'N'"
tmp = tmp & wc
tmp = tmp & " group by mcl.navcode, mcl.rnavcode, mcl.clientname, crc.NewCRID , crc.sm_prefix, cr.crdate"
tmp = tmp & " Union All"
tmp = tmp & " select isnull(mcl.navcode, mdp.navcode) navcode, mcl.rnavcode, isnull( mcl.clientname, mdp.depositoryname) clientname,  crc.CRTDSAdjustmentID , crc.sm_prefix, cr.adjdate crdate, 0 as amount, sum(crc.rectds) rectds, sum(crc.recotherded) recotherded from Txn_CRTDSAdjustment cr"
tmp = tmp & " left join Txn_CRTDSAdjustmentDetail crc on crc.CRTDSAdjustmentID = cr.CRTDSAdjustmentID and crc.sm_prefix = cr.sm_prefix"
tmp = tmp & " left join mst_client mcl on mcl.ClientID = cr.ClientID and cr.showdp = 0"
tmp = tmp & " left join mst_depository mdp on mdp.depositoryid = cr.DepositoryID and cr.showdp = 1"
tmp = tmp & " left join mst_location ml on ml.locationid = cr.locationid"
tmp = tmp & " left join mst_state ms on ms.stateid = ml.stateid"
tmp = tmp & " Where crc.rectds + crc.recotherded > 0 And isnull(cr.batchid, 0) = 0 and cr.adjdate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "' and cr.flag <> 'R'"
tmp = tmp & wc
tmp = tmp & " group by mcl.navcode, mdp.navcode, mcl.rnavcode, mcl.clientname, mdp.depositoryname, crc.CRTDSAdjustmentID , crc.sm_prefix, cr.adjdate"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Label22.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      ProgressBar1.Max = .Rows
      For I = 2 To TmpRs.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!crtdsid), "", TmpRs!crtdsid)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!CRDate), "", Format(TmpRs!CRDate, "dd-MMM-yyyy"))
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!Amount), "0.00", Format(TmpRs!Amount, "#0.00"))
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!RecTDS), "0.00", Format(TmpRs!RecTDS, "#0.00"))
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs!RecOtherDed), "0.00", Format(TmpRs!RecOtherDed, "#0.00"))
         .TextMatrix(I, 9) = Format(Val(.TextMatrix(I, 6)) + Val(.TextMatrix(I, 7)) + Val(.TextMatrix(I, 8)), "#0.00")
         .TextMatrix(0, 6) = Format(Val(.TextMatrix(I, 6)) + Val(.TextMatrix(0, 6)), "#0.00")
         .TextMatrix(0, 7) = Format(Val(.TextMatrix(I, 7)) + Val(.TextMatrix(0, 7)), "#0.00")
         .TextMatrix(0, 8) = Format(Val(.TextMatrix(I, 8)) + Val(.TextMatrix(0, 8)), "#0.00")
         .TextMatrix(0, 9) = Format(Val(.TextMatrix(I, 9)) + Val(.TextMatrix(0, 9)), "#0.00")
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
   MsgBox "No Record Found!!!"
End If
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
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
HScroll1.Max = Me.ScaleWidth - Frame1.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True
If Frame1.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

End Sub

Private Sub VScroll1_Scroll()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

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

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call CmdClear_Click
End Sub

Private Sub Grid_Head()
Label22.Caption = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 10
    .FixedCols = 1
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    
    .TextMatrix(1, 0) = "NAV Code"
    .TextMatrix(1, 1) = "RNAV Code"
    .TextMatrix(1, 2) = "Client Name"
    .TextMatrix(1, 3) = "CR No/TDS Adjustment ID"
    .TextMatrix(1, 4) = "SM_Prefix"
    .TextMatrix(1, 5) = "CR Date"
    .TextMatrix(1, 6) = "CR Amount"
    .TextMatrix(1, 7) = "TDS"
    .TextMatrix(1, 8) = "Other Deduction"
    .TextMatrix(1, 9) = "Total"
         
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 3200
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 1400
    .ColWidth(8) = 1400
    .ColWidth(9) = 1400
    
    .RowHeight(1) = 600
End With
End Sub
