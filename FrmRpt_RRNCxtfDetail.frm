VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_RRNCxtfDetail 
   Caption         =   "RRN CITF Detail"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   45
      TabIndex        =   13
      Top             =   9255
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   45
      TabIndex        =   10
      Top             =   -30
      Width           =   15105
      Begin VB.ComboBox CmbSClientName 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   825
         Width           =   10155
      End
      Begin VB.ComboBox CmbCommodity 
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
         Left            =   8895
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   6135
      End
      Begin VB.TextBox TxtTruckNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1260
         Width           =   4335
      End
      Begin VB.TextBox TxtRRNNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   12765
         TabIndex        =   5
         Top             =   825
         Width           =   2250
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   3900
         TabIndex        =   1
         Top             =   375
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
         Format          =   111607809
         CurrentDate     =   36494
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   255
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
         Left            =   13170
         TabIndex        =   8
         Top             =   1320
         Width           =   885
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
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
         Left            =   12285
         TabIndex        =   7
         Top             =   1320
         Width           =   885
      End
      Begin VB.ComboBox CmbSLocationID 
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
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   390
         Width           =   3585
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
         Height          =   350
         Left            =   14055
         TabIndex        =   9
         Top             =   1320
         Width           =   885
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   390
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
         Format          =   111607809
         CurrentDate     =   36494
      End
      Begin VB.Label Label7 
         Caption         =   "Client Name"
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
         TabIndex        =   22
         Top             =   855
         Width           =   1215
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
         Left            =   9555
         TabIndex        =   21
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label5 
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
         Left            =   7095
         TabIndex        =   20
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label Label6 
         Caption         =   "Truck No"
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
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "RRN No"
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
         Left            =   11880
         TabIndex        =   18
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   3000
         TabIndex        =   17
         Top             =   457
         Width           =   735
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
         Left            =   480
         TabIndex        =   15
         Top             =   457
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity"
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
         Left            =   11317
         TabIndex        =   14
         Top             =   135
         Width           =   1290
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
         Left            =   6667
         TabIndex        =   11
         Top             =   150
         Width           =   810
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7380
      Left            =   45
      TabIndex        =   12
      Top             =   1800
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   13018
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
Attribute VB_Name = "FrmRpt_RRNCxtfDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSLocationID, mSExchangeTypeID, mCommodityID As Variant

Private Sub Check1_Click()
If Check1.Value = 1 Then
   TxtFromDate.Enabled = True
   TxtToDate.Enabled = True
Else
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
End If
End Sub

Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodity.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID

End Sub

Private Sub CmbSClientName_GotFocus()
tmp1 = CmbSClientName.Text

'Call TableMst_Client("Where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
tmp = "Select clientName + '~' + ClientID + '~' + Convert(Varchar,ClientIDRow) 'ClientName' From Mst_Client Where ExchangeTypeId = 2 "
Call Tmp1Table

CmbSClientName.Clear
If TmpRs1.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbSClientName.AddItem TmpRs1!ClientName
    TmpRs1.MoveNext
Next
CmbSClientName.Text = tmp1
End Sub
Private Sub CmbSClientName_LostFocus()
If CmbSClientName.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "ClientName = '" & CmbSClientName.Text & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbSClientName.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmbSLocationID.Text = ""
CmbCommodity.Text = ""
TxtRRNNo.Text = ""
TxtTruckNo.Text = ""
ProgressBar1.Value = 0
Check1.Value = 0
TxtFromDate.Enabled = False
TxtToDate.Enabled = False
CmdExcel.Enabled = False
CmbSClientName.Text = ""
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_RRNCITFDetail.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_RRNCITFDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        ElseIf C = 8 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGo_Click()
Dim wc As Variant
LblTotalRecord.Caption = "-"

CmdExcel.Enabled = False

ProgressBar1.Value = 0
Call Grid_Head
MsgBox "Wait till next message !!!"

wc = ""

If Check1.Value = 1 Then
   If TxtFromDate.Value > Date Then
      MsgBox "Future Date not Allowed!!!"
      TxtFromDate.SetFocus
      Exit Sub
   End If
   If TxtToDate.Value < TxtFromDate.Value Then
      MsgBox "To date must be always greater than from date!!!"
      TxtToDate.SetFocus
      Exit Sub
   End If
   wc = " And a.cxtfDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
End If
    
tmp = " Select l.locationname,c.commodity,a.CxTF_DetailsID,a.cxtfDate,r.RRNNo,r.rrndate,mc.clientName + '~' + mc.ClientIDRow + '~' + Convert(Varchar,mc.ClientID) 'RRNClient' ,b.NoOfBags,b.DematWeight,b.TruckNo,MG.ISINID,a.Flag "
tmp = tmp & " From Txn_CxTF_Details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on a.CxTF_DetailsID =  b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix"
tmp = tmp & " inner join mst_commodity c on a.CommodityID = c.CommodityID"
tmp = tmp & " inner join mst_location l on  a.SM_Prefix = l.SM_Prefix"
tmp = tmp & " inner join mst_rrn r on  b.WHRNo = r.RRNID"
tmp = tmp & " inner join mst_Client MC On r.ClientID = mc.ClientID"
tmp = tmp & " inner join mst_GSL MG On b.Mst_GSL_ID = MG.Mst_GSL_ID and b.SM_Prefix = MG.SM_Prefix "
tmp = tmp & " where a.TxnType = 'W' and a.exchangetypeid = 2 "

If CmbSLocationID.Text <> "" Then
   wc = wc & " and l.LocationName = '" & CmbSLocationID.Text & "' "
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And l." & Gen_WcLocation
   End If
End If

If CmbCommodity.Text <> "" Then
   wc = wc & " And c.Commodity = '" & CmbCommodity.Text & "' "
End If

If Trim(TxtRRNNo.Text) <> "" Then
   wc = wc & " And r.RRNNo= '" & Trim(TxtRRNNo.Text) & "'"
End If

If Trim(TxtTruckNo.Text) <> "" Then
   wc = wc & " And b.TruckNo like '%" & Trim(TxtTruckNo.Text) & "%' "
End If

If CmbSClientName.Text <> "" Then
   wc = wc & " And mc.clientName + '~' + mc.ClientIDRow + '~' + Convert(Varchar,mc.ClientID) = '" & CmbSClientName.Text & "' "
End If

tmp1 = " Order By l.locationname,c.commodity,a.CxTF_DetailsID,a.cxtfDate"

tmp = tmp & " " & wc & " " & tmp1

Call TmpTable

If TmpRs.RecordCount > 0 Then
   CmdExcel.Enabled = True
   ProgressBar1.Max = TmpRs.RecordCount + 2

   For mRow = 1 To TmpRs.RecordCount
       For mCol = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(mRow, mCol) = IIf(IsNull(TmpRs.Fields(mCol)), "", TmpRs.Fields(mCol))
       Next
       MSHFlexGrid1.TextMatrix(mRow, 3) = Format(MSHFlexGrid1.TextMatrix(mRow, 3), "dd-MMM-yyyy")
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   
   ProgressBar1.Value = ProgressBar1.Max
   LblTotalRecord.Caption = TmpRs.RecordCount
Else
   MsgBox " No Record Found "
   Exit Sub
End If
MsgBox "Done !!!"
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If

Call CmdClear_Click
End Sub
Private Sub Grid_Head()

LblTotalRecord.Caption = ""
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 12 '11 '10 '9
    .RowHeight(0) = 400
    
    .AllowUserResizing = flexResizeColumns
    .Font.Size = 10
    .WordWrap = True
    
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "Commodity"
    .TextMatrix(0, 2) = "CITF No"
    .TextMatrix(0, 3) = "CITF Date"
    .TextMatrix(0, 4) = "RRN NO"
    .TextMatrix(0, 5) = "RRN Date"
    .TextMatrix(0, 6) = "RRN ClientName ~ ID ~ RowID "
    .TextMatrix(0, 7) = "No of Bags"
    .TextMatrix(0, 8) = "Weight"
    .TextMatrix(0, 9) = "Truck No"
    .TextMatrix(0, 10) = "CMSE Lot No"
    .TextMatrix(0, 11) = "CITF Print Flag"

    .ColWidth(0) = 2500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 2000
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    .RowHeight(0) = 600

End With
End Sub

Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbSLocationID.AddItem RsMst_Location!Locationname
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
