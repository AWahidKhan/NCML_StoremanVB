VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_StorageContractDetail 
   Caption         =   "Storage Contract Detail"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   8925
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   9360
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   7215
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12091
         _Version        =   393216
         FixedCols       =   0
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
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      Begin VB.OptionButton OptExpiredContract 
         Caption         =   "Expired Contract Having GSL Stock"
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
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
      Begin VB.OptionButton OptDetailRpt 
         Caption         =   "Detail Report"
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
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CheckBox ChkCurrDt 
         Appearance      =   0  'Flat
         Caption         =   "Apply Current Date Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   2625
      End
      Begin VB.TextBox TxtResevationQty 
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
         Left            =   13365
         TabIndex        =   10
         Top             =   765
         Width           =   1515
      End
      Begin VB.TextBox TxtContractID 
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
         Height          =   330
         Left            =   11685
         TabIndex        =   7
         Top             =   765
         Width           =   1640
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmRpt_StorageContractDetail.frx":0000
         Left            =   10230
         List            =   "FrmRpt_StorageContractDetail.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   750
         Width           =   1410
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   2955
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   750
         Width           =   2745
      End
      Begin VB.ComboBox CmbClient 
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
         Left            =   5715
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   750
         Width           =   4485
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
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   750
         Width           =   2745
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
         Left            =   13455
         TabIndex        =   12
         Top             =   1545
         Width           =   1335
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
         Left            =   12120
         TabIndex        =   11
         Top             =   1545
         Width           =   1335
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
         Left            =   10785
         TabIndex        =   9
         Top             =   1545
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Reservation Qty"
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
         Left            =   13365
         TabIndex        =   22
         Top             =   525
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "Contract ID"
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
         Left            =   12000
         TabIndex        =   21
         Top             =   525
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
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
         Left            =   10620
         TabIndex        =   20
         Top             =   510
         Width           =   645
      End
      Begin VB.Label LblCommodityID 
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
         Height          =   300
         Left            =   3780
         TabIndex        =   19
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Client"
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
         Left            =   7650
         TabIndex        =   18
         Top             =   525
         Width           =   600
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   1155
         TabIndex        =   17
         Top             =   510
         Width           =   780
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
         Left            =   5760
         TabIndex        =   14
         Top             =   1620
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
         Left            =   8070
         TabIndex        =   13
         Top             =   1620
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmRpt_StorageContractDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mClientIDRow, mCommodityID As Variant

Private Sub CmbClient_GotFocus()
tmp = CmbClient.Text
Call TableMst_Client(" Where ExchangeTypeId = 1 ")
CmbClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClient.Text = tmp
End Sub

Private Sub CmbClient_LostFocus()
If CmbClient.Text = "" Then
   mClientIDRow = Null
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> "active" And LCase(CmbStatus.Text) <> "deactive" Then
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbClient.Text = ""
CmbStatus.Text = ""
TxtContractID.Text = ""
TxtResevationQty.Text = ""
CmbCommodityID.Text = ""
LblTotalRecord.Caption = "-"
End Sub

Private Sub CmdExcel_Click()

If MSHFlexGrid1.Rows < 3 And MSHFlexGrid1.TextMatrix(1, 0) = "" Then MsgBox "No Record Found": Exit Sub

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FrmRpt_StorageContractDetail.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "FrmRpt_StorageContractDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1
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


Private Sub CmdGetReport_Click()

Call Grid_Head

MsgBox "Wait till next message!!!"

wc = ""

If CmbLocation.Text <> "" Then
   If wc = "" Then
      wc = " ML.LocationID =" & mLocationID
   Else
      wc = wc & " And ML.LocationID =" & mLocationID
   End If
Else
   If Gen_WcLocation <> "" Then
      If wc = "" Then
         wc = " ML." & Gen_WcLocation
      Else
         wc = wc & " And ML." & Gen_WcLocation
      End If
   End If
End If

If CmbCommodityID.Text <> "" Then
   If wc = "" Then
      wc = " MC.CommodityID =" & mCommodityID
   Else
      wc = wc & " And MC.CommodityID =" & mCommodityID
   End If
End If

If CmbClient.Text <> "" Then
   If wc = "" Then
      wc = " MCL.ClientID ='" & mClientIDRow & "'"
   Else
      wc = wc & " And MCL.ClientID ='" & mClientIDRow & "'"
   End If
End If

If CmbStatus.Text <> "" Then
   If wc = "" Then
      wc = " a.Flag ='" & Left(CmbStatus.Text, 1) & "'" 'IIf(CmbStatus.Text <> "Active", "D", "A") & "'"
   Else
      wc = wc & " And a.Flag ='" & Left(CmbStatus.Text, 1) & "'" 'a.Flag ='" & IIf(CmbStatus.Text <> "Active", "D", "A") & "'"
   End If
End If

If TxtContractID.Text <> "" Then
   If wc = "" Then
      wc = " a.ContractID =" & Val(TxtContractID)
   Else
      wc = wc & " And a.ContractID =" & Val(TxtContractID)
   End If
End If

If TxtResevationQty.Text <> "" Then
   If wc = "" Then
      wc = " b.BillingQty " & TxtResevationQty
   Else
      wc = wc & " And b.BillingQty " & TxtResevationQty
   End If
End If


If OptExpiredContract.Value = False Then

    If ChkCurrDt.Value = 1 Then
       If wc = "" Then
          wc = "'" & Format(Date, Gen_DatFormat) & "'" & " between b.FromDate And b.ToDate "
       Else
          wc = wc & " And " & "'" & Format(Date, Gen_DatFormat) & "'" & " between b.FromDate And b.ToDate "
       End If
    End If
    
    tmp = "Select a.ContractID, a.ContractName, ML.LocationName, MCL.ClientName +'-'+ MCL.ClientIDRow 'Client', MC.Commodity,"
    tmp = tmp & " a.BookedFrom 'ContractStartDate', a.ExtendedPeriod 'ContractEndDate', a.Remarks,"
    tmp = tmp & " b.FromDate 'PeriodFromDate', b.ToDate'PeriodToDate', MBC.BillingCycle, b.BillingQty 'ReservationQty',"
    tmp = tmp & " b.BillingRate 'Rate', b.Remarks 'DetailsRemarks', Case when upper(a.Flag) <>'A' then 'DeActive' else 'Active' end As Flag"
    tmp = tmp & " From Mst_StorageContract a"
    tmp = tmp & " Left Join  Mst_SContractBillingDetail b On a.ContractID = b.ContractID"
    tmp = tmp & " Left Join Mst_Location Ml on a.LocationID = Ml.LocationID"
    tmp = tmp & " Left  Join Mst_Commodity MC on a.CommodityID =MC.CommodityID"
    tmp = tmp & " Left Join Mst_BillingCycle MBC On b.BillingCycleID = MBC.BillingCycleID"
    tmp = tmp & " Left Join Mst_Client MCl on a.ClientID = MCl.ClientID"

    If wc <> "" Then
       tmp = tmp & " Where " & wc
    End If
    
    tmp = tmp & " Order By ML.LocationName,a.ContractID"

Else

    tmp = " Select s.statename,ML.LocationName,  a.ContractID,a.ContractName,MCL.ClientName,b.SCMFlag,a.BookedFrom 'StartDate',a.ExtendedPeriod,a.Flag,SUM(b.balancebags) 'BalBags',SUM(b.balanceweight) 'BalWeight'"
    tmp = tmp & " from Mst_GSL b"
    tmp = tmp & " inner join Mst_StorageContract a on b.NContractID = a.ContractID"
    tmp = tmp & " inner join Mst_Client MCL on a.ClientID = MCL.ClientID"
    tmp = tmp & " inner join Mst_Location ML on a.locationID = ML.LocationID"
    tmp = tmp & " inner join Mst_State s on ML.StateID = s.StateID"
    tmp = tmp & " Where b.BalanceBags > 0  And (b.ContractID = 0 Or b.ContractID is null) and b.NContractID in (Select Distinct ContractID From Mst_StorageContract Where ExtendedPeriod <= '" & Date & "')"
    If wc <> "" Then
       tmp = tmp & " And " & wc
    End If
    tmp = tmp & " Group by s.statename,ML.LocationName,  a.ContractID,a.ContractName,MCL.ClientName,b.SCMFlag,a.BookedFrom,a.ExtendedPeriod,a.Flag"


End If


Call TmpTable

ProgressBar1.Value = 0
LblTotalRecord.Caption = TmpRs.RecordCount
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 1
   For I = 1 To TmpRs.RecordCount
       If OptExpiredContract.Value = False Then
            For C = 0 To MSHFlexGrid1.Cols - 1
                If C = 5 Or C = 6 Or C = 8 Or C = 9 Then
                   MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", Format(TmpRs.Fields(C), "dd-MMM-yyyy"))
                Else
                   MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
                End If
            Next
       Else
       
            For C = 0 To MSHFlexGrid1.Cols - 1
                If C = 6 Or C = 7 Then
                   MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", Format(TmpRs.Fields(C), "dd-MMM-yyyy"))
                Else
                   MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
                End If
            Next
       
       End If
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   CmdExcel.Enabled = True
   MsgBox "Done!!!"
Else
   CmdExcel.Enabled = False
   MsgBox " No Record Found "
End If

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
CmdExcel.Enabled = False
Call Grid_Head
End Sub
Private Sub Grid_Head()
 
If OptExpiredContract.Value = True Then


    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 11
        .Font.Size = 10
        .AllowUserResizing = flexResizeBoth
        .TextMatrix(0, 0) = "State Name"
        .TextMatrix(0, 1) = "Location Name"
        .TextMatrix(0, 2) = "Contract ID"
        .TextMatrix(0, 3) = "Contract Name"
        .TextMatrix(0, 4) = "Client Name"
        .TextMatrix(0, 5) = "SCM Flag"
        .TextMatrix(0, 6) = "Contract Start Date"
        .TextMatrix(0, 7) = "Contract End Date "
        .TextMatrix(0, 8) = "Flag"
        .TextMatrix(0, 9) = "Bal Bags"
        .TextMatrix(0, 10) = "Bal Weight"
        
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .ColWidth(2) = 1800
        .ColWidth(3) = 2500
        .ColWidth(4) = 2000
        .ColWidth(5) = 1900
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1900
        .ColWidth(9) = 1500
        .ColWidth(10) = 1200

        .RowHeight(0) = 800
    End With

Else
    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 15
        .Font.Size = 10
        .AllowUserResizing = flexResizeBoth
        .TextMatrix(0, 0) = "Contract ID"
        .TextMatrix(0, 1) = "Contract Name"
        .TextMatrix(0, 2) = "Location Name"
        .TextMatrix(0, 3) = "Client"
        .TextMatrix(0, 4) = "Commodity"
        .TextMatrix(0, 5) = "Contract Start Date"
        .TextMatrix(0, 6) = "Contract End Date"
        .TextMatrix(0, 7) = "Remarks "
        .TextMatrix(0, 8) = "Period From Date"
        .TextMatrix(0, 9) = "Period To Date"
        .TextMatrix(0, 10) = "Billing Cycle"
        .TextMatrix(0, 11) = "Reservation Qty "
        .TextMatrix(0, 12) = "Rate"
        .TextMatrix(0, 13) = "Details Remarks"
        .TextMatrix(0, 14) = "Flag"
        
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 4500
        .ColWidth(2) = 1800
        .ColWidth(3) = 2500
        .ColWidth(4) = 2000
        .ColWidth(5) = 1900
        .ColWidth(6) = 1500
        .ColWidth(7) = 3500
        .ColWidth(8) = 1900
        .ColWidth(9) = 1500
        .ColWidth(10) = 1200
        .ColWidth(11) = 1500
        .ColWidth(12) = 1000
        .ColWidth(13) = 1500
        .ColWidth(14) = 1000
        .RowHeight(0) = 800
    End With
End If

End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
'If CmbLocation.Text = "" Then
'   CmbCMPName.Text = ""
'   Exit Sub
'End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   'MsgBox "Invalid selection "
   'CmbLocation.SetFocus
   mLocationID = 0
   Exit Sub
End If

'If tmp <> CmbLocation.Text Then
'   CmbCMPName.Text = ""
'End If

mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   mCommodityID = 0
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!! "
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID
End Sub

Private Sub TxtResevationQty_LostFocus()
If TxtResevationQty = "" Then Exit Sub

If Len(TxtResevationQty) = 1 Then
   MsgBox "Invalid format!!!"
End If


If Len(TxtResevationQty) > 1 Then
   If InStr(1, TxtResevationQty, ">") > 0 Then
       If IsNumeric(Replace(TxtResevationQty, ">", "")) = False Then
          MsgBox "Enter numeric value for data extracting!!!"
          Exit Sub
       End If
   ElseIf InStr(1, TxtResevationQty, "=") > 0 Then
       If IsNumeric(Replace(TxtResevationQty, "=", "")) = False Then
          MsgBox "Enter numeric value for data extracting!!!"
          Exit Sub
       End If
   Else
      MsgBox "Invalid format!!!"
      Exit Sub
   End If
End If

End Sub
