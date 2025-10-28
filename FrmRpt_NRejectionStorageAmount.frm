VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_NRejectionStorageAmount 
   Caption         =   "NCDEX Rejection Storage Amount"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9795
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15030
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   405
         Left            =   120
         TabIndex        =   23
         Top             =   9195
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.OptionButton Option2 
         Caption         =   "NCDEX 7 Days Storage Amount"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "NCDEX Rejection Storage Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   350
         Left            =   13770
         TabIndex        =   9
         Top             =   1230
         Width           =   1005
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12765
         TabIndex        =   8
         Top             =   1230
         Width           =   1005
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   11760
         TabIndex        =   7
         Top             =   1230
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   810
         Width           =   4125
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   9720
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   810
         Width           =   2685
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   3165
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   810
         Width           =   2325
      End
      Begin VB.TextBox TxtSCxTFNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12450
         MaxLength       =   25
         TabIndex        =   6
         Top             =   810
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         Format          =   109379585
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         Format          =   109379585
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7395
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   13044
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
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
         Left            =   4080
         TabIndex        =   20
         Top             =   1305
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Storage Rate not availabe in Storage Rate Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4725
         TabIndex        =   19
         Top             =   1200
         Width           =   2415
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
         Left            =   9720
         TabIndex        =   18
         Top             =   1260
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
         Left            =   7395
         TabIndex        =   17
         Top             =   1260
         Width           =   2205
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
         Height          =   255
         Left            =   2085
         TabIndex        =   15
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
         Height          =   255
         Left            =   442
         TabIndex        =   14
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label34 
         Caption         =   "CMP"
         Height          =   300
         Left            =   7455
         TabIndex        =   13
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label39 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   10695
         TabIndex        =   12
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label40 
         Caption         =   "Location"
         Height          =   300
         Left            =   3990
         TabIndex        =   11
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label49 
         Caption         =   "Store Man CITF No"
         Height          =   225
         Left            =   12405
         TabIndex        =   10
         Top             =   600
         Width           =   1530
      End
   End
End
Attribute VB_Name = "FrmRpt_NRejectionStorageAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSLocationID, msSM_prefix As Variant

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 17 '16 '17

If Option1.Value = True Then
   MSHFlexGrid1.TextMatrix(0, 0) = "CITF No"
   MSHFlexGrid1.TextMatrix(0, 1) = "CITF Date"
Else
   MSHFlexGrid1.TextMatrix(0, 0) = "CDTF No"
   MSHFlexGrid1.TextMatrix(0, 1) = "CDTF Date"
End If
MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 6) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 7) = "G~S~L"
MSHFlexGrid1.TextMatrix(0, 8) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 9) = "No Of Days"
MSHFlexGrid1.TextMatrix(0, 10) = "Billing Cycle"
MSHFlexGrid1.TextMatrix(0, 11) = "No Of Bags"
MSHFlexGrid1.TextMatrix(0, 12) = "Weight"
MSHFlexGrid1.TextMatrix(0, 13) = "Billing Unit"
MSHFlexGrid1.TextMatrix(0, 14) = "Rate"
MSHFlexGrid1.TextMatrix(0, 15) = "Amount"
MSHFlexGrid1.TextMatrix(0, 16) = "ADHoc Invoice Amount"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1200
MSHFlexGrid1.ColWidth(2) = 2400
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 2400
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(14) = 1200
MSHFlexGrid1.ColWidth(15) = 1600
MSHFlexGrid1.ColWidth(16) = 1800


If Option2.Value = True Then
   MSHFlexGrid1.ColWidth(8) = 0
End If

MSHFlexGrid1.RowHeight(0) = 800
MSHFlexGrid1.WordWrap = True


End Sub

Private Sub CmdClear_Click()

SFrom.Value = Date - 30
STo.Value = Date
CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSCommodity.Text = ""
TxtSCxTFNo = ""
LblTotalRecord.Caption = ""
Call Grid_Head
ProgressBar1.Value = 0

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
If Option1.Value = True Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXRejectionStorage.xls")
Else
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEX7DaysStorage.xls")
End If
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
If Option1.Value = True Then
   Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXRejectionStorage.xls")
Else
   Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEX7DaysStorage.xls")
End If
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
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
        
        MSHFlexGrid1.row = I
        MSHFlexGrid1.Col = C
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created!!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()


On Error GoTo msgError

Dim wc As String
Dim tmp1 As Variant
wc = ""
Call Grid_Head
LblTotalRecord.Caption = ""

MsgBox "Wait till next message!!!"

If STo.Value < SFrom.Value Then
   MsgBox "'To Date' must be greater than 'From Date'!!!"
   STo.SetFocus
   Exit Sub
End If

If Option1.Value = True Then

   tmp = "Select TCD.CxTFNo,TCD.CxTFDate,MC.CMPName,MComm.Commodity,MClient.ClientID,MClient.ClientName,MClient.DepositoryID, "
   tmp = tmp & " MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as 'G~S~L',MG.CDTFDate,(datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) NoOfDay, "
   tmp = tmp & " 'Daily' Cycle,Sum(TCG.NoOfBags) as 'Bags',"
   tmp = tmp & " Sum(TCG.DematWeight) as 'Weight','Quantity' Unit,MNRC.Rate, (datediff(day,MG.CDTFDate,TCD.CxTFDate) + 1) * Sum(TCG.DematWeight) * Rate as 'Amount',ML.LocationID,MClient.ClientIDRow "
   tmp = tmp & " From Txn_CxTF_Details TCD"
   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
   tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MComm.CommodityGroupID = MCG.CommodityGroupID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
   tmp = tmp & " Left Join Mst_Client MClient On MG.ClientIDRow = MClient.ClientIDRow"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
   tmp = tmp & " Where TCD.TxnType = 'W' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 1"
   tmp1 = " Group By TCD.CxTFNo,TCD.CxTFDate,MC.CMPName,MComm.Commodity,MClient.ClientID,MClient.ClientName,MClient.DepositoryID,"
   tmp1 = tmp1 & " MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo,MG.CDTFDate,MNRC.Rate,ML.LocationID,MClient.ClientIDRow "
   tmp1 = tmp1 & " Order By TCD.CxTFNo,TCD.CxTFDate,MC.CMPName "

   If Gen_DBType = "MDB" Then
      wc = " And TCD.CxTFDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
   Else
      wc = " And TCD.CxTFDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
   End If
ElseIf Option2.Value = True Then

   tmp = " Select TCG.CxTFNo,TCD.CxTFDate,MC.CMPName,MComm.Commodity,MClient.ClientID,MClient.ClientName,MClient.DepositoryID,"
   tmp = tmp & " MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo as 'G~S~L','','7' NoOfDay,'Daily' Cycle,Sum(TCG.NoOfBags) as 'Bags',"
   tmp = tmp & " Sum(TCG.GrossWeight) as 'Weight','Quantity' Unit,MNRC.Rate, 7 * Sum(TCG.GrossWeight) * Rate as 'Amount',ML.LocationID,MClient.ClientIDRow "
   tmp = tmp & " From Txn_CxTF_Details TCD"
   tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTFNo = TCG.CxTFNo And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType"
   tmp = tmp & " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Commodity MComm On TCD.CommodityID = MComm.CommodityID"
'   tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MComm.CommodityGroupID = MCG.CommodityGroupID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Godown MGod On MG.GodownID = MGod.GodownID"
   tmp = tmp & " Left Join Mst_Client MClient On TCG.ClientIDRow = MClient.ClientIDRow"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCard MNR On ML.LocationID = MNR.LocationID And MComm.CommodityGroupID = MNR.CommodityGroupID"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail MNRC On MNR.NSRCID = MNRC.NSRCID And MG.GodownID = MNRC.GodownID"
   tmp = tmp & " Where TCG.TxnType = 'D' and TCD.ExchangeTypeID = 2 and TCD.BookedSpaceFlag = 0"
   tmp1 = " Group By TCG.CxTFNo,TCD.CxTFDate,MC.CMPName,MComm.Commodity,MClient.ClientID,MClient.ClientName,MClient.DepositoryID,"
   tmp1 = tmp1 & " MGod.GodownNo +'~'+ MG.StackNo  +'~'+ MG.LotNo,MNRC.Rate,ML.LocationID,MClient.ClientIDRow "
   tmp1 = tmp1 & " Order By TCG.CxTFNo,TCD.CxTFDate,MC.CMPName "

End If
'Wc = " "

'If Gen_DBType = "MDB" Then
'   wc = " And TCD.CxTFDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
'Else
'   wc = " And TCD.CxTFDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
'End If

If CmbPreFix.Text <> "" Then
   wc = wc & " And TCD.SM_Prefix = '" & msSM_prefix & "'" ' ''" & CmbPreFix.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If

If CmbSCmp.Text <> "" Then
   wc = wc & " And MC.CMPName = '" & CmbSCmp.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      wc = wc & " And MC." & Gen_WcCMP
   End If
End If

If CmbSCommodity.Text <> "" Then
   wc = wc & " And MComm.Commodity = '" & CmbSCommodity.Text & "'"
End If


If TxtSCxTFNo <> "" Then
   wc = wc & " And TCG.CxTFNo = " & Val(TxtSCxTFNo) & ""
End If


tmp = tmp & wc & tmp1

Call TmpTable

x = 1
ProgressBar1.Max = TmpRs.RecordCount

For I = 1 To TmpRs.RecordCount
    
    If TmpRs!CxTFDate >= SFrom.Value And TmpRs!CxTFDate <= STo.Value Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1 'TmpRs.RecordCount + 2
       For C = 0 To TmpRs.Fields.Count - 2
           MSHFlexGrid1.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(x, 1) = Format(MSHFlexGrid1.TextMatrix(x, 1), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(x, 8) = Format(MSHFlexGrid1.TextMatrix(x, 8), "dd-mmm-yyyy")
       
       If IsNull(TmpRs!ClientIDRow) = False Then
          MSHFlexGrid1.TextMatrix(x, 16) = GetADhocInvoiceAmount(TmpRs!LocationID, TmpRs!ClientIDRow) 'Format(MSHFlexGrid1.TextMatrix(I, 8), "dd-mmm-yyyy")
       End If
       If (Val(MSHFlexGrid1.TextMatrix(x, 14)) = 0) Then
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.Col = C
              MSHFlexGrid1.row = x
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       End If
       x = x + 1
    End If
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
LblTotalRecord.Caption = MSHFlexGrid1.Rows - 2 ' mpRs.RecordCount
CmdExcel.Enabled = True

MsgBox "Done!!!"

Exit Sub

msgError:
MsgBox "Please add more filter to generate report!!!"



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
CmdExcel.Enabled = False

End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
'Call TableMst_Location

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If



CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   mLocationID = Null
   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbPreFix.SetFocus
   Exit Sub
End If

If tmp <> CmbPreFix.Text Then
   CmbSCmp.Text = ""
   mCMPID = Null
End If
mSLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!Locationname
msSM_prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbSCmp_GotFocus()
tmp = CmbSCmp.Text
If CmbPreFix.Text = "" Then
   MsgBox "Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
If Gen_WcCMP <> "" Then
   Call TableMst_CMP(" Where LocationID = " & mSLocationID & " And " & Gen_WcCMP)
Else
   Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
End If

CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp
End Sub



Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSCmp.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbSCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbSCommodity.Text = tmp
End Sub

Private Sub CmbSCommodity_LostFocus()
If CmbSCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSCommodity.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID

End Sub

Private Sub Option1_Click()


If Option1.Value = True Then
   Label49.Caption = "Store Man CITF No"
ElseIf Option2.Value = True Then
   Label49.Caption = "Store Man CDTF No"
End If

Call Grid_Head
End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub

Private Function GetADhocInvoiceAmount(mLocID_ As Variant, mClientID_ As Variant)
Dim Retval As Variant
Retval = 0

tmp = "Select LocationID,ClientIDRow,Sum(Amount) As 'Amount' From Txn_TempInvoice "
tmp = tmp & " Where LocationID = " & mLocID_ & " And ClientID = " & mClientID_ & " "
tmp = tmp & " And Month(InvoiceDate) >= " & Month(SFrom.Value) & " And  Month(InvoiceDate) <= " & Month(STo.Value) & ""
tmp = tmp & " And Year(InvoiceDate) >= " & Year(SFrom.Value) & " And Year(InvoiceDate) <= " & Year(SFrom.Value) & " And Flag <> 'R'  "

If Option1.Value = True Then
   tmp = tmp & " And  IETypeID = 2 "
ElseIf Option2.Value = True Then
   tmp = tmp & " And  IETypeID = 1 "
End If

tmp = tmp & " Group By LocationID,ClientID "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1!Amount
End If

GetADhocInvoiceAmount = Retval

End Function
