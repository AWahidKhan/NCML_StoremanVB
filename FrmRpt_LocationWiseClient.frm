VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_LocationWiseClient 
   Caption         =   "Location wise Client List"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   45
      TabIndex        =   18
      Top             =   9255
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   15105
      Begin VB.TextBox TxtNAVCode 
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
         Left            =   105
         TabIndex        =   21
         Top             =   1080
         Width           =   1905
      End
      Begin VB.CheckBox CheckDate 
         Caption         =   "Apply Date filter"
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
         Left            =   2595
         TabIndex        =   5
         Top             =   825
         Width           =   1770
      End
      Begin VB.TextBox TxtSClientName 
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
         Left            =   8355
         TabIndex        =   4
         Top             =   405
         Width           =   6585
      End
      Begin VB.ComboBox CmbSNAVCodeMapped 
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
         ItemData        =   "FrmRpt_LocationWiseClient.frx":0000
         Left            =   6270
         List            =   "FrmRpt_LocationWiseClient.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   2040
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
         TabIndex        =   9
         Top             =   960
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
         TabIndex        =   8
         Top             =   960
         Width           =   885
      End
      Begin VB.ComboBox CmbSExchangeTypeID 
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
         ItemData        =   "FrmRpt_LocationWiseClient.frx":0017
         Left            =   3720
         List            =   "FrmRpt_LocationWiseClient.frx":0019
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   390
         Width           =   2520
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
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   1
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
         TabIndex        =   10
         Top             =   960
         Width           =   885
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   4995
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
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
         Format          =   61865985
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   6915
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
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
         Format          =   61865985
         CurrentDate     =   36494
      End
      Begin VB.Label Label4 
         Caption         =   "NAV Code"
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
         Left            =   585
         TabIndex        =   22
         Top             =   795
         Width           =   1170
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Stock Deposited between"
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
         Left            =   2595
         TabIndex        =   20
         Top             =   1140
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "And"
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
         Left            =   6435
         TabIndex        =   19
         Top             =   1140
         Visible         =   0   'False
         Width           =   360
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
         Left            =   8355
         TabIndex        =   17
         Top             =   960
         Width           =   2205
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
         Left            =   10755
         TabIndex        =   16
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   11062
         TabIndex        =   15
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label30 
         Caption         =   "NAV Code Mapped"
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
         Left            =   6368
         TabIndex        =   14
         Top             =   150
         Width           =   1845
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
         Left            =   1485
         TabIndex        =   12
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Exchange Type"
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
         Left            =   4320
         TabIndex        =   11
         Top             =   150
         Width           =   1455
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7620
      Left            =   45
      TabIndex        =   13
      Top             =   1560
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   13441
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
Attribute VB_Name = "FrmRpt_LocationWiseClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSLocationID, mSExchangeTypeID As Variant

Private Sub CheckDate_Click()
If CheckDate.Value = vbChecked Then
   Label44.Visible = True
   Label2.Visible = True
   TxtFromDate.Visible = True
   TxtToDate.Visible = True
Else
   Label44.Visible = False
   Label2.Visible = False
   TxtFromDate.Visible = False
   TxtToDate.Visible = False
End If
End Sub

Private Sub CmbsExchangeTypeID_GotFocus()
tmp = CmbSExchangeTypeID.Text
Call TableMst_ExchangeType
CmbSExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbsExchangeTypeID_LostFocus()
If CmbSExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeTypeID.SetFocus
   Exit Sub
End If

mSExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmbSNAVCodeMapped_LostFocus()
If CmbSNAVCodeMapped.Text = "" Then Exit Sub
If LCase(CmbSNAVCodeMapped.Text) <> "no" And LCase(CmbSNAVCodeMapped.Text) <> "yes" Then
   MsgBox "invalid Selection."
   CmbSNAVCodeMapped.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date - 365
TxtToDate.Value = Date
CheckDate.Value = vbUnchecked
CmbSLocationID.Text = ""
ProgressBar1.Value = 0
CmbSExchangeTypeID.Text = ""
CmbSNAVCodeMapped.Text = ""
TxtSClientName = ""
TxtNAVCode = ""
Call Grid_Head

End Sub

Private Sub CmdExcel_Click()

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_LocationwiseClient.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_LocationwiseClient.xls")
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

ProgressBar1.Value = 0
Call Grid_Head
wc = ""

If CheckDate.Value = vbChecked Then
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
   wc = " Where MG.CDTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
End If
MsgBox "Wait till next message !!!"

'tmp = "Select  Distinct ML.LocationName,MC.ClientID,MC.ClientName,MC.ClientIDRow,ME.ExchangeType,MC.NAVCode  "
'tmp = tmp & " From Mst_GSL MG"
'tmp = tmp & " Inner Join Mst_Client MC On MG.ClientIDRow = MC.ClientIDRow"
'tmp = tmp & " Inner Join Mst_CMP MCMP On MG.CMPID = MCMP.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_ExchangeType ME On MG.ExchangeTypeID = ME.ExchangeTypeID"

tmp = " Select  Distinct ML.LocationName,MC.ClientID,MC.ClientName,MC.ClientIDRow,ME.ExchangeType,MC.NAVCode,MC.RNAVCode, MCom.Commodity,Sum(BalanceBags) 'Bags',Sum(BalanceWeight) 'Weight', ML.LocationName+MC.NAVCode LocationNAVCode,ML.StateID,MG.CommodityID"
tmp = tmp & " From Mst_GSL MG"
tmp = tmp & " Inner Join Mst_Client MC On MG.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_CMP MCMP On MG.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On MG.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Commodity MCom On  MG.CommodityID = MCom.CommodityID"


If CmbSLocationID.Text <> "" Then
   If wc <> "" Then
      wc = wc & " and ML.LocationName = '" & CmbSLocationID.Text & "' "
   Else
      wc = " Where ML.LocationName = '" & CmbSLocationID.Text & "' "
   End If
Else
   If Gen_WcLocation <> "" Then
      If wc <> "" Then
         wc = wc & " And ML." & Gen_WcLocation
      Else
         wc = " Where ML." & Gen_WcLocation
      End If
   End If
End If

If CmbSExchangeTypeID.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And ME.ExchangeType = '" & CmbSExchangeTypeID.Text & "' "
   Else
      wc = "Where ME.ExchangeType = '" & CmbSExchangeTypeID.Text & "' "
   End If
End If

If TxtNAVCode.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And (MC.NAVCode = '" & TxtNAVCode.Text & "' or MC.RNAVCode = '" & TxtNAVCode.Text & "')"
   Else
      wc = "Where (MC.NAVCode = '" & TxtNAVCode.Text & "' or MC.RNAVCode = '" & TxtNAVCode.Text & "')"
   End If
End If

If CmbSNAVCodeMapped.Text <> "" Then
   If LCase(CmbSNAVCodeMapped.Text) = "yes" Then
      If wc <> "" Then
         wc = wc & " And (MC.NAVCode <> '' Or MC.RNAVCode <> '') "
      Else
         wc = " Where (MC.NAVCode <> '' Or MC.RNAVCode <> '') "
      End If
   Else
      If wc <> "" Then
         wc = wc & " And (MC.NAVCode is Null Or MC.NAVCode = '') And (MC.RNAVCode is Null Or MC.RNAVCode = '') "
      Else
         wc = " Where (MC.NAVCode is Null Or MC.NAVCode = '') And (MC.RNAVCode is Null Or MC.RNAVCode = '') "
      End If
   End If
End If


If TxtSClientName.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And (MC.ClientName Like '%" & TxtSClientName & "%') "
   Else
      wc = " Where (MC.ClientName Like '%" & TxtSClientName & "%') "
   End If
End If

tmp1 = " Group By ML.LocationName,MC.ClientID,MC.ClientName,MC.ClientIDRow,ME.ExchangeType,MC.NAVCode, MC.RNAVCode, MCom.Commodity ,ML.StateID,MG.CommodityID"
tmp1 = tmp1 & " Order By ML.LocationName,MC.NAVCode,ME.ExchangeType"
tmp = tmp & " " & wc & " " & tmp1
Call TmpTable

tmp = " Select LocationName+NAVCode LocationNAVCode, sum(TIAmount - RecAmount) OutStandingAmount"
tmp = tmp & " From ("
tmp = tmp & " Select ML.LocationName, TI.NAVCode, TI.TIAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ isnull(TI.AdjAmountFAG,0) +"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) +"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID) as RecAmount"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Inner Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Where TINo<>''  And"
If CheckDate.Value = vbChecked Then
   wc = " Where TI.TIDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' and '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
End If
tmp = tmp & " TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+ (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R')+AdjAmountFAG)>0"
tmp = tmp & " group by ML.LocationName, TI.NAVCode, TI.TIAmount, TI.AdjAmountFAG, TI.TIID"
tmp = tmp & " )A"
tmp = tmp & " Group By LocationName, NAVCode"
Call Tmp7Table

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 2

   For mRow = 1 To TmpRs.RecordCount
       For mCol = 0 To MSHFlexGrid1.Cols - 4
           MSHFlexGrid1.TextMatrix(mRow, mCol) = IIf(IsNull(TmpRs.Fields(mCol)), "", TmpRs.Fields(mCol))
       Next
       If Val(MSHFlexGrid1.TextMatrix(mRow, 8)) <= 0 Then
          MSHFlexGrid1.TextMatrix(mRow, 9) = "0"
       End If
       If MSHFlexGrid1.TextMatrix(mRow, 0) <> MSHFlexGrid1.TextMatrix(mRow - 1, 0) Or MSHFlexGrid1.TextMatrix(mRow, 5) <> MSHFlexGrid1.TextMatrix(mRow - 1, 5) Then
         If TmpRs7.RecordCount > 0 Then
             TmpRs7.MoveFirst
             TmpRs7.Find " LocationNAVCode = '" & TmpRs!LocationNAVCode & "'"
             If Not TmpRs7.EOF Then
                 MSHFlexGrid1.TextMatrix(mRow, 10) = IIf(IsNull(TmpRs7!OutStandingAmount), "", TmpRs7!OutStandingAmount)
             End If
         End If
       End If
       MSHFlexGrid1.TextMatrix(mRow, 11) = GetCommodityLatestRate(TmpRs!CommodityID, TmpRs!StateID)
       MSHFlexGrid1.TextMatrix(mRow, 12) = Format(Val(MSHFlexGrid1.TextMatrix(mRow, 9)) * Val(MSHFlexGrid1.TextMatrix(mRow, 11)), "#####0.00")
   
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

Call Grid_Head

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If

Call CmdClear_Click
Call ClearData

End Sub
Private Sub ClearData()

CmbSLocationID.Text = ""
CmbSExchangeTypeID.Text = ""
CmbSNAVCodeMapped.Text = ""

Call Grid_Head


End Sub
Private Sub Grid_Head()
LblTotalRecord.Caption = ""
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 13
    .RowHeight(0) = 400
    .AllowUserResizing = flexResizeColumns
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "Client ID"
    .TextMatrix(0, 2) = "Client Name"
    .TextMatrix(0, 3) = "ClientIDRow"
    .TextMatrix(0, 4) = "ExchangeType"
    .TextMatrix(0, 5) = "NAV Code"
    .TextMatrix(0, 6) = "RNAV Code"
    .TextMatrix(0, 7) = "Commodity"
    .TextMatrix(0, 8) = "Bags"
    .TextMatrix(0, 9) = "Weight"
    If CheckDate.Value = 0 Then
       .TextMatrix(0, 10) = "Out Standing Amount of Invoices"
    Else
       .TextMatrix(0, 10) = "Out Standing Amount of Invoices generated between this dates"
    End If
    .TextMatrix(0, 11) = "Commodity Price"
    .TextMatrix(0, 12) = "AUM"
    
    .ColWidth(0) = 1800
    .ColWidth(1) = 1100
    .ColWidth(2) = 4000
    .ColWidth(3) = 0
    .ColWidth(4) = 1400
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1600
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1400
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
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

Private Function GetCommodityLatestRate(mCommodityID_ As Variant, mStateID_ As Variant) As Variant

Dim mComRate As Variant

'tmp = " Select MCP.YearMonth, MCP.Day28,MCP.Day29,MCP.Day30,MCP.Day31 from Mst_CommodityPrice MCP"
'tmp = tmp & " Where MCP.StateID = " & mStateID_ & " And MCP.CommodityID = " & mCommodityID_ & ""
'tmp = tmp & " Order By MCP.YearMonth "


tmp = " Select MCP.StateID,MCP.CommodityID,MCP.YearMonth,"
tmp = tmp & " Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
tmp = tmp & " Then"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " Else"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " End [Rate]"
tmp = tmp & " from Mst_CommodityPrice MCP"
tmp = tmp & " Where StateID = " & mStateID_ & " And CommodityID = " & mCommodityID_ & "  And Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Group By StateID,CommodityID )"


Call Tmp2Table

mComRate = 0
'If TmpRs2.RecordCount > 0 Then
'   TmpRs2.MoveLast
'   If Right(TmpRs2!YearMonth, 2) = "01" Or Right(TmpRs2!YearMonth, 2) = "03" Or Right(TmpRs2!YearMonth, 2) = "05" Or Right(TmpRs2!YearMonth, 2) = "07" Or Right(TmpRs2!YearMonth, 2) = "08" Or Right(TmpRs2!YearMonth, 2) = "10" Or Right(TmpRs2!YearMonth, 2) = "12" Then
'      mComRate = TmpRs2!Day31
'   ElseIf Right(TmpRs2!YearMonth, 2) = "04" Or Right(TmpRs2!YearMonth, 2) = "06" Or Right(TmpRs2!YearMonth, 2) = "09" Or Right(TmpRs2!YearMonth, 2) = "11" Then
'      mComRate = TmpRs2!Day30
'   Else
'      If IsNull(TmpRs2!Day29) = False Then
'         mComRate = TmpRs2!Day29
'      Else
'         mComRate = TmpRs2!Day28
'      End If
'   End If
'End If

If TmpRs2.RecordCount > 0 Then
   mComRate = IIf(IsNull(TmpRs2!Rate), 0, TmpRs2!Rate)
End If

GetCommodityLatestRate = mComRate
End Function
