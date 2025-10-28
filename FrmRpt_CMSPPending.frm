VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_CMSPPending 
   Caption         =   "CMSP-CMSE Pending Report"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   14895
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
         Left            =   9810
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   5025
      End
      Begin VB.CheckBox ChkMAW 
         Appearance      =   0  'Flat
         Caption         =   "MAW Weight"
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
         Height          =   255
         Left            =   8760
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox CmbCommodity 
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
         Left            =   9810
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   165
         Width           =   5010
      End
      Begin VB.OptionButton Option4 
         Caption         =   "CMS-P Pending"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1830
      End
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
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
         Height          =   255
         Left            =   11085
         TabIndex        =   20
         Top             =   1065
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox ChkUpdateBatch 
         Appearance      =   0  'Flat
         Caption         =   "Update CMS-P Batch"
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
         Height          =   345
         Left            =   12480
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox CmbBatch 
         Enabled         =   0   'False
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
         Left            =   5760
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         Caption         =   "CMS-P Old Batch"
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
         Left            =   2160
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
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
         Left            =   13350
         TabIndex        =   10
         Top             =   1440
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
         Height          =   375
         Left            =   10440
         TabIndex        =   6
         Top             =   1440
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
         Height          =   375
         Left            =   11895
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CMS-P Pending"
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
         Left            =   120
         TabIndex        =   0
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CMS-P Mapped but CMS-E Pending for Mapping"
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
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4695
      End
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "FrmRpt_CMSPPending.frx":0000
         Left            =   5760
         List            =   "FrmRpt_CMSPPending.frx":0007
         TabIndex        =   4
         Top             =   165
         Width           =   2415
      End
      Begin VB.Label Label4 
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
         Left            =   8760
         TabIndex        =   25
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label3 
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
         Height          =   285
         Left            =   8760
         TabIndex        =   22
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label Label8 
         Caption         =   "Non Marketable Lot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   780
         TabIndex        =   19
         Top             =   1110
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FFFF&
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
         Left            =   120
         TabIndex        =   18
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Batch No"
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
         Left            =   4920
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   8040
         TabIndex        =   13
         Top             =   1590
         Width           =   210
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
         Left            =   5640
         TabIndex        =   12
         Top             =   1590
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Type"
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
         Left            =   4800
         TabIndex        =   11
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6975
      Left            =   120
      TabIndex        =   7
      Top             =   1950
      Width           =   14895
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   6555
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11033
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
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
Attribute VB_Name = "FrmRpt_CMSPPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Dim MaxID As Variant
Dim mSM_Prefix As Variant

Private Sub ChkCheckAll_Click()

If MSHFlexGrid1.Rows = 1 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If Option1.Value = True Then
   If ChkCheckAll.Value = vbChecked Then
      For I = 1 To MSHFlexGrid1.Rows - 2
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 20) = strChecked
      Next
   Else
      For I = 1 To MSHFlexGrid1.Rows - 2
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 20) = strUnChecked
      Next
   End If
End If

End Sub

Private Sub CmbBatch_GotFocus()
tmp = "Select distinct CMSPBatchID from Mst_GSL Where CMSPBatchID > 0  order by CMSPBatchID"
Call Tmp1Table
CmbBatch.Clear
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       CmbBatch.AddItem TmpRs1.Fields(0)
       MaxID = TmpRs1.Fields(0) + 1
       TmpRs1.MoveNext
   Next
Else
   MaxID = 1
End If
End Sub

Private Sub CmbBatch_LostFocus()
If CmbBatch.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "CMSPBatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbBatch.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity(" Where CommTrackFlag = '1'")

CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Commodity found"
   CmdCommodity.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodity.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbCommodity")
End If
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"

If RsMst_Commodity.EOF Then
   MsgBox "Invalid Commodity selection "
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity.Fields(0)
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
   CmbSLocationID.AddItem RsMst_Location!LocationName
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
mLocationID = RsMst_Location!LocationID
mSM_Prefix = IIf(IsNull(RsMst_Location!SM_Prefix), "", RsMst_Location!SM_Prefix)

End Sub

Private Sub CmbType_LostFocus()

If CmbType.Text = "" Then Exit Sub

If LCase(CmbType.Text) <> "marketable" Then
   MsgBox "Invalid Selection!!!"
   CmbType.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbType.Text = ""
CmdExcel.Enabled = False
Option4.Value = True
CmbBatch.Text = ""
ChkUpdateBatch.Value = 0
ChkUpdateBatch.Enabled = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc, DispFlag, Filt, Sdat, Edat As Variant
ChkUpdateBatch.Value = 0
Call Grid_Head
ProgressBar1.Value = 0
If Option3.Value = True Then
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch!!!"
      CmbBatch.SetFocus
      Exit Sub
   End If
End If

 If Option1.Value = True Then
   tmp = "SELECT  GM.SM_Prefix 'Location ID',GM.Mst_GSL_ID 'GSL ID.',GM.CxTF_DetailsID 'CDTF No',GM.CDTFDate 'CDTF Date',CMP.CMPName 'CMP', "
   tmp = tmp & " COM.Commodity 'Commodity',CM.DepositoryID 'dp_id',CM.ClientIDRow 'CM_ClientID',CM.ClientName 'Client Name', "
   tmp = tmp & " GoM.GodownNo + '~' + GM.StackNo + '~' + GM.LotNo 'GSL No.', GoM.Address,GM.DepositeBags 'No of Bags', GM.GrossWt 'Weight in MT', "
   tmp = tmp & " CM.Address 'Depositor Address',DM.DepositoryName 'Depositor Name',GM.CRAmount,GM.CRDate,GM.KRNo 'KR No',GM.KRDate 'KR Date',MA.AgentName,GM.CMSPBatchID,GM.DepositeWeight,CMD.LotSize,CMD.Variation,GM.CMSPBatchGeneratedBy,GM.CMSPBatchGeneratedDate,GM.CMSPUpdatedBy,GM.CMSPUpdatedDate"
   ''''added after trf module
   tmp = tmp & " ,TRF.TRFNo, TRF.SamplingDate, CM.CMSPCode, GoM.CommtrackNCDEXCode"
   tmp = tmp & " FROM Mst_GSL GM Inner Join Mst_Client CM On GM.ClientID=CM.ClientID "
   tmp = tmp & " Inner Join Mst_Depository DM On CM.DepositoryID=DM.DepositoryID Inner Join Mst_Godown GoM On GoM.GodownID=GM.GodownID "
   tmp = tmp & " Inner Join Mst_CMP CMP On CMP.CMPID=GM.CMPID Inner Join Mst_Commodity COM On COM.CommodityID=GM.CommodityID "
   tmp = tmp & " Inner join Mst_CommodityDetail CMD On CMD.CommodityID=GM.CommodityID AND CMD.ExchangeTypeID=GM.ExchangeTypeID "
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On GM.CxTF_DetailsID = TCD.CxTF_DetailsID And GM.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D'"
   tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID"
   ''''added after trf module
   tmp = tmp & " Left Join Txn_TRFDetail TRFD on GM.Mst_GSL_ID = TRFD.Mst_GSL_ID and GM.SM_Prefix = TRFD.SM_Prefix"
   tmp = tmp & " Left Join Txn_TRF TRF on TRFD.TRFID = TRF.TRFID"
   tmp = tmp & " WHERE  GM.CMSPBatchID = 0 And COM.CommTrackFlag = '1' And GM.ExchangeTypeID=2 AND (GM.ISINID='0' OR GM.ISINID='' OR GM.ISINID IS NULL) AND (GM.Status ='O'  ) AND (KRNo=0 OR KRNo IS NULL) "
   If CmbCommodity.Text <> "" Then
      tmp = tmp & " And COM.Commodity = '" & CmbCommodity.Text & "' "
   End If
   
   If CmbSLocationID.Text <> "" Then
      tmp = tmp & " And GM.SM_Prefix = '" & mSM_Prefix & "' "
   End If
   
   If LCase(CmbType.Text) = "marketable" Then
      If ChkMAW.Value = 0 Then
         tmp = tmp & " AND (GM.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
      Else
         tmp = tmp & " AND (GM.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
      End If
'   ElseIf LCase(CmbType.Text) = "non marketable" Then
'      tmp = tmp & " AND (GM.DepositeWeight<CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
   End If
   tmp = tmp & " and gm.Mst_GSL_ID not in(select a.Mst_GSL_ID from Txn_CxTF_GsL  a "
   tmp = tmp & " inner join Txn_CxTF_Details b on b.SM_Prefix = a.SM_Prefix and b.CxTF_DetailsID  = a.CxTF_DetailsID and b.TxnType = a.TxnType "
   tmp = tmp & " where b.ExchangeTypeID = 2 and b.TxnType = 'D' and (b.Flag <> 'P' or b.flag is null)) "
   tmp = tmp & " order by GM.SM_Prefix "
ElseIf Option3.Value = True Then
   tmp = "SELECT  GM.SM_Prefix 'Location ID',GM.Mst_GSL_ID 'GSL ID.',GM.CxTF_DetailsID 'CDTF No',GM.CDTFDate 'CDTF Date',CMP.CMPName 'CMP', "
   tmp = tmp & " COM.Commodity 'Commodity',CM.DepositoryID 'dp_id',CM.ClientIDRow 'CM_ClientID',CM.ClientName 'Client Name', "
   tmp = tmp & " GoM.GodownNo + '~' + GM.StackNo + '~' + GM.LotNo 'GSL No.', GoM.Address,GM.DepositeBags 'No of Bags',GM.GrossWt 'Weight in MT', "
   tmp = tmp & " CM.Address 'Depositor Address',DM.DepositoryName 'Depositor Name',GM.CRAmount,GM.CRDate,GM.KRNo 'KR No',GM.KRDate 'KR Date',MA.AgentName,GM.CMSPBatchID,GM.DepositeWeight,CMD.LotSize,CMD.Variation,GM.CMSPBatchGeneratedBy,GM.CMSPBatchGeneratedDate,GM.CMSPUpdatedBy,GM.CMSPUpdatedDate"
   ''''added after trf module
   tmp = tmp & " ,TRF.TRFNo, TRF.SamplingDate, CM.CMSPCode, GoM.CommtrackNCDEXCode"
   tmp = tmp & " FROM    Mst_GSL GM Inner Join Mst_Client CM On GM.ClientID=CM.ClientID "
   tmp = tmp & " Inner Join Mst_Depository DM On CM.DepositoryID=DM.DepositoryID Inner Join Mst_Godown GoM On GoM.GodownID=GM.GodownID "
   tmp = tmp & " Inner Join Mst_CMP CMP On CMP.CMPID=GM.CMPID Inner Join Mst_Commodity COM On COM.CommodityID=GM.CommodityID "
   tmp = tmp & " Inner join Mst_CommodityDetail CMD On CMD.CommodityID=GM.CommodityID AND CMD.ExchangeTypeID=GM.ExchangeTypeID "
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On GM.CxTF_DetailsID = TCD.CxTF_DetailsID And GM.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D'"
   tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID"
   ''''added after trf module
   tmp = tmp & " Left Join Txn_TRFDetail TRFD on GM.Mst_GSL_ID = TRFD.Mst_GSL_ID and GM.SM_Prefix = TRFD.SM_Prefix"
   tmp = tmp & " Left Join Txn_TRF TRF on TRFD.TRFID = TRF.TRFID"
   tmp = tmp & " WHERE  COM.CommTrackFlag = '1'  And GM.CMSPBatchID = " & CmbBatch.Text & " "
   If CmbCommodity.Text <> "" Then
      tmp = tmp & " And COM.Commodity = '" & CmbCommodity.Text & "' "
   End If
   If CmbSLocationID.Text <> "" Then
      tmp = tmp & " And GM.SM_Prefix = '" & mSM_Prefix & "' "
   End If
   tmp = tmp & " order by GM.SM_Prefix "
ElseIf Option4.Value = True Then
   tmp = "SELECT  GM.SM_Prefix 'Location ID',GM.Mst_GSL_ID 'GSL ID.',GM.CxTF_DetailsID 'CDTF No',GM.CDTFDate 'CDTF Date',CMP.CMPName 'CMP', "
   tmp = tmp & " COM.Commodity 'Commodity',CM.DepositoryID 'dp_id',CM.ClientIDRow 'CM_ClientID',CM.ClientName 'Client Name', "
   tmp = tmp & " GoM.GodownNo + '~' + GM.StackNo + '~' + GM.LotNo 'GSL No.', GoM.Address,GM.DepositeBags 'No of Bags',GM.GrossWt 'Weight in MT',"
   tmp = tmp & " CM.Address 'Depositor Address',DM.DepositoryName 'Depositor Name',GM.CRAmount,GM.CRDate,GM.KRNo 'KR No',GM.KRDate 'KR Date',MA.AgentName,GM.CMSPBatchID,GM.DepositeWeight,CMD.LotSize,CMD.Variation,GM.CMSPBatchGeneratedBy,GM.CMSPBatchGeneratedDate,GM.CMSPUpdatedBy,GM.CMSPUpdatedDate"
   ''''added after trf module
   tmp = tmp & " ,TRF.TRFNo, TRF.SamplingDate, CM.CMSPCode, GoM.CommtrackNCDEXCode,GM.MoisturePer"
   
   If ChkMAW.Value = vbChecked Then
       tmp = tmp & " ,GM.MAWWt"
   End If
   
   tmp = tmp & " FROM Mst_GSL GM Inner Join Mst_Client CM On GM.ClientID=CM.ClientID "
   tmp = tmp & " Inner Join Mst_Depository DM On CM.DepositoryID=DM.DepositoryID Inner Join Mst_Godown GoM On GoM.GodownID=GM.GodownID "
   tmp = tmp & " Inner Join Mst_CMP CMP On CMP.CMPID=GM.CMPID Inner Join Mst_Commodity COM On COM.CommodityID=GM.CommodityID "
   
   If ChkMAW.Value = vbChecked Then
    tmp = tmp & " AND COM.MAWFlag = 1 "
   End If
   
   tmp = tmp & " Inner join Mst_CommodityDetail CMD On CMD.CommodityID=GM.CommodityID AND CMD.ExchangeTypeID=GM.ExchangeTypeID "
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On GM.CxTF_DetailsID = TCD.CxTF_DetailsID And GM.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D'"
   tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID"
   ''''added after trf module
   tmp = tmp & " Left Join Txn_TRFDetail TRFD on GM.Mst_GSL_ID = TRFD.Mst_GSL_ID and GM.SM_Prefix = TRFD.SM_Prefix"
   tmp = tmp & " Left Join Txn_TRF TRF on TRFD.TRFID = TRF.TRFID"
   tmp = tmp & " WHERE  GM.CMSPBatchID = 0 And COM.CommTrackFlag = '1' And GM.ExchangeTypeID=2 AND (GM.ISINID='0' OR GM.ISINID='' OR GM.ISINID IS NULL) AND (GM.Status ='O'  ) AND (KRNo=0 OR KRNo IS NULL) "
   If CmbBatch.Text <> "" Then
      tmp = tmp & " And GM.CMSPBatchID = " & CmbBatch.Text & " "
   End If
   
   If CmbCommodity.Text <> "" Then
      tmp = tmp & " And COM.Commodity = '" & CmbCommodity.Text & "' "
   End If
   If CmbSLocationID.Text <> "" Then
      tmp = tmp & " And GM.SM_Prefix = '" & mSM_Prefix & "' "
   End If
   If LCase(CmbType.Text) = "marketable" Then
'      tmp = tmp & " AND (GM.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
'   ElseIf LCase(CmbType.Text) = "non marketable" Then
'      tmp = tmp & " AND (GM.DepositeWeight<CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
   
      If ChkMAW.Value = 0 Then
         tmp = tmp & " AND (GM.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
      Else
         tmp = tmp & " AND (GM.MAWWt>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
      End If
   End If

   tmp = tmp & " and gm.Mst_GSL_ID not in(select a.Mst_GSL_ID from Txn_CxTF_GsL  a "
   tmp = tmp & " inner join Txn_CxTF_Details b on b.SM_Prefix = a.SM_Prefix and b.CxTF_DetailsID  = a.CxTF_DetailsID and b.TxnType = a.TxnType "
   tmp = tmp & " where b.ExchangeTypeID = 2 and b.TxnType = 'D' and (b.Flag <> 'P' or b.flag is null)) "
   tmp = tmp & " order by GM.SM_Prefix,GM.Mst_GSL_ID "
Else
   tmp = "SELECT  GM.SM_Prefix 'Location ID',GM.Mst_GSL_ID 'GSL ID.',GM.CxTF_DetailsID 'CDTF No',GM.CDTFDate 'CDTF Date',CMP.CMPName 'CMP', "
   tmp = tmp & " COM.Commodity 'Commodity',CM.DepositoryID 'dp_id',CM.ClientIDRow 'CM_ClientID',CM.ClientName 'Client Name', "
         tmp = tmp & " GoM.GodownNo + '~' + GM.StackNo + '~' + GM.LotNo 'GSL No.', GoM.Address,GM.DepositeBags 'No of Bags',GM.GrossWt 'Weight in MT', "
         tmp = tmp & " CM.Address 'Depositor Address',DM.DepositoryName 'Depositor Name',GM.CRAmount,GM.CRDate,GM.KRNo 'KR No',GM.KRDate 'KR Date',MA.AgentName,GM.CMSPBatchID,GM.DepositeWeight,CMD.LotSize,CMD.Variation,GM.CMSPBatchGeneratedBy,GM.CMSPBatchGeneratedDate,GM.CMSPUpdatedBy,GM.CMSPUpdatedDate,GM.CMSPUpdatedBy,GM.CMSPUpdatedDate "
         ''''added after trf module
         tmp = tmp & " ,TRF.TRFNo, TRF.SamplingDate, CM.CMSPCode, GoM.CommtrackNCDEXCode,GM.MoisturePer"
         tmp = tmp & " FROM    Mst_GSL GM Inner Join Mst_Client CM On GM.ClientID=CM.ClientID "
         tmp = tmp & " Inner Join Mst_Depository DM On CM.DepositoryID=DM.DepositoryID Inner Join Mst_Godown GoM On GoM.GodownID=GM.GodownID "
         tmp = tmp & " Inner Join Mst_CMP CMP On CMP.CMPID=GM.CMPID Inner Join Mst_Commodity COM On COM.CommodityID=GM.CommodityID "
         tmp = tmp & " Inner join Mst_CommodityDetail CMD On CMD.CommodityID=GM.CommodityID AND CMD.ExchangeTypeID=GM.ExchangeTypeID "
         tmp = tmp & "  Inner Join Txn_CxTF_Details TCD On GM.CxTF_DetailsID = TCD.CxTF_DetailsID And GM.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D'   "
         tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID "
         ''''added after trf module
         tmp = tmp & " Left Join Txn_TRFDetail TRFD on GM.Mst_GSL_ID = TRFD.Mst_GSL_ID and GM.SM_Prefix = TRFD.SM_Prefix"
         tmp = tmp & " Left Join Txn_TRF TRF on TRFD.TRFID = TRF.TRFID"
         tmp = tmp & "  WHERE  COM.CommTrackFlag = '1' And GM.ExchangeTypeID=2 AND (GM.ISINID='0' OR GM.ISINID='' OR GM.ISINID IS NULL) AND (GM.Status ='O'  ) AND (KRNo >1000) "
         If CmbCommodity.Text <> "" Then
            tmp = tmp & "  And COM.Commodity = '" & CmbCommodity.Text & "' "
         End If
         If CmbSLocationID.Text <> "" Then
            tmp = tmp & " And GM.SM_Prefix = '" & mSM_Prefix & "' "
         End If
         tmp = tmp & "  Order by GM.SM_Prefix,GM.Mst_GSL_ID "
'            "AND (GM.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) order by GM.SM_Prefix "
'Or GM.Status = 'J'
End If


Call TmpTable
Call dispdata
CmdExcel.Enabled = True
lblRecordCount.Caption = TmpRs.RecordCount

End Sub
Public Sub dispdata()
    Dim x As Double
    x = 0
    Dim TotCol As Variant
    If Option1.Value = True Then
       TotCol = 16
    Else
       TotCol = 18
    End If
    
    If TmpRs.RecordCount > 0 Then
       ProgressBar1.Value = 0
       ProgressBar1.Max = TmpRs.RecordCount + 2
       For I = 1 To TmpRs.RecordCount
           x = MSHFlexGrid1.Rows - 1
           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           For C = 0 To TotCol
               MSHFlexGrid1.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           Next
           MSHFlexGrid1.TextMatrix(x, 19) = IIf(IsNull(TmpRs!AgentName), "", TmpRs!AgentName)
           If Option1.Value = 0 Then
              MSHFlexGrid1.TextMatrix(x, 20) = IIf(IsNull(TmpRs!CMSPBatchID), "", TmpRs!CMSPBatchID)
           'Else
           '   MSHFlexGrid1.TextMatrix(x, 20) = IIf(IsNull(TmpRs!CMSPBatchID), "", TmpRs!CMSPBatchID)
           End If
           MSHFlexGrid1.TextMatrix(x, 21) = IIf(IsNull(TmpRs!CMSPBatchGeneratedBy), "", TmpRs!CMSPBatchGeneratedBy)
           MSHFlexGrid1.TextMatrix(x, 22) = IIf(IsNull(TmpRs!CMSPBatchGeneratedDate), "", TmpRs!CMSPBatchGeneratedDate)
           MSHFlexGrid1.TextMatrix(x, 23) = IIf(IsNull(TmpRs!CMSPUpdatedBy), "", TmpRs!CMSPUpdatedBy)
           MSHFlexGrid1.TextMatrix(x, 24) = IIf(IsNull(TmpRs!CMSPUpdatedDate), "", TmpRs!CMSPUpdatedDate)
           MSHFlexGrid1.TextMatrix(x, 25) = IIf(IsNull(TmpRs!TRFNo), "", TmpRs!TRFNo)
           MSHFlexGrid1.TextMatrix(x, 26) = IIf(IsNull(TmpRs!SamplingDate), "", TmpRs!SamplingDate)
           MSHFlexGrid1.TextMatrix(x, 27) = IIf(IsNull(TmpRs!CMSPCode), "", TmpRs!CMSPCode)
           MSHFlexGrid1.TextMatrix(x, 28) = IIf(IsNull(TmpRs!CommtrackNCDEXCode), "", TmpRs!CommtrackNCDEXCode)
           MSHFlexGrid1.TextMatrix(x, 29) = IIf(IsNull(TmpRs!MoisturePer), "", TmpRs!MoisturePer)
           If ChkMAW.Value = vbChecked Then
              MSHFlexGrid1.TextMatrix(x, 30) = IIf(IsNull(TmpRs!MAWWt), "", TmpRs!MAWWt)
           End If
           'Commented on 25 April 2013
           'If Option1.Value = True Then
              If LCase(CmbType.Text) = "" Then
                 If ChkMAW.Value = 0 Then
                    If TmpRs!DepositeWeight < TmpRs!LotSize - (TmpRs!LotSize * TmpRs!Variation / 100) Then 'And TmpRs!DepositeWeight > TmpRs!LotSize + (TmpRs!LotSize * TmpRs!Variation / 100) Then
                       For C = 0 To MSHFlexGrid1.Cols - 1
                           MSHFlexGrid1.Col = C
                           MSHFlexGrid1.row = x
                           MSHFlexGrid1.CellBackColor = vbYellow
                       Next
                    End If
                 Else
                    If TmpRs!MAWWt < TmpRs!LotSize - (TmpRs!LotSize * TmpRs!Variation / 100) Then 'And TmpRs!DepositeWeight > TmpRs!LotSize + (TmpRs!LotSize * TmpRs!Variation / 100) Then
                       For C = 0 To MSHFlexGrid1.Cols - 1
                           MSHFlexGrid1.Col = C
                           MSHFlexGrid1.row = x
                           MSHFlexGrid1.CellBackColor = vbYellow
                       Next
                    End If
                 End If
              End If
           'End If
           If Option1.Value = True Then
              MSHFlexGrid1.row = x 'I + 1
              MSHFlexGrid1.Col = 20
              MSHFlexGrid1.CellFontName = "Wingdings"
              MSHFlexGrid1.CellFontSize = 12
              MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
              MSHFlexGrid1.Text = strUnChecked
           End If
           
           TmpRs.MoveNext
           ProgressBar1.Value = ProgressBar1.Value + 1
        Next
        ProgressBar1.Value = ProgressBar1.Max
    Else
       MsgBox "No Record Found!!!"
    End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   
'   CmdGetReport.Enabled = False
'   CmdExcel.Enabled = False
'   CmdClear.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Option4.Value = True
Call Grid_Head
CmdExcel.Enabled = False
ProgressBar1.Value = 0

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 800
If ChkMAW.Value = vbChecked Then
    MSHFlexGrid1.Cols = 31
Else
    MSHFlexGrid1.Cols = 30
End If

'If Option1.Value = True Then
'   MSHFlexGrid1.Cols = 31 '29 '25 '21
'Else
'   MSHFlexGrid1.Cols = 20
'End If

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 1) = "GSL ID."
MSHFlexGrid1.TextMatrix(0, 2) = "CDTF No"
MSHFlexGrid1.TextMatrix(0, 3) = "CDTF Date"
MSHFlexGrid1.TextMatrix(0, 4) = "CMP"
MSHFlexGrid1.TextMatrix(0, 5) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 6) = "dp_id"
MSHFlexGrid1.TextMatrix(0, 7) = "CM_ClientID"
MSHFlexGrid1.TextMatrix(0, 8) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 9) = "GSL No."
MSHFlexGrid1.TextMatrix(0, 10) = "Godown Address"
MSHFlexGrid1.TextMatrix(0, 11) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 12) = "Weight in MT"
MSHFlexGrid1.TextMatrix(0, 13) = "Depositor Address"
MSHFlexGrid1.TextMatrix(0, 14) = "Depositor Name"
MSHFlexGrid1.TextMatrix(0, 15) = "Amount Received"
MSHFlexGrid1.TextMatrix(0, 16) = "SC upto Date"
MSHFlexGrid1.TextMatrix(0, 17) = "CMS-P Code" '"KR No"
MSHFlexGrid1.TextMatrix(0, 18) = "CMS-P Date" '"KR Date"
MSHFlexGrid1.TextMatrix(0, 19) = "CDTF Agent Name"
If Option1.Value = True Then
   MSHFlexGrid1.TextMatrix(0, 20) = "Select GSL"
Else
   MSHFlexGrid1.TextMatrix(0, 20) = "CMS-P Batch ID"
End If
MSHFlexGrid1.TextMatrix(0, 21) = "CMS-P Generated By"
MSHFlexGrid1.TextMatrix(0, 22) = "CMS-P Generated Date"
MSHFlexGrid1.TextMatrix(0, 23) = "CMS-P Updated By"
MSHFlexGrid1.TextMatrix(0, 24) = "CMS-P Updated Date"
''added after trf module
MSHFlexGrid1.TextMatrix(0, 25) = "TRF No"
MSHFlexGrid1.TextMatrix(0, 26) = "TRF Date"
MSHFlexGrid1.TextMatrix(0, 27) = "Client CMSP Code"
MSHFlexGrid1.TextMatrix(0, 28) = "NCDEX COMMTRACK CODE"
MSHFlexGrid1.TextMatrix(0, 29) = "Moisture (in %)"

If ChkMAW.Value = vbChecked Then
   MSHFlexGrid1.TextMatrix(0, 30) = "MAW Weight in MT"
End If


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 800
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 2000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 2200
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 3000
MSHFlexGrid1.ColWidth(13) = 1500
MSHFlexGrid1.ColWidth(14) = 1000
MSHFlexGrid1.ColWidth(15) = 1600
MSHFlexGrid1.ColWidth(16) = 1600
MSHFlexGrid1.ColWidth(17) = 1000
MSHFlexGrid1.ColWidth(18) = 1000
MSHFlexGrid1.ColWidth(19) = 2000
MSHFlexGrid1.ColWidth(28) = 3500
MSHFlexGrid1.ColWidth(29) = 1000


End Sub

Private Sub CmdExcel_Click()
'Call Gen_Create_Xls("Rpt_KR Pending.xls")

'If ChkUpdateBatch.Value = 1 Then
'   MaxID = GetMaxCMSPBatchID
'   For i = 0 To MSHFlexGrid1.Rows - 1
'       'mRow = I + 1 ' + 6
'       If MSHFlexGrid1.TextMatrix(i, 20) = strChecked Then
'          tmp = "Update Mst_GSL Set CMSPBatchID = " & MaxID & ",CMSPBatchGeneratedBy = '" & Gen_UserLoginID & "',CMSPBatchGeneratedDate = getdate() Where GSLID = " & MSHFlexGrid1.TextMatrix(i, 1) & " And SM_Prefix = '" & MSHFlexGrid1.TextMatrix(i, 0) & "' And CMSPBatchID = 0 "
'          Call TmpTable
'       End If
'   Next
'End If

Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CMSPPending.xls")
Call Xls_Detail_Rpt

'If ChkUpdateBatch.Value = 1 Then
'   MaxID = GetMaxCMSPBatchID
'   For I = 0 To MSHFlexGrid1.Rows - 1
'       'mRow = I + 1 ' + 6
'       If MSHFlexGrid1.TextMatrix(I, 20) = strChecked Then
'          tmp = "Update Mst_GSL Set CMSPBatchID = " & MaxID & " Where GSLID = " & MSHFlexGrid1.TextMatrix(I, 1) & " And SM_Prefix = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
'          Call TmpTable
'       End If
'   Next
'End If

ChkUpdateBatch.Value = 0
Call Grid_Head

End Sub
Private Function GetMaxCMSPBatchID()
tmp = "select Max(CMSPBatchID) from Mst_GSL"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxCMSPBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxCMSPBatchID = 1
End If
End Function

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CMSPPending.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = 1
'If ChkUpdateBatch.Value = 0 Then
   ProgressBar1.Max = MSHFlexGrid1.Rows + 2
   For I = 0 To MSHFlexGrid1.Rows - 1
       mRow = I + 1 ' + 6
       For C = 0 To MSHFlexGrid1.Cols - 1
           If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "mm/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
      
           MSHFlexGrid1.row = I
           MSHFlexGrid1.Col = C
           If MSHFlexGrid1.CellBackColor = vbYellow Then
              oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
           End If
       Next
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
'ElseIf ChkUpdateBatch.Value = 1 Then
'   ProgressBar1.Max = MSHFlexGrid1.Rows + 2
'   For i = 0 To MSHFlexGrid1.Rows - 1
'
'       If i = 0 Then
'          mRow = 1 ' + 6
 '         For c = 0 To MSHFlexGrid1.Cols - 1
'              If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
'                 If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
'                    oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "mm/dd/yyyy")
'                 Else
'                    oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
'                 End If
'              Else
'                 oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
'              End If
'              MSHFlexGrid1.Row = i
'              MSHFlexGrid1.Col = c
'              If MSHFlexGrid1.CellBackColor = vbYellow Then
'                 oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbYellow
'              End If
'          Next
'       ElseIf MSHFlexGrid1.TextMatrix(i, 20) = strChecked Then
'          mRow = mRow + 1 ' + 6
'          For c = 0 To MSHFlexGrid1.Cols - 1
'              If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
'                 If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
'                    oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "mm/dd/yyyy")
'                 Else
'                    oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
'                 End If
'              Else
'                 oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
'              End If
'              If c = 20 Then
'                 oWS.Cells(1, c + 1) = "CMS-P Batch ID"
'                 oWS.Cells(mRow, c + 1) = MaxID
'              End If
'              MSHFlexGrid1.Row = i
'              MSHFlexGrid1.Col = c
'              If MSHFlexGrid1.CellBackColor = vbYellow Then
'                 oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbYellow
'              End If
'          Next
'       End If
'       ProgressBar1.Value = ProgressBar1.Value + 1
'   Next
'End If

ProgressBar1.Value = ProgressBar1.Max

oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Public Sub Create_Xls(FileName As String)
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FLName = FileName
oBook.SaveAs (Pat & "\excel\" & FileName)
 
Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)
 
oExcel.Quit
'oExcel.Close
End Sub

Private Sub MSHFlexGrid1_Click()
If Option1.Value = True Then
   With MSHFlexGrid1
        If .Col = 20 Then
           If .TextMatrix(.row, .Col) = "" Then Exit Sub
           .Col = 20
           If .TextMatrix(.row, .Col) = strUnChecked Then
              .TextMatrix(.row, .Col) = strChecked
           Else
              .TextMatrix(.row, .Col) = strUnChecked
           End If
        Else
           .Col = 20
           .TextMatrix(.row, .Col) = strUnChecked
        End If

   End With
End If
End Sub

Private Sub Option1_Click()
CmbBatch.Text = ""
CmbBatch.Enabled = False
lblRecordCount.Caption = ""
If Option1.Value = True Then
   CmbType.Enabled = True
   ChkUpdateBatch.Value = 0
   ChkUpdateBatch.Enabled = True
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = True
Else
   CmbType.Enabled = False
   CmbType.Text = ""
   ChkUpdateBatch.Value = 0
   ChkUpdateBatch.Enabled = False
   ChkCheckAll.Value = 0
   ChkCheckAll.Enabled = False
End If
Call Grid_Head
End Sub
Private Sub Option2_Click()
Call Option1_Click
End Sub
Private Sub Option3_Click()
lblRecordCount.Caption = ""
CmbBatch.Text = ""
CmbBatch.Enabled = True
'CmbType.Text = ""
'CmbType.Enabled = False
ChkUpdateBatch.Value = 0
ChkUpdateBatch.Enabled = False
ChkCheckAll.Enabled = False
Call Grid_Head

End Sub

Private Sub Option4_Click()
Call Option3_Click
End Sub
