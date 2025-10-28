VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_SCMQualityParameter 
   Caption         =   "SCM Quality Parameter"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9765
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   15165
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   9330
         Width           =   14970
         _ExtentX        =   26405
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.ComboBox CmbQualityParameterID 
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
         ItemData        =   "FrmRpt_SCMQualityParameter.frx":0000
         Left            =   10560
         List            =   "FrmRpt_SCMQualityParameter.frx":0002
         TabIndex        =   12
         Top             =   1680
         Width           =   4515
      End
      Begin VB.TextBox TxtWBSlipNo 
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
         Left            =   6000
         MaxLength       =   100
         TabIndex        =   11
         Top             =   1680
         Width           =   2640
      End
      Begin VB.TextBox TxtTruckNo 
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
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1680
         Width           =   2760
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
         Left            =   13875
         TabIndex        =   29
         Top             =   2100
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Entry Date"
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
         TabIndex        =   28
         Top             =   600
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CDTF Date"
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
         Left            =   1680
         TabIndex        =   27
         Top             =   600
         Width           =   1440
      End
      Begin VB.ComboBox CmbSCommodity 
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
         Left            =   3045
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   3705
      End
      Begin VB.ComboBox CmbSCxTFNo 
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
         ItemData        =   "FrmRpt_SCMQualityParameter.frx":0004
         Left            =   9060
         List            =   "FrmRpt_SCMQualityParameter.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   555
         Width           =   1590
      End
      Begin VB.ComboBox CmbSLocation 
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
         Left            =   6165
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   555
         Width           =   2895
      End
      Begin VB.ComboBox CmbSCmp 
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
         Left            =   10680
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   555
         Width           =   4410
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
         Left            =   12660
         TabIndex        =   14
         Top             =   2100
         Width           =   1215
      End
      Begin VB.CommandButton CmdSearch 
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
         Height          =   375
         Left            =   11445
         TabIndex        =   13
         Top             =   2100
         Width           =   1215
      End
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmRpt_SCMQualityParameter.frx":0008
         Left            =   12120
         List            =   "FrmRpt_SCMQualityParameter.frx":0012
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   2940
      End
      Begin VB.ComboBox CmbSExchangeType 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox CmbAgent 
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
         Left            =   6780
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   5280
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   4710
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   112066561
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   3255
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   112066561
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6735
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   14970
         _ExtentX        =   26405
         _ExtentY        =   11880
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
      Begin VB.Label Label8 
         Caption         =   "Quality Parameter"
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
         Left            =   8880
         TabIndex        =   32
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Weigh Bridge Slip No"
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
         Left            =   3960
         TabIndex        =   31
         Top             =   1725
         Width           =   2130
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   1725
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Agent"
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
         Left            =   9158
         TabIndex        =   26
         Top             =   960
         Width           =   525
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
         Height          =   300
         Left            =   4365
         TabIndex        =   25
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label45 
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
         Left            =   5010
         TabIndex        =   24
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label44 
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
         Height          =   255
         Left            =   3495
         TabIndex        =   23
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label28 
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
         Height          =   255
         Left            =   5340
         TabIndex        =   22
         Top             =   2160
         Width           =   2415
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
         Left            =   7830
         TabIndex        =   21
         Top             =   2167
         Width           =   90
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "CxTF No"
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
         Left            =   9458
         TabIndex        =   20
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label13 
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
         Height          =   270
         Left            =   7200
         TabIndex        =   19
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   12630
         TabIndex        =   18
         Top             =   255
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Mode"
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
         Left            =   1282
         TabIndex        =   17
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   13275
         TabIndex        =   16
         Top             =   960
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmRpt_SCMQualityParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mQualityParameterID, mTruckNo, mWBSLN, mCxTFNo, mSSMPrefix, mSLocationID As Variant

Private Sub CmbQualityParameterID_GotFocus()
tmp = CmbQualityParameterID.Text
CmbQualityParameterID.Clear
Call TableMst_QualityParameter
If RsMst_QualityParameter.RecordCount = 0 Then
   MsgBox "Quality Parameter Not Found."
   Exit Sub
End If
For I = 1 To RsMst_QualityParameter.RecordCount
    CmbQualityParameterID.AddItem RsMst_QualityParameter!QualityParameter
    RsMst_QualityParameter.MoveNext
Next
CmbQualityParameterID.Text = tmp
End Sub

Private Sub CmbQualityParameterID_LostFocus()
If CmbQualityParameterID.Text = "" Then Exit Sub
RsMst_QualityParameter.MoveFirst
RsMst_QualityParameter.Find "QualityParameter='" & CmbQualityParameterID.Text & "'"
If RsMst_QualityParameter.EOF Then
   MsgBox "Invalid Selection."
   Exit Sub
End If
mQualityParameterID = RsMst_QualityParameter!QualityParameterID
End Sub

Private Sub CmdClear_Click()

Call Grid_Head
SFrom.Value = Date - 30
STo.Value = Date
CmbSLocation.Text = ""
CmbSCxTFNo.Text = ""
CmbSCommodity.Text = ""
CmbAgent.Text = ""
CmbSExchangeType.Text = ""
TxtTruckNo = ""
TxtWBSlipNo = ""
CmbQualityParameterID.Text = ""
ProgressBar1.Value = 0

End Sub

Private Sub CmdSearch_Click()

Dim mTruckNo, mWBSLN As Variant
Dim mRow As Integer
Dim Wc1 As String
Wc1 = ""

Call Grid_Head

MsgBox "Wait till next message!!!"

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To Date'!!! "
   SFrom.SetFocus
   Exit Sub
End If


If CmbQualityParameterID.Text = "" Then
   mCol = 23
   Call TableMst_QualityParameter '(" Where StatusFlag='A' ")
   MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + RsMst_QualityParameter.RecordCount
   For I = 1 To RsMst_QualityParameter.RecordCount
       MSHFlexGrid1.ColAlignment(mCol) = flexAlignLeftCenter
       MSHFlexGrid1.TextMatrix(0, mCol) = RsMst_QualityParameter!QualityParameterID
       MSHFlexGrid1.TextMatrix(1, mCol) = 0
       MSHFlexGrid1.TextMatrix(2, mCol) = RsMst_QualityParameter!QualityParameter
       MSHFlexGrid1.ColWidth(mCol) = 1000 'Len(MSHFlexGrid1.TextMatrix(1, mCol)) * 150
       mCol = mCol + 1
       RsMst_QualityParameter.MoveNext
   Next
Else
   MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 1
   MSHFlexGrid1.TextMatrix(0, 23) = mQualityParameterID
   MSHFlexGrid1.TextMatrix(2, 23) = CmbQualityParameterID.Text
   wc = " And MQP.QualityParameterID =" & mQualityParameterID & ""
End If

If Option1.Value = True Then
   If Gen_DBType = "MDB" Then
      Wc1 = " Where TSQP.EntryDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "#"
   Else
      Wc1 = " Where TSQP.EntryDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "'"
   End If
ElseIf Option2.Value = True Then
   If Gen_DBType = "MDB" Then
      Wc1 = " Where TCD.CxTFDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "#"
   Else
      Wc1 = " Where TCD.CxTFDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "'"
   End If
End If

If CmbSLocation.Text <> "" Then
   Wc1 = Wc1 & " And ML.LocationName='" & CmbSLocation.Text & "'"
End If

If CmbSCxTFNo.Text <> "" Then
   Wc1 = Wc1 & " And TSQP.CxTFNO=" & Val(CmbSCxTFNo.Text) & ""
End If

If CmbSCommodity.Text <> "" Then
   Wc1 = Wc1 & " And MCM.Commodity='" & CmbSCommodity.Text & "'"
End If

If CmbSExchangeType.Text <> "" Then
   Wc1 = Wc1 & " And MET.ExchangeType='" & CmbSExchangeType.Text & "'"
End If

If CmbSCmp.Text <> "" Then
   Wc1 = Wc1 & " And MC.CMPName='" & CmbSCmp.Text & "'"
End If

If CmbAgent.Text <> "" Then
   Wc1 = Wc1 & " And MA.AgentName='" & CmbAgent.Text & "'"
End If

If CmbSStatus.Text <> "" Then
   Wc1 = Wc1 & " And TSQP.Status='" & UCase(Left(CmbSStatus.Text, 1)) & "'"
End If

If Trim(TxtTruckNo) <> "" Then
   Wc1 = Wc1 & " And TSQPD.TruckNo like '%" & Trim(TxtTruckNo) & "%'"
End If

If Trim(TxtWBSlipNo) <> "" Then
   Wc1 = Wc1 & " And TSQPD.WeighBridgeSlipno like '%" & Trim(TxtWBSlipNo) & "%'"
End If

If CmbQualityParameterID.Text <> "" Then
   Wc1 = Wc1 & " And MQP.QualityParameter='" & CmbQualityParameterID.Text & "'"
End If

tmp = " Select Distinct MS.StateName,ML.LocationName,TSQP.ID,TSQP.EntryDate,MC.CMPName,MCM.Commodity,TSQP.Remarks,TSQP.CxTFNO,"
tmp = tmp & " TCD.CxTFDate,MET.ExchangeType,MA.AgentName,TSQP.Status,TSQPD.TruckNo,TSQPD.WeighBridgeSlipno,"
tmp = tmp & " TSQPD.TotalWeight,TSQPD.TareWeight,TSQPD.GrossWeight,TSQPD.NoOfBags,"
tmp = tmp & " TSQPD.Remarks 'DetailRemarks',TSQPD.DID,MQP.QualityParameterID,MQP.QualityParameter,TSQPD.Value,MU.EmployeeName +'~'+ MU.EmployeeNo 'CreatedBy',MU1.EmployeeName +'~'+ MU1.EmployeeNo 'UpdatedBy',TSQP.UpdatedDate"
tmp = tmp & " From Txn_ScmQualityParameter TSQP"
tmp = tmp & " Inner Join Mst_Location ML on TSQP.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TSQP.CxTFNo= TCD.CxTFNo And TSQP.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D'"
tmp = tmp & " Inner Join Mst_Cmp MC on TCD.CmpID=MC.CmpID"
tmp = tmp & " Inner Join Mst_Commodity MCM on TCD.CommodityID=MCM.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType MET on TCD.ExchangeTypeID=MET.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Agent MA on TCD.AgentID=MA.AgentID"
tmp = tmp & " Inner Join Txn_ScmQualityParameterDetail TSQPD On TSQP.ID = TSQPD.ID"
tmp = tmp & " Inner Join Mst_QualityParameter MQP On TSQPD.QualityParameterID = MQP.QualityParameterID"
tmp = tmp & " Left Join Mst_Users MU On TSQP.CreatedBy = MU.EmployeeNo"
tmp = tmp & " Left Join Mst_Users MU1 On TSQP.UpdatedBy = MU1.EmployeeNo"
tmp = tmp & Wc1
tmp = tmp & " Order By MS.StateName,ML.LocationName,TSQP.CxTFNo,TSQP.ID,TSQPD.TruckNo,TSQPD.WeighBridgeSlipno"

Call TmpTable


mTruckNo = ""
mWBSLN = ""
mRow = 2
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    If mTruckNo <> TmpRs!TruckNo And mWBSLN <> TmpRs!WeighBridgeSlipno Then
       mRow = mRow + 1
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(mRow, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(mRow, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid1.TextMatrix(mRow, 2) = IIf(IsNull(TmpRs!ID), "", TmpRs!ID)
       MSHFlexGrid1.TextMatrix(mRow, 3) = IIf(IsNull(TmpRs!EntryDate), "", Format(TmpRs!EntryDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(mRow, 4) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
       MSHFlexGrid1.TextMatrix(mRow, 5) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(mRow, 6) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
       MSHFlexGrid1.TextMatrix(mRow, 7) = IIf(IsNull(TmpRs!CxTFNO), "", TmpRs!CxTFNO)
       MSHFlexGrid1.TextMatrix(mRow, 8) = IIf(IsNull(TmpRs!CxTFDate), "", Format(TmpRs!CxTFDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(mRow, 9) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
       MSHFlexGrid1.TextMatrix(mRow, 10) = IIf(IsNull(TmpRs!AgentName), "", TmpRs!AgentName)
       MSHFlexGrid1.TextMatrix(mRow, 11) = IIf(IsNull(TmpRs!Status), "", TmpRs!Status)
       If LCase(MSHFlexGrid1.TextMatrix(mRow, 11)) = "p" Then
          MSHFlexGrid1.TextMatrix(mRow, 11) = "Pending"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(mRow, 11)) = "a" Then
          MSHFlexGrid1.TextMatrix(mRow, 11) = "Approved"
       End If
       MSHFlexGrid1.TextMatrix(mRow, 12) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy)
       MSHFlexGrid1.TextMatrix(mRow, 13) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
       MSHFlexGrid1.TextMatrix(mRow, 14) = IIf(IsNull(TmpRs!UpdatedDate), "", Format(TmpRs!UpdatedDate, "dd-mmm-yyyy hh:mm"))
       MSHFlexGrid1.TextMatrix(mRow, 15) = IIf(IsNull(TmpRs!TruckNo), "", TmpRs!TruckNo)
       MSHFlexGrid1.TextMatrix(mRow, 16) = IIf(IsNull(TmpRs!WeighBridgeSlipno), "", TmpRs!WeighBridgeSlipno)
       MSHFlexGrid1.TextMatrix(mRow, 17) = IIf(IsNull(TmpRs!NoofBags), "", TmpRs!NoofBags)
       MSHFlexGrid1.TextMatrix(mRow, 18) = IIf(IsNull(TmpRs!TotalWeight), "", TmpRs!TotalWeight)
       MSHFlexGrid1.TextMatrix(mRow, 19) = IIf(IsNull(TmpRs!TareWeight), "", TmpRs!TareWeight)
       MSHFlexGrid1.TextMatrix(mRow, 20) = IIf(IsNull(TmpRs!GrossWeight), "", TmpRs!GrossWeight)
       MSHFlexGrid1.TextMatrix(mRow, 21) = IIf(IsNull(TmpRs!DetailRemarks), "", TmpRs!DetailRemarks)
       MSHFlexGrid1.TextMatrix(mRow, 22) = IIf(IsNull(TmpRs!DID), "", TmpRs!DID)
       
       For C = 23 To MSHFlexGrid1.Cols - 1
           If MSHFlexGrid1.TextMatrix(0, C) = TmpRs!QualityParameterID Then
              MSHFlexGrid1.TextMatrix(mRow, C) = TmpRs!Value
              If Val(MSHFlexGrid1.TextMatrix(mRow, C)) < 0 Then
                 MSHFlexGrid1.TextMatrix(mRow, C) = "NA"
              End If
              Exit For
           End If
       Next
       mTruckNo = TmpRs!TruckNo
       mWBSLN = TmpRs!WeighBridgeSlipno
    Else
       For C = 23 To MSHFlexGrid1.Cols - 1
           If MSHFlexGrid1.TextMatrix(0, C) = TmpRs!QualityParameterID Then
              MSHFlexGrid1.TextMatrix(mRow, C) = TmpRs!Value
              If Val(MSHFlexGrid1.TextMatrix(mRow, C)) < 0 Then
                 MSHFlexGrid1.TextMatrix(mRow, C) = "NA"
              End If
              Exit For
           End If
       Next
    End If
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
Label22.Caption = MSHFlexGrid1.Rows - 4
MsgBox "Done"

End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 4
     .FixedRows = 3
     .Cols = 23 '20
     .AllowUserResizing = flexResizeBoth
     
     .TextMatrix(2, 0) = "State Name"
     .TextMatrix(2, 1) = "Location Name"
     .TextMatrix(2, 2) = "ID"
     .TextMatrix(2, 3) = "Entry Date"
     .TextMatrix(2, 4) = "CMP Name"
     .TextMatrix(2, 5) = "Commodity"
     .TextMatrix(2, 6) = "Remarks"
     .TextMatrix(2, 7) = "CDTF No"
     .TextMatrix(2, 8) = "CDTF Date"
     .TextMatrix(2, 9) = "Exchange Type"
     .TextMatrix(2, 10) = "Agent Name"
     .TextMatrix(2, 11) = "Status"
     .TextMatrix(2, 12) = "Created By"
     .TextMatrix(2, 13) = "Update By"
     .TextMatrix(2, 14) = "Updated On"
     .TextMatrix(2, 15) = "Truck No"
     .TextMatrix(2, 16) = "Weigh Bridge Slip No"
     .TextMatrix(2, 17) = "No Of Bags"
     .TextMatrix(2, 18) = "Total Weight"
     .TextMatrix(2, 19) = "Tare weight"
     .TextMatrix(2, 20) = "Gross Weight"
     .TextMatrix(2, 21) = "Details Remarks"
     .TextMatrix(2, 22) = "Detail ID"
     
     .ColWidth(0) = 1200
     .ColWidth(1) = 2000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1600
     .ColWidth(4) = 1600
     .ColWidth(5) = 1600
     .ColWidth(6) = 1400
     
'     .ColAlignment(1) = flexAlignRightCenter
     .RowHeight(0) = 0
     .RowHeight(1) = 300
     .RowHeight(2) = 700
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame0.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame0.Enabled = True
Call CmdClear_Click
End Sub


Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbSLocation.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub


Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   CmbSCxTFNo.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
CmbSCxTFNo.Clear
mSLocationID = RsMst_Location.Fields("LocationID")
mSSMPrefix = RsMst_Location.Fields("SM_Prefix")
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" Then
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub



Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_SCMQualityParameter.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_SCMQualityParameter.xls")
Set oWS = oWB.Worksheets("Sheet1")

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
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
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Agent.RecordCount
    CmbAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbAgent.Text = tmp
End Sub


Private Sub CmbAgent_LostFocus()
If CmbAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection "
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID
End Sub


Private Sub CmbSCmp_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text

If Gen_WcCMP = "" Then
   Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
Else
   Call TableMst_CMP(" Where LocationID = " & mSLocationID & " And " & Gen_WcCMP & "")
End If

CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
CmbSCmp.Text = tmp
End Sub

Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mSCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If

mSCommodityID = RsMst_Commodity!CommodityID
End Sub


Private Sub CmbSCxTFNo_GotFocus()

If CmbSLocation.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbSLocation.SetFocus
   Exit Sub
End If

CmbSCxTFNo.Clear

tmp1 = CmbSCxTFNo.Text
tmp = " Select Distinct CxTFNo "
tmp = tmp & " from Txn_CxTF_Details where TxnType='D' And SCMFlag='1' And SM_Prefix='" & mSSMPrefix & "'"
tmp = tmp & " And CxTFNo In (Select CxTFNo From Txn_ScmQualityParameter) "
Call Tmp1Table

If TmpRs1.RecordCount = 0 Then
   MsgBox "No CxTF No found"
   CmbSCxTFNo.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbSCxTFNo.AddItem TmpRs1!CxTFNO
    TmpRs1.MoveNext
Next
CmbSCxTFNo.Text = tmp1
End Sub


Private Sub CmbSCxTFNo_LostFocus()
If CmbSCxTFNo.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "CxTFNo = '" & CmbSCxTFNo.Text & "'"

If TmpRs1.EOF Then
   MsgBox "Invalid CxTFNo selection "
   CmbSCxTFNo.SetFocus
   Exit Sub
End If
mSCxTFNo = TmpRs1.Fields("CxTFNo")
End Sub

Private Sub CmbSExchangeType_GotFocus()
tmp = CmbSExchangeTypeID.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeType.Text = tmp
End Sub

Private Sub CmbSExchangeType_LostFocus()
If CmbSExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mSExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

