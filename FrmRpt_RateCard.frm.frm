VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_RateCard 
   Caption         =   "Report on Rate Card"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Special Rate Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   90
      TabIndex        =   17
      Top             =   5760
      Width           =   14970
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3495
         Left            =   135
         TabIndex        =   18
         Top             =   270
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   6165
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
   Begin VB.Frame Frame2 
      Caption         =   "Storage Rate Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   90
      TabIndex        =   15
      Top             =   1635
      Width           =   14970
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3735
         Left            =   120
         TabIndex        =   16
         Top             =   255
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   6588
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
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15000
      Begin VB.OptionButton Option1 
         Caption         =   "In Rate Card"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   233
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Not in Rate Card"
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   233
         Width           =   1545
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
         Left            =   13770
         TabIndex        =   19
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchClientName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   7470
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2865
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
         ItemData        =   "FrmRpt_RateCard.frm.frx":0000
         Left            =   8430
         List            =   "FrmRpt_RateCard.frm.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2400
      End
      Begin VB.ComboBox CmbSCommodityID 
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
         Left            =   11790
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2880
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
         Left            =   11760
         TabIndex        =   7
         Top             =   1080
         Width           =   1000
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
         Left            =   12765
         TabIndex        =   8
         Top             =   1080
         Width           =   1000
      End
      Begin VB.ComboBox CmbSBillingUnit 
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
         ItemData        =   "FrmRpt_RateCard.frm.frx":0004
         Left            =   5670
         List            =   "FrmRpt_RateCard.frm.frx":0011
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   645
         Width           =   2880
      End
      Begin VB.ComboBox CmbSBillingCycle 
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
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   645
         Width           =   2880
      End
      Begin VB.Label Label24 
         Caption         =   "Client Name "
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Exchange Type"
         Height          =   240
         Left            =   7200
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         Height          =   210
         Left            =   3480
         TabIndex        =   12
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity"
         Height          =   240
         Left            =   10920
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Billing Unit"
         Height          =   240
         Left            =   4080
         TabIndex        =   10
         Top             =   705
         Width           =   1605
      End
      Begin VB.Label Label16 
         Caption         =   "Billing Cycle"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   960
      End
   End
End
Attribute VB_Name = "FrmRpt_RateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
For i = 1 To RsMst_Location.RecordCount
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
mSLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbsExchangeTypeID_GotFocus()
tmp = CmbSExchangeTypeID.Text
Call TableMst_ExchangeType
CmbSExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
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

Private Sub CmbSBillingCycle_GotFocus()
tmp = CmbSBillingCycle.Text
Call TableMst_BillingCycle
CmbSBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_BillingCycle.RecordCount
   CmbSBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbSBillingCycle.Text = tmp
End Sub

Private Sub CmbSBillingCycle_LostFocus()
If CmbSBillingCycle.Text = "" Then
   mSBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbSBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid selection "
   CmbSBillingCycle.SetFocus
   Exit Sub
End If
mSBillingCycleID = RsMst_BillingCycle!BillingCycleID
End Sub

Private Sub CmbSCommodityID_GotFocus()

tmp = CmbSCommodityID.Text
Call TableMst_CommodityGroup
CmbSCommodityID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_CommodityGroup.RecordCount
   CmbSCommodityID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbSCommodityID.Text = tmp
    
End Sub

'----
Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbSCommodityID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityID.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_CommodityGroup!CommodityGroupID
End Sub

Private Sub CmdClear_Click()

CmbSLocationID.Text = ""
CmbSExchangeTypeID.Text = ""
CmbSCommodityID.Text = ""
CmbSBillingCycle.Text = ""
CmbSBillingUnit.Text = ""
TxtSearchClientName = ""
Option1.Value = True
Call Grid_Head
Call Grid_Head1

End Sub

Private Sub CmdGo_Click()

Dim wc As String
Dim tmp1 As String

If Option2.Value = True Then
   Call Grid_HeadNotPresent
   Call Grid_HeadNotPresentSpl
Else
   Call Grid_Head
   Call Grid_Head1
End If

MsgBox "Wait till next message!!!"

wc = ""

If CmbSLocationID.Text <> "" Then
   wc = GenWc(wc, CmbSLocationID.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp2 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If
If CmbSCommodityID.Text <> "" Then
   wc = GenWc(wc, CmbSCommodityID.Text, "CommodityGroup", "S", "N")
End If
If CmbSExchangeTypeID.Text <> "" Then
   wc = GenWc(wc, CmbSExchangeTypeID.Text, "ExchangeType", "S", "N")
End If
If CmbSBillingCycle.Text <> "" Then
   wc = GenWc(wc, CmbSBillingCycle.Text, "BillingCycle", "S", "N")
End If

If CmbSBillingUnit.Text <> "" Then
   wc = GenWc(wc, Left(CmbSBillingUnit.Text, 1), "MSRC.Flag", "S", "N")
End If

If Option1.Value = True Then

   tmp = "Select ML.LocationName,ME.ExchangeType,MCG.CommodityGroup,MB.BillingCycle,MSRC.Flag,MSRCD.FromBagSize,MSRCD.ToBagSize,"
   tmp = tmp & " MSRCD.RatePerBag , MSRC.RateOnQty, MSRC.DefaultRate,MSRC.CreatedDate,MSRC.CreatedBy,MSRC.UpdatedDate,MSRC.UpdatedBy "
   tmp = tmp & " From Mst_StorageRateCard MSRC"
   tmp = tmp & " Inner Join Mst_Location ML On MSRC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On MSRC.ExchangeTypeID= ME.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MSRC.CommodityGroupID = MCG.CommodityGroupID"
   tmp = tmp & " Inner Join Mst_BillingCycle MB On MSRC.BillingCycleID= MB.BillingCycleID"
   tmp = tmp & " Left Join Mst_StorageRateCardDetail MSRCD On MSRC.SRCID = MSRCD.SRCID"
      
   tmp1 = " Order By ML.LocationName,ME.ExchangeType,MCG.CommodityGroup,MB.BillingCycle,MSRC.Flag"
   
   tmp = tmp & " " & wc & " " & tmp1
   Call TmpTable
      
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
      For i = 1 To TmpRs.RecordCount
       
          MSHFlexGrid1.TextMatrix(i, 0) = TmpRs!LocationName
          MSHFlexGrid1.TextMatrix(i, 1) = TmpRs!ExchangeType
          MSHFlexGrid1.TextMatrix(i, 2) = TmpRs!COmmodityGroup
          MSHFlexGrid1.TextMatrix(i, 3) = TmpRs!BillingCycle
          If LCase(TmpRs!Flag) = "b" Then
             MSHFlexGrid1.TextMatrix(i, 4) = "Bags" 'TmpRs!BillingUnit
          ElseIf LCase(TmpRs!Flag) = "q" Then
             MSHFlexGrid1.TextMatrix(i, 4) = "Quantity" 'TmpRs!BillingUnit
          ElseIf LCase(TmpRs!Flag) = "s" Then
             MSHFlexGrid1.TextMatrix(i, 4) = "Sq Ft." 'TmpRs!BillingUnit
          End If
          MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs!FromBagSize), "", TmpRs!FromBagSize)
          MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRs!ToBagSize), "", TmpRs!ToBagSize)
          If MSHFlexGrid1.TextMatrix(i, 4) = "Bags" Then
             MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(TmpRs!RatePerBag), "", TmpRs!RatePerBag)
          ElseIf MSHFlexGrid1.TextMatrix(i, 4) = "Quantity" Or MSHFlexGrid1.TextMatrix(i, 4) = "Sq Ft." Then
             MSHFlexGrid1.TextMatrix(i, 7) = TmpRs!RateOnQty
          End If
          MSHFlexGrid1.TextMatrix(i, 8) = TmpRs!DefaultRate
          If MSHFlexGrid1.TextMatrix(i, 8) = 0 Then
             MSHFlexGrid1.TextMatrix(i, 8) = "No"
          Else
             MSHFlexGrid1.TextMatrix(i, 8) = "Yes"
          End If
          
          MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "DD-MMM-YYYY"))
          MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy)
          MSHFlexGrid1.TextMatrix(i, 11) = IIf(IsNull(TmpRs!UpdatedDate), "", Format(TmpRs!UpdatedDate, "DD-MMM-YYYY"))
          MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
          
          TmpRs.MoveNext
      Next
   End If
      
      
   tmp = "Select ML.LocationName,ME.ExchangeType,MCG.CommodityGroup,MC.ClientName,MB.BillingCycle,MSPRC.BillingUnit,MSPRCD.FromBagSize,MSPRCD.ToBagSize,"
   tmp = tmp & " MSPRCD.RatePerBag , MSPRC.BaseRate, MSPRC.AbsoluteDiscount, MSPRC.Amount,MSPRC.CreatedDate,MSPRC.CreatedBy,MSPRC.UpdatedDate,MSPRC.UpdatedBy"
   tmp = tmp & " From Mst_SpecialRateCard MSPRC"
   tmp = tmp & " Inner Join Mst_Location ML On MSPRC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On MSPRC.ExchangeTypeID= ME.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MSPRC.CommodityGroupID = MCG.CommodityGroupID"
   tmp = tmp & " Inner Join Mst_BillingCycle MB On MSPRC.BillingCycleID= MB.BillingCycleID"
   tmp = tmp & " Inner Join Mst_Client MC On MSPRC.ClientIDRow = MC.ClientIDRow"
   tmp = tmp & " Left Join Mst_SpecialRateCardDetail MSPRCD On MSPRC.SPRCID = MSPRCD.SPRCID"
   
   tmp1 = " Order By ML.LocationName,ME.ExchangeType,MCG.CommodityGroup,MC.ClientName,MB.BillingCycle,MSPRC.BillingUnit"
   
   
   If TxtSearchClientName <> "" Then
         wc = GenWc(wc, TxtSearchClientName, "MC.ClientName", "S", "L")
   End If
   
   tmp = tmp & " " & wc & " " & tmp1
   
   Call TmpTable
   
   
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
      For i = 1 To TmpRs.RecordCount
       
          MSHFlexGrid2.TextMatrix(i, 0) = TmpRs!LocationName
          MSHFlexGrid2.TextMatrix(i, 1) = TmpRs!ExchangeType
          MSHFlexGrid2.TextMatrix(i, 2) = TmpRs!COmmodityGroup
          MSHFlexGrid2.TextMatrix(i, 3) = TmpRs!ClientName
          MSHFlexGrid2.TextMatrix(i, 4) = TmpRs!BillingCycle
          If LCase(TmpRs!BillingUnit) = "b" Then
             MSHFlexGrid2.TextMatrix(i, 5) = "Bags" 'TmpRs!BillingUnit
          ElseIf LCase(TmpRs!BillingUnit) = "q" Then
             MSHFlexGrid2.TextMatrix(i, 5) = "Quantity" 'TmpRs!BillingUnit
          ElseIf LCase(TmpRs!BillingUnit) = "s" Then
             MSHFlexGrid2.TextMatrix(i, 5) = "Sq Ft." 'TmpRs!BillingUnit
          End If
          MSHFlexGrid2.TextMatrix(i, 6) = IIf(IsNull(TmpRs!FromBagSize), "", TmpRs!FromBagSize)
          MSHFlexGrid2.TextMatrix(i, 7) = IIf(IsNull(TmpRs!ToBagSize), "", TmpRs!ToBagSize)
          If LCase(TmpRs!BaseRate) = "y" Then
             MSHFlexGrid2.TextMatrix(i, 8) = "Yes" 'TmpRs!BaseRate
          ElseIf LCase(TmpRs!BaseRate) = "n" Then
             MSHFlexGrid2.TextMatrix(i, 8) = "No" 'TmpRs!BaseRate
          End If
          MSHFlexGrid2.TextMatrix(i, 9) = TmpRs!AbsoluteDiscount
          
          If MSHFlexGrid2.TextMatrix(i, 5) = "Bags" Then
             MSHFlexGrid2.TextMatrix(i, 10) = TmpRs!RatePerBag
          ElseIf MSHFlexGrid2.TextMatrix(i, 5) = "Quantity" Or MSHFlexGrid2.TextMatrix(i, 5) = "Sq Ft." Then
             MSHFlexGrid2.TextMatrix(i, 10) = TmpRs!Amount
          End If
          
          MSHFlexGrid2.TextMatrix(i, 11) = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "DD-MMM-YYYY"))
          MSHFlexGrid2.TextMatrix(i, 12) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy)
          MSHFlexGrid2.TextMatrix(i, 13) = IIf(IsNull(TmpRs!UpdatedDate), "", Format(TmpRs!UpdatedDate, "DD-MMM-YYYY"))
          MSHFlexGrid2.TextMatrix(i, 14) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
          
          
          
          
          
          TmpRs.MoveNext
      Next
   End If

Else

   tmp = " Select Distinct ML.LocationName,MCG.CommodityGroup"
   tmp = tmp & " From Mst_GSL MGSL"
   tmp = tmp & " Left Join Mst_Godown  MG On MGSL.GodownID = MG.GodownID"
   tmp = tmp & " Left Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Left Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MCom On MGSL.CommodityID = MCom.CommodityID"
   tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MCom.CommodityGroupID = MCG.CommodityGroupID"
   
   tmp1 = ""
   
   If wc = "" Then
      tmp1 = " Where ML.LocationID Not in (Select LocationID from Mst_StorageRateCard) Or MCom.CommodityGroupID Not in (Select CommodityGroupID from Mst_StorageRateCard)"
      tmp1 = tmp1 & " And MGSL.ExchangeTypeID = 1"
   Else
      tmp1 = " And (ML.LocationID Not in (Select LocationID from Mst_StorageRateCard) Or MCom.CommodityGroupID Not in (Select CommodityGroupID from Mst_StorageRateCard))"
      tmp1 = tmp1 & " And MGSL.ExchangeTypeID = 1"
   End If
   
   tmp1 = tmp1 & " Order By ML.LocationName,MCG.CommodityGroup"
   
   tmp = tmp & " " & wc & " " & tmp1
   Call TmpTable
      
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
      For i = 1 To TmpRs.RecordCount
          MSHFlexGrid1.TextMatrix(i, 0) = TmpRs!LocationName
          MSHFlexGrid1.TextMatrix(i, 1) = TmpRs!COmmodityGroup
          TmpRs.MoveNext
      Next
   End If

   tmp = " Select Distinct ML.LocationName,MCG.CommodityGroup"
   tmp = tmp & "  From Mst_GSL MGSL"
   tmp = tmp & "  Left Join Mst_Godown  MG On MGSL.GodownID = MG.GodownID"
   tmp = tmp & " Left Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Left Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MCom On MGSL.CommodityID = MCom.CommodityID"
   tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MCom.CommodityGroupID = MCG.CommodityGroupID"
   tmp1 = ""
   
   If wc = "" Then
      tmp1 = " Where ML.LocationID Not in (Select LocationID from Mst_SpecialRateCard) Or MCom.CommodityGroupID Not in (Select CommodityGroupID from Mst_SpecialRateCard)"
      tmp1 = tmp1 & " And MGSL.ExchangeTypeID = 1"
   Else
      tmp1 = " And (ML.LocationID Not in (Select LocationID from Mst_SpecialRateCard) Or MCom.CommodityGroupID Not in (Select CommodityGroupID from Mst_SpecialRateCard))"
      tmp1 = tmp1 & " And MGSL.ExchangeTypeID = 1"
   End If
   
   tmp1 = tmp1 & " Order By ML.LocationName,MCG.CommodityGroup"
   
   tmp = tmp & " " & wc & " " & tmp1
   
   Call TmpTable
      
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
      For i = 1 To TmpRs.RecordCount
          MSHFlexGrid2.TextMatrix(i, 0) = TmpRs!LocationName
          MSHFlexGrid2.TextMatrix(i, 1) = TmpRs!COmmodityGroup
          TmpRs.MoveNext
      Next
   End If

End If

MsgBox "Done!!!"

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

Call CmdClear_Click
  

End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 13 '9
       
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "Exchange Type"
    .TextMatrix(0, 2) = "Commodity"
    .TextMatrix(0, 3) = "Billing Cycle"
    .TextMatrix(0, 4) = "Billing Unit(Bags/Qty in MT)"
    .TextMatrix(0, 5) = "Unit From"
    .TextMatrix(0, 6) = "Unit To"
    .TextMatrix(0, 7) = "Rate Per Cycle"
    .TextMatrix(0, 8) = "Default Rate"
    
    .TextMatrix(0, 9) = "Created Date"
    .TextMatrix(0, 10) = "Created By"
    .TextMatrix(0, 11) = "Updated Date"
    .TextMatrix(0, 12) = "Updated By"
    
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 0
    .ColWidth(2) = 2500
    .ColWidth(3) = 2000
    .ColWidth(4) = 2500
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
       
    
    .RowHeight(0) = 700
End With
End Sub
Private Sub Grid_HeadNotPresent()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 2
       
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "Commodity"
    
    .ColWidth(0) = 3000
    .ColWidth(1) = 3000
       
    
    .RowHeight(0) = 700
End With

End Sub
Private Sub Grid_HeadNotPresentSpl()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 2
       
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "Commodity"
    
    .ColWidth(0) = 3000
    .ColWidth(1) = 3000
       
    
    .RowHeight(0) = 700
End With

End Sub

Private Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 15
               
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "Exchange Type"
    .TextMatrix(0, 2) = "Commodity"
    .TextMatrix(0, 3) = "Client Name"
    .TextMatrix(0, 4) = "Billing Cycle"
    .TextMatrix(0, 5) = "Billing Unit(Bags/Qty in MT)"
    .TextMatrix(0, 6) = "Unit From"
    .TextMatrix(0, 7) = "Unit To"
    .TextMatrix(0, 8) = "Connected to Base Rate"
    .TextMatrix(0, 9) = "Absolute Discount"
    .TextMatrix(0, 10) = "Rate Per Cycle"
    .TextMatrix(0, 11) = "Created Date"
    .TextMatrix(0, 12) = "Created By"
    .TextMatrix(0, 13) = "Updated Date"
    .TextMatrix(0, 14) = "Updated By"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 0
    .ColWidth(2) = 1500
    .ColWidth(3) = 3000
    .ColWidth(4) = 1500
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
       
    .RowHeight(0) = 700
End With
End Sub




Private Sub CmdExcel_Click()
If MSHFlexGrid1.TextMatrix(1, 0) <> "" Then
   Gen_mTimeStamp = GetTimeStamp
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_RateCard.xls")
   Call Xls_Detail_Rpt
Else
   MsgBox "No Record To Export"
End If
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_RateCard.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Storage Rate Card"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
Next


Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Special Rate Card"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid2.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Option1_Click()

If Option2.Value = True Then
   Call Grid_HeadNotPresent
   Call Grid_HeadNotPresentSpl
Else
   Call Grid_Head
   Call Grid_Head1
End If

End Sub
