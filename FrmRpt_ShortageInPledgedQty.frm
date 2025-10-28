VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_ShortageInPledgedQty 
   Caption         =   "Shortage In Pledged Quantity"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8415
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8055
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14208
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmRpt_ShortageInPledgedQty.frx":0000
         Left            =   4920
         List            =   "FrmRpt_ShortageInPledgedQty.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   2400
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
         Left            =   7350
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   2520
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   2550
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2325
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
         TabIndex        =   1
         Top             =   600
         Width           =   2325
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
         Left            =   13470
         TabIndex        =   7
         Top             =   540
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
         Left            =   12135
         TabIndex        =   6
         Top             =   540
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
         Left            =   10800
         TabIndex        =   5
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   5340
         TabIndex        =   15
         Top             =   285
         Width           =   1575
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
         Left            =   8070
         TabIndex        =   14
         Top             =   285
         Width           =   1080
      End
      Begin VB.Label Label3 
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
         Height          =   300
         Left            =   3465
         TabIndex        =   13
         Top             =   285
         Width           =   480
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
         Left            =   795
         TabIndex        =   12
         Top             =   285
         Width           =   1080
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
         Left            =   10800
         TabIndex        =   9
         Top             =   240
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
         Left            =   13200
         TabIndex        =   8
         Top             =   240
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmRpt_ShortageInPledgedQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPName, mCommodityID, mExchangeTypeID As Variant


Private Sub CmdClear_Click()
Call Grid_Head
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbCMPName.Text = ""
CmbExchangeTypeID.Text = ""
CmbCommodityID.Text = ""
LblTotalRecord.Caption = "-"
End Sub

Private Sub CmdExcel_Click()
If MSHFlexGrid1.Rows < 3 And MSHFlexGrid1.TextMatrix(1, 0) = "" Then MsgBox "No Record Found": Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ShortagePledgeQuantity.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ShortagePledgeQuantity.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
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

Call Grid_Head
Call TableTxn_PledgeCreation_GSL
Call TableTxn_PledgeCreation
Call TableMst_Bank

tmp = "Select ML.LocationName,MCMP.CMPName,MG.PledgeNo,"
tmp = tmp & " MC.Commodity,MEX.ExchangeType,"
tmp = tmp & " MGO.GodownNo+'~'+ MG.StackNo+'~'+MG.LotNo 'GSL',"
tmp = tmp & " MG.PledgeBags, MG.BalanceBags ,"
tmp = tmp & " MG.PledgeWeight ,MG.BalanceWeight ,"
tmp = tmp & " MG.GSLID,MG.SM_Prefix,MCMP.CMPID"
tmp = tmp & " From Mst_GSL MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID=MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID=MCMP.LocationID"
tmp = tmp & " Inner Join Mst_Commodity MC on MC.CommodityID =MG.CommodityID"
tmp = tmp & " Left Join Mst_ExchangeType MEX on MEX.ExchangeTypeID=MG.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Godown MGO on MGO.GodownID=MG.GoDownID"
tmp = tmp & " Where (MG.PledgeBags < MG.BalanceBags Or MG.PledgeWeight < MG.BalanceWeight) "
tmp = tmp & " And (MG.PledgeBags <> 0 And MG.PledgeWeight <> 0)"

If CmbLocation.Text <> "" Then
   tmp = tmp & " And ML.LocationID=" & mLocationID & ""
End If
If CmbCMPName.Text <> "" Then
   tmp = tmp & " And MCMP.CMPID=" & mCMPName & ""
End If
If CmbExchangeTypeID.Text <> "" Then
   tmp = tmp & " And MEX.ExchangeTypeID=" & mExchangeTypeID & " "
End If
If CmbCommodityID.Text <> "" Then
   tmp = tmp & " And MC.CommodityID=" & mCommodityID & ""
End If

tmp = tmp & " Order by ML.LocationName,MEX.ExchangeType"

Call TmpTable


If TmpRs.RecordCount > 0 Then
   For mRow = 1 To TmpRs.RecordCount
       
       MSHFlexGrid1.TextMatrix(mRow, 0) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid1.TextMatrix(mRow, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
       MSHFlexGrid1.TextMatrix(mRow, 2) = IIf(IsNull(TmpRs!PledgeNo), "", TmpRs!PledgeNo)
        
       MSHFlexGrid1.TextMatrix(mRow, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(mRow, 5) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
       
       'WHR/ISIN
       
       RsTxn_PledgeCreation_GSL.Filter = "GSLID=" & TmpRs!GSLID & " And SM_Prefix='" & TmpRs!SM_Prefix & "'"
       
       If Not RsTxn_PledgeCreation_GSL.EOF Then
          RsTxn_PledgeCreation_GSL.MoveLast
          RsTxn_PledgeCreation.Filter = "PledgeID=" & RsTxn_PledgeCreation_GSL!PledgeID & " And SM_Prefix='" & RsTxn_PledgeCreation_GSL!SM_Prefix & "'"
          If Not RsTxn_PledgeCreation.EOF Then
             If LCase(TmpRs!ExchangeType) = "non ncdex" Then 'Non NCDEX WHRNO
                'MSHFlexGrid1.TextMatrix(mRow, 6) = IIf(IsNull(RsTxn_PledgeCreation!WHRNO), "", RsTxn_PledgeCreation!WHRNO)
                MSHFlexGrid1.TextMatrix(mRow, 6) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) & "-" & RsTxn_PledgeCreation!WHRNO - (IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) * 1000000)
             ElseIf LCase(TmpRs!ExchangeType) = "ncdex" Then 'NCDEX ISINID
                MSHFlexGrid1.TextMatrix(mRow, 7) = IIf(IsNull(RsTxn_PledgeCreation!WHRNO), "", RsTxn_PledgeCreation!WHRNO)
             ElseIf LCase(TmpRs!ExchangeType) = "spot" Then   'Spot
                MSHFlexGrid1.TextMatrix(mRow, 8) = IIf(IsNull(RsTxn_PledgeCreation!WHRNO), "", RsTxn_PledgeCreation!WHRNO)
             End If
          End If
       End If
       
       'G~S~L
       MSHFlexGrid1.TextMatrix(mRow, 9) = IIf(IsNull(TmpRs!GSL), "", TmpRs!GSL)
       MSHFlexGrid1.TextMatrix(mRow, 10) = IIf(IsNull(TmpRs!PledgeBags), "", TmpRs!PledgeBags)
       MSHFlexGrid1.TextMatrix(mRow, 11) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
       MSHFlexGrid1.TextMatrix(mRow, 12) = Val(MSHFlexGrid1.TextMatrix(mRow, 11)) - Val(MSHFlexGrid1.TextMatrix(mRow, 10))
       ''MSHFlexGrid1.TextMatrix(mRow, 12) = Format(MSHFlexGrid1.TextMatrix(mRow, 12), "########0.00")
       MSHFlexGrid1.TextMatrix(mRow, 13) = IIf(IsNull(TmpRs!PledgeWeight), "", TmpRs!PledgeWeight)
       MSHFlexGrid1.TextMatrix(mRow, 14) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
       MSHFlexGrid1.TextMatrix(mRow, 15) = Val(MSHFlexGrid1.TextMatrix(mRow, 14)) - Val(MSHFlexGrid1.TextMatrix(mRow, 13))
       MSHFlexGrid1.TextMatrix(mRow, 15) = Format(MSHFlexGrid1.TextMatrix(mRow, 15), "########0.000")
       
       If RsMst_Bank.RecordCount > 0 Then
          RsMst_Bank.MoveFirst
       End If
       RsMst_Bank.Find "BankID=" & RsTxn_PledgeCreation!BankID & ""
       
       'Bank Name
       MSHFlexGrid1.TextMatrix(mRow, 3) = IIf(IsNull(RsMst_Bank!BankName), "", RsMst_Bank!BankName)
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       RsTxn_PledgeCreation_GSL.Filter = ""
       RsTxn_PledgeCreation.Filter = ""
       TmpRs.MoveNext
   Next
   CmdExcel.Enabled = True
   MsgBox "Done !!!", vbInformation
Else
   CmdExcel.Enabled = False
   MsgBox " No Record Found "
End If
LblTotalRecord.Caption = TmpRs.RecordCount
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

With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 16
    .Font.Size = 10
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "CMP Name"
    .TextMatrix(0, 2) = "Pledge NO"
    .TextMatrix(0, 3) = "Bank Name"
    .TextMatrix(0, 4) = "Commodity"
    .TextMatrix(0, 5) = "Exchange Type" 'NCDEX / Non - NCDEX
    .TextMatrix(0, 6) = "WHR No Pledged "
    .TextMatrix(0, 7) = "ISIN Pledged"
    .TextMatrix(0, 8) = "CDC No"
    .TextMatrix(0, 9) = "G~S~L"
    .TextMatrix(0, 10) = "Pledged Bags "
    .TextMatrix(0, 11) = "Balance Bags"
    .TextMatrix(0, 12) = "Bags Difference"
    .TextMatrix(0, 13) = "Pledged Quantity"
    .TextMatrix(0, 14) = "Balance Quantity "    'in the GSL
    .TextMatrix(0, 15) = "Quantity Difference"

    .ColWidth(0) = 2000
    .ColWidth(1) = 1800
    .ColWidth(2) = 1200
    .ColWidth(3) = 2500
    .ColWidth(4) = 2000
    .ColWidth(5) = 1500
    .ColWidth(6) = 1400
    .ColWidth(7) = 1200
    .ColWidth(8) = 0
    .ColWidth(9) = 1500
    .ColWidth(10) = 1500
    .ColWidth(11) = 1500
    .ColWidth(12) = 1700
    .ColWidth(13) = 1700
    .ColWidth(14) = 1700
    .ColWidth(15) = 1750
    .RowHeight(0) = 400
End With

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
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   CmbCMPName.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If

If tmp <> CmbLocation.Text Then
   CmbCMPName.Text = ""
End If

mLocationID = RsMst_Location!LocationID

End Sub
Private Sub CmbCMPName_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text
Call TableMst_CMP(" Where LocationID = " & mLocationID & "")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
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
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
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
'----
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

