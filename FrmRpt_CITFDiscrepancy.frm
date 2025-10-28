VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_CITFDiscrepancy 
   Caption         =   "CITF Discrepancy"
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
      Height          =   750
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbExchangeType 
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2280
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
         Left            =   8175
         TabIndex        =   2
         Top             =   240
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
         Height          =   375
         Left            =   9390
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CmbLocationID 
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
         TabIndex        =   1
         Top             =   240
         Width           =   2640
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
         Left            =   10605
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
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
         Left            =   12000
         TabIndex        =   9
         Top             =   300
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
         Left            =   14520
         TabIndex        =   8
         Top             =   300
         Width           =   90
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4320
         TabIndex        =   7
         Top             =   300
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8535
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15055
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_CITFDiscrepancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mExchangeTypeID As Variant


Private Sub CmdClear_Click()
CmbLocationID.Text = ""
CmbExchangeType.Text = ""
Call Grid_Head1
LblTotalRecord.Caption = ""
End Sub

Private Sub CmdGetReport_Click()
Dim wc As Variant



If CmbExchangeType.Text = "" Then
   MsgBox "Please Select Exchange Type !!!"
   CmbExchangeType.SetFocus
   Exit Sub
End If


If CmbLocationID.Text <> "" Then
   wc = " And ML.LocationId=" & mLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      wc = " And ML." & Gen_WcLocation
   End If
End If


If mExchangeTypeID = 1 Then
   
   Call Grid_Head1
   tmp = " Select TC.CxTFNo,ML.LocationName,ME.ExchangeType,MC.Commodity,TWD.WHRNo,MCL.ClientName,Count(*) As DoubleRecord,MCMP.CMPID"
   tmp = tmp & " From Txn_Cxtf_ClientDetail TC"
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TC.CxtFNo=TCD.CxTFNo and TC.SM_PreFix=TCD.SM_PreFix And TC.TxnType=TCD.TxnType"
   tmp = tmp & " Inner Join Txn_WHR_Detail  TWD On TWD.WHRNo = TC.RRNID And TWD.SM_PreFix = TC.SM_Prefix"
   tmp = tmp & " Inner Join Mst_Client MCL On MCL.ClientIDRow = TC.ClientIDRow"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCD.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID=TCD.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MCMP.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCD.CommodityID"
   tmp = tmp & " Where TC.TxnType='W'  And  TCD.ExchangeTypeID=1 And TCD.BookedSpaceFlag=0"
   If wc <> "" Then
      tmp = tmp & wc
   End If
   tmp = tmp & " Group By ML.LocationName,MC.Commodity,TC.CxTFNo,TWD.WHRNo,ME.ExchangeType,MCL.ClientName,MCMP.CMPID"
   tmp = tmp & " having Count(*)>1"
   tmp = tmp & " Order by  TC.CxTFNo,TWD.WHRNo"

ElseIf mExchangeTypeID = 2 Then
   
   Call Grid_Head2
   tmp = "Select TC.CxTFNo,ML.LocationName,ME.ExchangeType,MC.Commodity,MR.RRNNo,MR.ISINID,MCL.ClientName,Count(*) As DoubleRecord"
   tmp = tmp & " From Txn_Cxtf_ClientDetail TC"
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TC.CxtFNo=TCD.CxTFNo and TC.SM_PreFix=TCD.SM_PreFix And TC.TxnType=TCD.TxnType"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID=TCD.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MCMP.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCD.CommodityID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCD.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_RRN MR On MR.RRNID=TC.RRNID"
   tmp = tmp & " Inner Join Mst_Client MCL On MCL.ClientIDRow = MR.ClientIDRow"
   tmp = tmp & " Where TC.TxnType='W'  And TCD.ExchangeTypeID=2 And TCD.BookedSpaceFlag=0"
   If wc <> "" Then
      tmp = tmp & wc
   End If
   tmp = tmp & " Group By ML.LocationName,MC.Commodity,TC.CxTFNo,MR.RRNNo,ME.ExchangeType,MR.ISINID,MCL.ClientName"
   tmp = tmp & " having Count(*)>1"
   tmp = tmp & " Order by  TC.CxTFNo,MR.RRNNo"


ElseIf mExchangeTypeID = 3 Then
   Call Grid_Head3

   tmp = " Select TC.CxTFNo,ML.LocationName,ME.ExchangeType,MC.Commodity,TCGD.CDCNo,MCL.ClientName,Count(*) As DoubleRecord"
   tmp = tmp & " From Txn_Cxtf_ClientDetail TC"
   tmp = tmp & " Inner Join Txn_cxtf_Details TCD On TC.CxtFNo=TCD.CxTFNo and TC.SM_PreFix=TCD.SM_PreFix And TC.TxnType=TCD.TxnType"
   tmp = tmp & " Inner Join Txn_CDC_GSL_Detail  TCGD On TCGD.CDCNo=TC.RRNID And TCGD.SM_PreFix=TC.SM_Prefix"
   tmp = tmp & " Inner Join Mst_CMP MCMP On  MCMP.CMPID=TCD.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MCMP.LocationID"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCD.CommodityID"
   tmp = tmp & " Inner Join Mst_Client MCL On MCL.ClientIDRow = TC.ClientIDRow"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCD.ExchangeTypeID"
   tmp = tmp & " Where TC.TxnType='W'  And TCD.BookedSpaceFlag=0  And  TCD.ExchangeTypeID=3"
   If wc <> "" Then
      tmp = tmp & wc
   End If
   tmp = tmp & " Group By ML.LocationName,MC.Commodity,TC.CxTFNo,TCGD.CDCNo,ME.ExchangeType,MCL.ClientName"
   tmp = tmp & " having Count(*)>1"
   tmp = tmp & " Order by  TC.CxTFNo,TCGD.CDCNo"

End If


Call TmpTable
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, c) = TmpRs.Fields(c)
       Next
       If mExchangeTypeID = 1 Then
          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) & "-" & TmpRs!WHRNO - (IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) * 1000000) 'TmpRs!WHRNo
       End If
       TmpRs.MoveNext
   Next
Else
   MsgBox "No Record Found !!"
End If
LblTotalRecord.Caption = TmpRs.RecordCount

End Sub
Private Sub Grid_Head1()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 6
     .AllowUserResizing = flexResizeColumns
     .WordWrap = True
     .TextMatrix(0, 0) = "CITF No"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "Exchange Type"
     .TextMatrix(0, 3) = "Commodity"
     .TextMatrix(0, 4) = "WHR No"
     .TextMatrix(0, 5) = "Client Name"
     
     .ColWidth(0) = 1000
     .ColWidth(1) = 2200
     .ColWidth(2) = 1800
     .ColWidth(3) = 3000
     .ColWidth(4) = 1400
     .ColWidth(5) = 4500
     .Font.Size = 10
     .RowHeight(0) = 400
End With

End Sub

Private Sub Grid_Head2()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 7
     .AllowUserResizing = flexResizeColumns
     .WordWrap = True
     .TextMatrix(0, 0) = "CITF No"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "Exchange Type"
     .TextMatrix(0, 3) = "Commodity"
     .TextMatrix(0, 4) = "RRN No"
     .TextMatrix(0, 5) = "ISIN No"
     .TextMatrix(0, 6) = "Client Name"
     .ColWidth(0) = 1000
     .ColWidth(1) = 2200
     .ColWidth(2) = 1800
     .ColWidth(3) = 3000
     .ColWidth(4) = 1400
     .ColWidth(5) = 1900
     .ColWidth(6) = 3500
     .Font.Size = 10
     .RowHeight(0) = 400
End With

End Sub

Private Sub Grid_Head3()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 6
     .AllowUserResizing = flexResizeColumns
     .WordWrap = True
     .TextMatrix(0, 0) = "CITF No"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "Exchange Type"
     .TextMatrix(0, 3) = "Commodity"
     .TextMatrix(0, 4) = "CDC No"
     .TextMatrix(0, 5) = "Client Name"
     
     .ColWidth(0) = 1000
     .ColWidth(1) = 2200
     .ColWidth(2) = 1800
     .ColWidth(3) = 3000
     .ColWidth(4) = 1400
     .ColWidth(5) = 4500
     .Font.Size = 10
     .RowHeight(0) = 400
End With

End Sub


Private Sub CmbExchangeType_GotFocus()
tmp = CmbExchangeType.Text
Call TableMst_ExchangeType
CmbExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeType.Text = tmp
End Sub

Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   Exit Sub
End If


RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub


Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location("Where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location Selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CITFDiscrepancy.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CITFDiscrepancy.xls")
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

Private Sub Form_Load()
Call Grid_Head1
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
End Sub
