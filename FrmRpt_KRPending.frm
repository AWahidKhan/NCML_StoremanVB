VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_KRPending 
   Caption         =   "KR Pending Report"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8535
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   14895
      Begin VB.PictureBox Picturebox1 
         Height          =   8175
         Left            =   120
         ScaleHeight     =   8115
         ScaleWidth      =   14595
         TabIndex        =   5
         Top             =   240
         Width           =   14655
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7815
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   13785
            _Version        =   393216
            FixedCols       =   0
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14895
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
         ItemData        =   "FrmRpt_KRPending.frx":0000
         Left            =   2760
         List            =   "FrmRpt_KRPending.frx":000A
         TabIndex        =   12
         Top             =   285
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DR Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "KR Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
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
         Height          =   495
         Left            =   8055
         TabIndex        =   3
         Top             =   360
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
         Height          =   495
         Left            =   6600
         TabIndex        =   2
         Top             =   360
         Width           =   1455
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
         Height          =   495
         Left            =   9510
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Type"
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
         Left            =   2160
         TabIndex        =   11
         Top             =   345
         Width           =   645
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
         Left            =   11160
         TabIndex        =   10
         Top             =   360
         Width           =   2325
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
         Left            =   13560
         TabIndex        =   9
         Top             =   360
         Width           =   210
      End
   End
End
Attribute VB_Name = "FrmRpt_KRPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Erl() As Integer


Private Sub CmdClear_Click()
Call Grid_Head
CmbType.Text = ""
CmdExcel.Enabled = False
End Sub

Private Sub CmdGetReport_Click()

Dim wc, DispFlag, Filt, Sdat, Edat As Variant
Call Grid_Head

If Option1.Value = True Then
   tmp = "SELECT  GM.SM_Prefix 'Location ID',GM.GSLID 'GSL ID.',GM.CDTFNo 'CDTF No',GM.CDTFDate 'CDTF Date',CMP.CMPName 'CMP', "
   tmp = tmp & " COM.Commodity 'Commodity',CM.DepositoryID 'dp_id',CM.ClientID 'CM_ClientID',CM.ClientName 'Client Name', "
   tmp = tmp & " GoM.GodownNo + '~' + GM.StackNo + '~' + GM.LotNo 'GSL No.', GoM.Address,GM.DepositeBags 'No of Bags',GM.GrossWt 'Weight in MT', "
   tmp = tmp & " CM.Address 'Depositor Address',DM.DepositoryName 'Depositor Name',GM.CRAmount,GM.CRDate,GM.KRNo 'KR No',GM.KRDate 'KR Date',MA.AgentName"
   ''''added after trf module
   tmp = tmp & " ,TRF.TRFNo, TRF.SamplingDate, GoM.NComNCDEXCode,GM.MoisturePer"
   tmp = tmp & " FROM    Mst_GSL GM Inner Join Mst_Client CM On GM.ClientIDRow=CM.ClientIDRow "
   tmp = tmp & " Inner Join Mst_Depository DM On CM.DepositoryID=DM.DepositoryID Inner Join Mst_Godown GoM On GoM.GodownID=GM.GodownID "
   tmp = tmp & " Inner Join Mst_CMP CMP On CMP.CMPID=GM.CMPID Inner Join Mst_Commodity COM On COM.CommodityID=GM.CommodityID "
   tmp = tmp & " Inner join Mst_CommodityDetail CMD On CMD.CommodityID=GM.CommodityID AND CMD.ExchangeTypeID=GM.ExchangeTypeID "
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On GM.CDTFNo = TCD.CxTFNo And GM.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D'"
   tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID"
   ''''added after trf module
   tmp = tmp & " Left Join Txn_TRFDetail TRFD on GM.GSLID = TRFD.GSLID and GM.SM_Prefix = TRFD.SM_Prefix"
   tmp = tmp & " Left Join Txn_TRF TRF on TRFD.TRFID = TRF.TRFID"
   tmp = tmp & " WHERE  COM.CommTrackFlag = '0'  And GM.ExchangeTypeID=2 AND (GM.ISINID='0' OR GM.ISINID='' OR GM.ISINID IS NULL) AND (GM.Status ='O'  ) AND (KRNo=0 OR KRNo IS NULL) "
   If LCase(CmbType.Text) = "marketable" Then
      tmp = tmp & " AND (GM.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
   ElseIf LCase(CmbType.Text) = "non marketable" Then
      tmp = tmp & " AND (GM.DepositeWeight<CMD.LotSize - CMD.LotSize*CMD.Variation/100) "
   End If
   tmp = tmp & " and gm.gslid not in(select a.gslid from Txn_CxTF_GsL  a "
   tmp = tmp & " inner join Txn_CxTF_Details b on b.SM_Prefix = a.SM_Prefix and b.CxTFNo  = a.CxTFNo and b.TxnType = a.TxnType "
   tmp = tmp & " where b.ExchangeTypeID = 2 and b.TxnType = 'D' and (b.Flag <> 'P' or b.flag is null)) "
   tmp = tmp & " order by GM.SM_Prefix "

Else
   tmp = "SELECT  GM.SM_Prefix 'Location ID',GM.GSLID 'GSL ID.',GM.CDTFNo 'CDTF No',GM.CDTFDate 'CDTF Date',CMP.CMPName 'CMP', " & _
         "COM.Commodity 'Commodity',CM.DepositoryID 'dp_id',CM.ClientID 'CM_ClientID',CM.ClientName 'Client Name', " & _
         "GoM.GodownNo + '~' + GM.StackNo + '~' + GM.LotNo 'GSL No.', GoM.Address,GM.DepositeBags 'No of Bags',GM.GrossWt 'Weight in MT', " & _
         "CM.Address 'Depositor Address',DM.DepositoryName 'Depositor Name',GM.CRAmount,GM.CRDate,GM.KRNo 'KR No',GM.KRDate 'KR Date',MA.AgentName " & _
         " ,TRF.TRFNo, TRF.SamplingDate, GoM.NComNCDEXCode ,GM.MoisturePer" & _
         "FROM    Mst_GSL GM Inner Join Mst_Client CM On GM.ClientIDRow=CM.ClientIDRow " & _
         "Inner Join Mst_Depository DM On CM.DepositoryID=DM.DepositoryID Inner Join Mst_Godown GoM On GoM.GodownID=GM.GodownID " & _
         "Inner Join Mst_CMP CMP On CMP.CMPID=GM.CMPID Inner Join Mst_Commodity COM On COM.CommodityID=GM.CommodityID " & _
         "Inner join Mst_CommodityDetail CMD On CMD.CommodityID=GM.CommodityID AND CMD.ExchangeTypeID=GM.ExchangeTypeID " & _
         " Inner Join Txn_CxTF_Details TCD On GM.CDTFNo = TCD.CxTFNo And GM.SM_Prefix = TCD.SM_Prefix And TCD.TxnType = 'D'   " & _
         " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID " & _
         " Left Join Txn_TRFDetail TRFD on GM.GSLID = TRFD.GSLID and GM.SM_Prefix = TRFD.SM_Prefix" & _
         " Left Join Txn_TRF TRF on TRFD.TRFID = TRF.TRFID" & _
         " WHERE  COM.CommTrackFlag = '0'  And GM.ExchangeTypeID=2 AND (GM.ISINID='0' OR GM.ISINID='' OR GM.ISINID IS NULL) AND (GM.Status ='O'  ) AND (KRNo >1000) " & _
         " Order by GM.SM_Prefix "
'            "AND (GM.DepositeWeight>=CMD.LotSize - CMD.LotSize*CMD.Variation/100) order by GM.SM_Prefix "
'Or GM.Flag = 'J'
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
       For I = 1 To TmpRs.RecordCount
           x = MSHFlexGrid1.Rows - 1
           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           For C = 0 To TotCol
               MSHFlexGrid1.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
           Next
           MSHFlexGrid1.TextMatrix(x, 19) = IIf(IsNull(TmpRs!AgentName), "", TmpRs!AgentName)
           'If Option1.Value = True Then
               MSHFlexGrid1.TextMatrix(x, 20) = IIf(IsNull(TmpRs!TRFNo), "", TmpRs!TRFNo)
               MSHFlexGrid1.TextMatrix(x, 21) = IIf(IsNull(TmpRs!SamplingDate), "", TmpRs!SamplingDate)
               MSHFlexGrid1.TextMatrix(x, 22) = IIf(IsNull(TmpRs!NComNCDEXCODE), "", TmpRs!NComNCDEXCODE)
           'End If
               MSHFlexGrid1.TextMatrix(x, 23) = IIf(IsNull(TmpRs!MoisturePer), "", TmpRs!MoisturePer)
           TmpRs.MoveNext
        Next
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
Call Grid_Head
CmdExcel.Enabled = False
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 24 ' 20

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
MSHFlexGrid1.TextMatrix(0, 17) = "KR No"
MSHFlexGrid1.TextMatrix(0, 18) = "KR Date"
MSHFlexGrid1.TextMatrix(0, 19) = "CDTF Agent Name"

'If Option1.Value = True Then
   'MSHFlexGrid1.Cols = 22
   MSHFlexGrid1.TextMatrix(0, 20) = "TRF No"
   MSHFlexGrid1.TextMatrix(0, 21) = "TRF Date"
   MSHFlexGrid1.TextMatrix(0, 22) = "NCOM NCDEX CODE"
   
'End If

MSHFlexGrid1.TextMatrix(0, 23) = "Moisture %"

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
MSHFlexGrid1.ColWidth(22) = 3500
MSHFlexGrid1.ColWidth(23) = 1200

End Sub

Private Sub CmdExcel_Click()
'Call Gen_Create_Xls("Rpt_KR Pending.xls")

Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_Rpt_KR Pending.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_Rpt_KR Pending.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
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
    Next
Next
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

Private Sub Option1_Click()
If Option1.Value = True Then
   CmbType.Enabled = True
Else
   CmbType.Enabled = False
   CmbType.Text = ""
End If
End Sub
Private Sub Option2_Click()
Call Option1_Click
End Sub
