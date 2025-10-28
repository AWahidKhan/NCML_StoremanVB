VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_StorageContractGodMapping 
   Caption         =   "Expired Storage Contract Godown Mapping"
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
      Height          =   9285
      Left            =   60
      TabIndex        =   8
      Top             =   45
      Width           =   15000
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   75
         TabIndex        =   17
         Top             =   8880
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.ComboBox CmbSGodown 
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
         Left            =   6375
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   390
         Width           =   1485
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
         Left            =   3075
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   390
         Width           =   3255
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
         Left            =   12645
         TabIndex        =   6
         Top             =   1320
         Width           =   1125
      End
      Begin VB.TextBox TxtSearchContract 
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
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   13695
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
         Left            =   11520
         TabIndex        =   5
         Top             =   1320
         Width           =   1125
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1125
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   390
         Width           =   2910
      End
      Begin VB.ComboBox CmbSClient 
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
         Left            =   7905
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   7005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7095
         Left            =   75
         TabIndex        =   9
         Top             =   1710
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   12515
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
      Begin VB.Label Label4 
         Caption         =   "Godown"
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
         Left            =   6690
         TabIndex        =   16
         Top             =   120
         Width           =   840
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   4185
         TabIndex        =   15
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Contract Name"
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
         Left            =   120
         TabIndex        =   14
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label15 
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
         Left            =   6915
         TabIndex        =   13
         Top             =   1335
         Width           =   90
      End
      Begin VB.Label Label23 
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
         Left            =   4440
         TabIndex        =   12
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label42 
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
         Height          =   225
         Left            =   1170
         TabIndex        =   11
         Top             =   120
         Width           =   810
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
         Height          =   225
         Left            =   11107
         TabIndex        =   10
         Top             =   120
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmRpt_StorageContractGodMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLOcationID, mCMPID As Variant

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mLOcationID = RsMst_Location!LocationID
End Sub


Private Sub CmbSClient_GotFocus()

tmp = CmbSClient.Text
Call TableMst_Client("Where ExchangeTypeID = " & GetExchangeID("Non NCDEX"))
CmbSClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbSClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClient.Text = tmp


End Sub

Private Sub CmbSClient_LostFocus()
If CmbSClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClient.SetFocus
   Exit Sub
End If

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 10
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 700

MSHFlexGrid1.TextMatrix(0, 0) = "Contract ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Contract Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Extended Period"
MSHFlexGrid1.TextMatrix(0, 3) = "From Date"
MSHFlexGrid1.TextMatrix(0, 4) = "To Date"
MSHFlexGrid1.TextMatrix(0, 5) = "Reservation Qty"
MSHFlexGrid1.TextMatrix(0, 6) = "Mapping Detail ID"
MSHFlexGrid1.TextMatrix(0, 7) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 9) = "Godown Reservation Qty"


MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1800

End Sub

Private Sub CmdClear_Click()
CmbSLocation.Text = ""
CmbCMPName.Text = ""
CmbSClient.Text = ""
CmbSGodown.Text = ""
TxtSearchContract = ""
Call Grid_Head
Label15.Caption = ""

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_StorageContractGodownMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_StorageContractGodownMapping.xls")
Set oWS = oWB.Worksheets("Sheet1")
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
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As String


Call Grid_Head
ProgressBar1.Value = 0

tmp = " Select MSC.ContractID,MSC.ContractName,MSC.ExtendedPeriod,MSBD.FromDate,"
tmp = tmp & " MSBD.ToDate , MSBD.BillingQty, TSGD.SRGodownMapDetailsID,MC.CMPName, MG.GodownNo, TSGD.ReservationQty"
tmp = tmp & " from Mst_StorageContract MSC"
tmp = tmp & " Inner Join Mst_SContractBillingDetail MSBD on MSC.ContractID = MSBD.ContractID"
tmp = tmp & " Inner Join Txn_SRGodownMapping TSG On MSC.ContractID = TSG.ContractID"
tmp = tmp & " Inner Join Txn_SRGodownMapDetails TSGD On TSG.SRGodownMappingID = TSGD.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_Godown MG On TSGD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MCl On MSC.ClientIDRow = MCl.ClientIDRow"
tmp = tmp & " Where MSC.Flag = 'A' And MSC.ExtendedPeriod > '2013-01-23' And MSBD.BillingQty = 0 "

wc = ""

If CmbSLocation.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbSLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If

If CmbSClient.Text <> "" Then
   wc = wc & " And MCl.ClientName = '" & CmbSClient.Text & "'"
End If

If CmbCMPName.Text <> "" Then
   wc = wc & " And MC.CMPName = '" & CmbCMPName.Text & "'"
Else
   If Gen_WcCMP <> "" Then
      wc = wc & " And ML." & Gen_WcCMP
   End If
End If

If CmbSGodown.Text <> "" Then
   wc = wc & " And MG.GodownNo = '" & CmbSGodown.Text & "'"
End If

If TxtSearchContract <> "" Then
   wc = wc & " And MSC.ContractName Like '%" & TxtSearchContract & "%'"
End If

tmp1 = " Order By MSC.ContractID"
tmp = tmp & wc & tmp1

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 5
For i = 1 To TmpRs.RecordCount
    For C = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(i, 2) = Format(MSHFlexGrid1.TextMatrix(i, 2), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(i, 3) = Format(MSHFlexGrid1.TextMatrix(i, 3), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(i, 4) = Format(MSHFlexGrid1.TextMatrix(i, 4), "dd-mmm-yyyy")
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max

Label15.Caption = TmpRs.RecordCount

End Sub


Private Sub CmbCMPName_GotFocus()
If CmbSLocation.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text

If Gen_WcCMP <> "" Then
   qry = " Where " & Gen_WcCMP & " and LocationID =" & mLOcationID & " "
Else
   qry = " Where LocationID =" & mLOcationID & ""
End If
'qry = " Where LocationID =" & mLocationID & " "
Call TableMst_CMP(qry)
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
'----
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
    mCMPID = Null
   Exit Sub
End If
RsMst_CMP.MoveFirst

RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
'dtCurrentDate = RsMst_CMP!CurrentDate 'CDate(Format(RsMst_CMP!CurrentDate, "mm/dd/yyyy"))
'RptDate.Value = RsMst_CMP!CurrentDate
End Sub


Private Sub CmbSGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP"
   CmbCMPName.SetFocus
   Exit Sub
End If
 
tmp = CmbSGodown.Text
Call TableMst_Godown(" Where CMPID = " & mCMPID & "")
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!godownno
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame0.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame0.Enabled = True
Call CmdClear_Click

End Sub
