VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_RRNPending 
   Caption         =   "RRN Pending"
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
      Height          =   825
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   15015
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
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   3045
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
         Left            =   3180
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3240
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
         Left            =   6840
         TabIndex        =   2
         Top             =   353
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
         Height          =   375
         Left            =   8175
         TabIndex        =   3
         Top             =   353
         Width           =   1335
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
         Left            =   9510
         TabIndex        =   4
         Top             =   353
         Width           =   1335
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
         Height          =   240
         Left            =   1222
         TabIndex        =   11
         Top             =   120
         Width           =   840
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
         Height          =   255
         Left            =   4290
         TabIndex        =   10
         Top             =   120
         Width           =   1080
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
         Left            =   11280
         TabIndex        =   9
         Top             =   413
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         Left            =   13800
         TabIndex        =   8
         Top             =   420
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   5
      Top             =   870
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_RRNPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityID, mLocationID As Variant

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

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If


mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmdClear_Click()
CmbLocation.Text = ""
CmbCommodityID.Text = ""
Call Grid_Head
Label2.Caption = ""
'CmbLocation.SetFocus
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
wc = ""
Dim tmp1 As String
tmp1 = ""

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If
If CmbCommodityID.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "MComm.CommodityID", "N", "N")
End If

tmp = "Select MR.RRNID,MR.RRNNo,MR.RRNDate,ML.LocationName,MComm.Commodity,MR.DepositoryID,MC.ClientIDRow,MC.ClientName ,MR.RRNQty,MR.DeliveredQty, "

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
   tmp = tmp & " MR.OutOfSystemQty,MR.CA_Qty,  MR.ISINID "
Else
   tmp = tmp & " MR.OutOfSystemQty,MR.CA_Qty,  MR.EntryDate    "
End If

tmp = tmp & " From Mst_RRN MR "
tmp = tmp & " Inner Join Mst_Client MC On MR.ClientID = MC.ClientID "
tmp = tmp & " Inner Join Mst_ISIN MI On MR.ISINID = MI.ISINID "
tmp = tmp & " Inner Join Mst_Location ML On MI.LocationID = ML.LocationID "
tmp = tmp & " Inner Join Mst_Commodity MComm On MI.CommodityID = MComm.CommodityID "



If wc <> "" Then
   tmp = tmp & wc
   tmp = tmp & " And MR.RRNQty - (MR.DeliveredQty + MR.OutOfSystemQty) <> 0 "
Else
   tmp = tmp & " Where MR.RRNQty - (MR.DeliveredQty + MR.OutOfSystemQty) <> 0 "
End If
tmp = tmp & " Order By MR.RRNID "

Call TmpTable

Call Grid_Head
Call dispdata
Label2.Refresh
Label2.Caption = TmpRs.RecordCount

MsgBox ("Done!!!")

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_RRNPending.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_RRNPending.xls")
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
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click
CmdExcel.Enabled = False

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
   MSHFlexGrid1.Cols = 14
Else
   MSHFlexGrid1.Cols = 13
End If

MSHFlexGrid1.TextMatrix(0, 0) = "RRN ID"
MSHFlexGrid1.TextMatrix(0, 1) = "RRN No"
MSHFlexGrid1.TextMatrix(0, 2) = "RRN Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 5) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 6) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 7) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 8) = "RRN Qty"
MSHFlexGrid1.TextMatrix(0, 9) = "Delivery Qty"
MSHFlexGrid1.TextMatrix(0, 10) = "Delivered before Storeman Date Qty"
MSHFlexGrid1.TextMatrix(0, 11) = "CA Qty"
MSHFlexGrid1.TextMatrix(0, 12) = "BalanceQty(RRNQty - (DeliveredQty + Delivered before Storeman Date Qty))"

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
   MSHFlexGrid1.TextMatrix(0, 13) = "ISIN ID"
End If

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1100
MSHFlexGrid1.ColWidth(5) = 1100
MSHFlexGrid1.ColWidth(6) = 900
MSHFlexGrid1.ColWidth(7) = 2000
MSHFlexGrid1.ColWidth(8) = 800
MSHFlexGrid1.ColWidth(9) = 800
MSHFlexGrid1.ColWidth(10) = 800
MSHFlexGrid1.ColWidth(11) = 800
MSHFlexGrid1.ColWidth(12) = 2800


If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
   MSHFlexGrid1.ColWidth(13) = 2800
End If


'MSHFlexGrid1.ColWidth(12) = 1000

MSHFlexGrid1.RowHeight(0) = 850

End Sub
Public Sub dispdata()
Dim x As Double
MsgBox ("Wait till Next Message")
x = 1

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 2) = Format(MSHFlexGrid1.TextMatrix(I, 2), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 12) = Round(Val(MSHFlexGrid1.TextMatrix(I, 8)) - (Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 10))), 3)
       If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
          MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!ISINID), "", TmpRs!ISINID)
       End If
       TmpRs.MoveNext
   Next
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If

End Sub





