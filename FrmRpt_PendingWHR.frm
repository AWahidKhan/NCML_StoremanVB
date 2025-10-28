VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_PendingWHR 
   Caption         =   "Pending WHR Report"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   14895
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   350
         Left            =   13200
         TabIndex        =   18
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox CmbSClientName 
         Height          =   315
         Left            =   8250
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   435
         Width           =   3645
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   2820
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   435
         Width           =   1605
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   4515
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   435
         Width           =   3645
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12195
         TabIndex        =   3
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox CmbFlag 
         Height          =   315
         ItemData        =   "FrmRpt_PendingWHR.frx":0000
         Left            =   12000
         List            =   "FrmRpt_PendingWHR.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   435
         Width           =   1500
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   350
         Left            =   11160
         TabIndex        =   1
         Top             =   960
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker TxtWHRFromDate 
         Height          =   345
         Left            =   105
         TabIndex        =   4
         Top             =   435
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         _Version        =   393216
         Format          =   113508353
         CurrentDate     =   40012
      End
      Begin MSComCtl2.DTPicker TxtWHRToDate 
         Height          =   345
         Left            =   1500
         TabIndex        =   5
         Top             =   435
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
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
         Format          =   113508353
         CurrentDate     =   40012
      End
      Begin VB.Label Label1 
         Caption         =   "Client Name"
         Height          =   300
         Left            =   9585
         TabIndex        =   17
         Top             =   195
         Width           =   960
      End
      Begin VB.Label Label36 
         Caption         =   "Location ID"
         Height          =   300
         Left            =   3142
         TabIndex        =   15
         Top             =   195
         Width           =   960
      End
      Begin VB.Label Label42 
         Caption         =   "CMP"
         Height          =   300
         Left            =   6097
         TabIndex        =   14
         Top             =   195
         Width           =   480
      End
      Begin VB.Label LblRecordCount 
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
         Left            =   7920
         TabIndex        =   11
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
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
         Left            =   14160
         TabIndex        =   10
         Top             =   240
         Width           =   75
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
         Left            =   5520
         TabIndex        =   9
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label24 
         Caption         =   "Status"
         Height          =   240
         Left            =   12465
         TabIndex        =   8
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lblWHRDateFrom 
         Caption         =   "From WHR Date"
         Height          =   225
         Left            =   135
         TabIndex        =   7
         Top             =   195
         Width           =   1260
      End
      Begin VB.Label LblWHRToDate 
         Caption         =   "To WHR Date"
         Height          =   240
         Left            =   1590
         TabIndex        =   6
         Top             =   195
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7545
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   13309
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_PendingWHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSLocationID As Double
Dim mFlag As Variant

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then
   mFlag = Null
   Exit Sub
End If

If LCase(CmbFlag.Text) = "pending print" Then
   mFlag = Null
ElseIf LCase(CmbFlag.Text) = "printed" Then
   mFlag = "P"
End If

End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

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
    CmbPreFix.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
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
End If

mSLocationID = RsMst_Location!LocationID

End Sub
Private Sub CmbSClientName_GotFocus()
tmp = CmbSClientName.Text

Call TableMst_Client("Where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
CmbSClientName.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClientName.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClientName.Text = tmp
End Sub
Private Sub CmbSClientName_LostFocus()
If CmbSClientName.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClientName.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClientName.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmdClear_Click()
TxtWHRFromDate.Value = Date - 365
TxtWHRToDate.Value = Date

CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSClientName.Text = ""
CmdExcel.Enabled = False
LblRecordCount.Caption = ""
Call Grid_Head

End Sub

Private Sub CmdGo_Click()

Dim wc As String
wc = ""
tmp = ""
Call Grid_Head
' --DeliveredSpillageBags,  --,--DeliveredSpillageQty

If TxtWHRFromDate.Value > TxtWHRToDate.Value Then
   MsgBox "'From Date' Must be less than 'To Date' !!!"
   TxtWHRFromDate.SetFocus
   Exit Sub
End If

tmp = "Select TWD.WHRNo,TWD.WHRDate,TWD.CMPID,ML.LocationName,MC.CMPName,MG.GodownNo,MClient.ClientName,TWD.WHRNoOfBags, "
tmp = tmp & " TWD.DeliveredNoOfBags , "
tmp = tmp & " TWD.WHRQuantity,TWD.DeliveredQuantity,TWD.Flag   "
tmp = tmp & " From Txn_WHR_Detail TWD "
tmp = tmp & " Inner Join Mst_CMP MC On TWD.CMPID = MC.CMPID "
tmp = tmp & " Inner Join Mst_Godown MG On TWD.GodownID = MG.GodownID "
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID "
tmp = tmp & " Inner Join Mst_Client MClient On TWD.ClientIDRow = MClient.ClientIDRow "
If Gen_DBType = "MDB" Then
   wc = " Where TWD.WHRDate Between #" & Format(TxtWHRFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtWHRToDate.Value, Gen_DatFormat) & "# And ((TWD.WHRNoOfBags - TWD.DeliveredNoOfBags) > 0 or (TWD.WHRQuantity - TWD.DeliveredQuantity) > 0) "
Else
   wc = " Where TWD.WHRDate Between '" & Format(TxtWHRFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtWHRToDate.Value, Gen_DatFormat) & "' And ((TWD.WHRNoOfBags - TWD.DeliveredNoOfBags) > 0 or (TWD.WHRQuantity - TWD.DeliveredQuantity) > 0) "
End If
'And ((TWD.WHRNoOfBags - TWD.DeliveredNoOfBags) > 0 OR (TWD.WHRQuantity - TWD.DeliveredQuantity) > 0)
Dim tmp1 As String
If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, CmbPreFix.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, CmbSCmp.Text, "MC.CmpName", "S", "N")
End If

If CmbSClientName.Text <> "" Then
   wc = GenWc(wc, CmbSClientName.Text, "MClient.ClientName", "S", "N")
End If

If CmbFlag.Text <> "" Then
   'Wc = GenWc(Wc, mFlag, "TWD.Flag", "S", "N")
   If mFlag <> "" Then
      wc = wc & " And TWD.Flag = 'P' "
   Else
      wc = wc & " And TWD.Flag is Null "
   End If
End If

tmp1 = "Order By ML.LocationName"

tmp = tmp & " " & wc & " " & tmp1
Call TmpTable
Call Grid_Head
If TmpRs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
    For I = 1 To TmpRs.RecordCount
        'IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) & "-" & TmpRs!WHRNo - (IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) * 1000000)
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) & "-" & TmpRs!WHRNo - (IIf(IsNull(TmpRs!CMPID), 0, TmpRs!CMPID) * 1000000) 'TmpRs!WHRNo
        MSHFlexGrid1.TextMatrix(I, 1) = Format(TmpRs!WHRDate, "dd-mmm-yyyy")
        MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!LocationName
        MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!CMPName
        MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!GodownNo
        MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!ClientName
        MSHFlexGrid1.TextMatrix(I, 6) = TmpRs!WHRNoOfBags
        MSHFlexGrid1.TextMatrix(I, 7) = TmpRs!DeliveredNoOfBags
        MSHFlexGrid1.TextMatrix(I, 8) = ""
        MSHFlexGrid1.TextMatrix(I, 9) = Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) - Val(MSHFlexGrid1.TextMatrix(I, 7)), "########0") 'TmpRs!WHRNo  'Balance Bags
        MSHFlexGrid1.TextMatrix(I, 10) = TmpRs!WHRQuantity
        MSHFlexGrid1.TextMatrix(I, 11) = TmpRs!DeliveredQuantity
        MSHFlexGrid1.TextMatrix(I, 12) = "" 'TmpRs!WHRNo
        MSHFlexGrid1.TextMatrix(I, 13) = Format(Val(MSHFlexGrid1.TextMatrix(I, 10)) - Val(MSHFlexGrid1.TextMatrix(I, 11)), "########0.00") 'TmpRs!WHRNo 'Balance Qty
        MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
        If MSHFlexGrid1.TextMatrix(I, 14) = "" Then
           MSHFlexGrid1.TextMatrix(I, 14) = "Pending"
        Else
           MSHFlexGrid1.TextMatrix(I, 14) = "Printed"
        End If
        TmpRs.MoveNext
    Next

End If
LblRecordCount.Refresh
LblRecordCount.Caption = TmpRs.RecordCount
CmdExcel.Enabled = True

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
Frame1.Enabled = True
Call CmdClear_Click

End Sub

Public Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 15
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "WHR No"
MSHFlexGrid1.TextMatrix(0, 1) = "WHR Date "
MSHFlexGrid1.TextMatrix(0, 2) = "Location "
MSHFlexGrid1.TextMatrix(0, 3) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 5) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 6) = "WHR Bags"
MSHFlexGrid1.TextMatrix(0, 7) = "Delivered Bags"
MSHFlexGrid1.TextMatrix(0, 8) = "Delivered Spillage Bags"
MSHFlexGrid1.TextMatrix(0, 9) = "Balance Bags"
MSHFlexGrid1.TextMatrix(0, 10) = "WHR Qty"
MSHFlexGrid1.TextMatrix(0, 11) = "Delivered Qty"
MSHFlexGrid1.TextMatrix(0, 12) = "Delivered Spillage Qty"
MSHFlexGrid1.TextMatrix(0, 13) = "Balance Quantity"
MSHFlexGrid1.TextMatrix(0, 14) = "Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1600
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 0  '1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 0 '1000
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 0



MSHFlexGrid1.RowHeight(0) = 700

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_PendingWHR.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_PendingWHR.xls")
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
