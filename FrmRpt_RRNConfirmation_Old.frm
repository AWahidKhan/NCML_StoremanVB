VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_RRNConfirmation 
   Caption         =   "Pending RRN  for Confirmation"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   5
      Top             =   855
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
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   5550
         TabIndex        =   7
         Top             =   285
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
         Left            =   4215
         TabIndex        =   2
         Top             =   285
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
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   285
         Width           =   1335
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
         Left            =   12120
         TabIndex        =   4
         Top             =   360
         Width           =   135
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
         Left            =   9600
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_RRNConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClear_Click()

Call Grid_Head
Label2.Caption = ""

End Sub

Private Sub CmdGetReport_Click()

tmp = "Select MR.RRNID,MR.RRNNo,MR.RRNDate,MR.ISINID,MR.DepositoryID,MC.ClientID,MC.ClientName ,MR.RRNQty,MR.DeliveredQty, " & _
      " MR.CA_Qty,  MR.EntryDate   From Mst_RRN MR " & _
      " Inner Join Mst_Client MC On MR.ClientIDRow = MC.ClientIDRow " & _
      " Where (MR.DeliveredQty - MR.CA_Qty <> 0) and MR.RRNID not in(select a.RRNID from " & _
      " Txn_CxTF_ClientDetail a " & _
      " inner join Txn_CxTF_Details b on b.SM_Prefix = a.SM_Prefix and b.CxTFNo  = a.CxTFNo and b.TxnType = a.TxnType " & _
      " where b.BookedSpaceFlag = 0 and b.ExchangeTypeID = 2 and b.TxnType = 'W' and (b.Flag <> 'P' or b.flag is null)) " & _
      " Order By MR.RRNID "

Call TmpTable

Call Grid_Head
Call DispData

MsgBox ("Done!!!")

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_RRNConfirmation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_RRNConfirmation.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
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
MSHFlexGrid1.Cols = 12


MSHFlexGrid1.TextMatrix(0, 0) = "RRN ID"
MSHFlexGrid1.TextMatrix(0, 1) = "RRN No"
MSHFlexGrid1.TextMatrix(0, 2) = "RRN Date"
MSHFlexGrid1.TextMatrix(0, 3) = "ISIN No"
MSHFlexGrid1.TextMatrix(0, 4) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 6) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 7) = "RRN Qty"
MSHFlexGrid1.TextMatrix(0, 8) = "Delivery Qty"
MSHFlexGrid1.TextMatrix(0, 9) = "CA Qty"
MSHFlexGrid1.TextMatrix(0, 10) = "Difference(DeliveredQty - CA_Qty)"
MSHFlexGrid1.TextMatrix(0, 11) = "BalanceQty(RRNQty - DeliveredQty)"
'MSHFlexGrid1.TextMatrix(0, 12) = "EntryDate"


MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400
'MSHFlexGrid1.ColWidth(12) = 1000

MSHFlexGrid1.RowHeight(0) = 400

End Sub
Public Sub DispData()
Dim x As Double
MsgBox ("Wait till Next Message")
x = 1

If TmpRs.RecordCount > 0 Then
   For i = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(i, 10) = Round(Val(MSHFlexGrid1.TextMatrix(i, 8)) - Val(MSHFlexGrid1.TextMatrix(i, 9)), 3)
       MSHFlexGrid1.TextMatrix(i, 11) = Round(Val(MSHFlexGrid1.TextMatrix(i, 7)) - Val(MSHFlexGrid1.TextMatrix(i, 8)), 3)
       TmpRs.MoveNext
   Next
   Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If

End Sub



