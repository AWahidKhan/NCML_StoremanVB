VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_RRNConfirmation 
   Caption         =   "Pending RRN  for Confirmation"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8895
      Left            =   120
      TabIndex        =   5
      Top             =   615
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8175
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14420
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   8520
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox ChkCommTrack 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2670
         TabIndex        =   10
         Top             =   240
         Width           =   195
      End
      Begin VB.ComboBox CmbCommodity 
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
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   870
         Visible         =   0   'False
         Width           =   885
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
         Left            =   7350
         TabIndex        =   7
         Top             =   165
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
         Left            =   6015
         TabIndex        =   2
         Top             =   165
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
         Left            =   4680
         TabIndex        =   1
         Top             =   165
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Commtrack Flag"
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
         Left            =   255
         TabIndex        =   11
         Top             =   210
         Width           =   2175
      End
      Begin VB.Label Label8 
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
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1080
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
         Left            =   11760
         TabIndex        =   4
         Top             =   210
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
         Left            =   9360
         TabIndex        =   3
         Top             =   210
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_RRNConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityID As Variant

Private Sub ChkCommTrack_Click()

If ChkCommTrack.Value = 1 Then
   Call Grid_Head_CommTrackFormat
Else
   Call Grid_Head
End If

End Sub
Private Sub CmdClear_Click()
Call Grid_Head
Label2.Caption = ""
ChkCommTrack.Value = 0
Call Grid_Head
ProgressBar1.Value = 0
End Sub

Private Sub CmdGetReport_Click()
  
MsgBox "Wait till next message!!!"

If ChkCommTrack.Value = 1 Then
   Call Grid_Head_CommTrackFormat

   tmp = "Select MR.RRNID, MR.RRNNo, MR.RRNDate, MR.ISINID, MR.DepositoryID,MC.ClientIDRow, MC.ClientName,"
   tmp = tmp & " MR.RRNQty, MR.DeliveredQty, MR.OutOfSystemQty, MR.CA_Qty, 0, 0, "
   tmp = tmp & " MRD.ISINID , MGSL.KRNo, MRD.RematQty, Sum(TGSL.DematWeight)"
   tmp = tmp & " From Mst_RRN MR"
   tmp = tmp & " Inner Join Mst_RRN_LotDetail MRD on MRD.RRNID = MR.RRNID"
   tmp = tmp & " Inner Join Mst_GSL as MGSL on  MGSL.ISINID = MRD.ISINID"
   tmp = tmp & " Inner Join Mst_Client MC On MR.ClientID = MC.ClientID and MC.ExchangeTypeID = 2"
   tmp = tmp & " Inner Join Txn_CxTF_GSL  as TGSL on TGSL.Mst_GSL_ID =MGSL.Mst_GSL_ID and TGSL.SM_Prefix = MGSL.SM_Prefix and TGSL.TxnType = 'W'"
   tmp = tmp & " Where "
   tmp = tmp & " MR.CommTrackFlag = 1 and"
   tmp = tmp & " ((MR.DeliveredQty + MR.OutOfSystemQty)  - MR.CA_Qty <> 0) "
   tmp = tmp & " Group By "
   tmp = tmp & " MR.RRNID, MR.RRNNo, MR.RRNDate, MR.ISINID, MR.DepositoryID, MC.ClientIDRow, MC.ClientName,"
   tmp = tmp & " MR.RRNQty, MR.DeliveredQty, MR.OutOfSystemQty, MR.CA_Qty,"
   tmp = tmp & " MRD.ISINID , MGSL.KRNo, MRD.RematQty "
   tmp = tmp & " Order By MR.RRNID"
   
   Call TmpTable
   
   
   tmp = " select a.RRNID from "
   tmp = tmp & " Txn_CxTF_ClientDetail a "
   tmp = tmp & " inner join Txn_CxTF_Details b on b.SM_Prefix = a.SM_Prefix and b.CxTF_DetailsID  = a.CxTF_DetailsID and b.TxnType = a.TxnType "
   tmp = tmp & " where b.BookedSpaceFlag = 0 and b.ExchangeTypeID = 2 and b.TxnType = 'W' and (b.Flag <> 'P' or b.flag is null) "
   
   Call Tmp1Table
   
   tmp = " Select RRNID From Txn_RRNPWR Where Flag <> 'A'"
   
   Call Tmp2Table
   
   Call DisCommTrackData

Else
  Call Grid_Head
   
  tmp = "Select MR.RRNID,MR.RRNNo,MR.RRNDate,MR.ISINID,MR.DepositoryID,MC.ClientIDRow,MC.ClientName ,MR.RRNQty,MR.DeliveredQty, " & _
     " MR.OutOfSystemQty,MR.CA_Qty,  MR.EntryDate   From Mst_RRN MR " & _
     " Inner Join Mst_Client MC On MR.ClientID = MC.ClientID " & _
     " Where ((MR.DeliveredQty + MR.OutOfSystemQty)  - MR.CA_Qty <> 0) and MR.RRNID not in(select a.RRNID from " & _
     " Txn_CxTF_ClientDetail a " & _
     " inner join Txn_CxTF_Details b on b.SM_Prefix = a.SM_Prefix and b.CxTF_DetailsID  = a.CxTF_DetailsID and b.TxnType = a.TxnType " & _
     " where b.BookedSpaceFlag = 0 and b.ExchangeTypeID = 2 and b.TxnType = 'W' and (b.Flag <> 'P' or b.flag is null)) And " & _
     " MR.RRNID not in (Select RRNID From Txn_RRNPWR Where Flag <> 'A')  And (mr.CommTrackFlag = 0 Or mr.CommTrackFlag is null) " & _
     " Order By MR.RRNID "
  Call TmpTable
  Call dispdata

End If
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
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_RRNConfirmation.xls")
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
           If ChkCommTrack.Value = 0 Then
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           Else
              If C = 3 Or C = 13 Or C = 14 Then
                 oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           End If
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

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13


MSHFlexGrid1.TextMatrix(0, 0) = "RRN ID"
MSHFlexGrid1.TextMatrix(0, 1) = "RRN No"
MSHFlexGrid1.TextMatrix(0, 2) = "RRN Date"
MSHFlexGrid1.TextMatrix(0, 3) = "ISIN No"
MSHFlexGrid1.TextMatrix(0, 4) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 6) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 7) = "RRN Qty"
MSHFlexGrid1.TextMatrix(0, 8) = "Delivery Qty"
MSHFlexGrid1.TextMatrix(0, 9) = "Out Of System Qty"
MSHFlexGrid1.TextMatrix(0, 10) = "CA Qty"
MSHFlexGrid1.TextMatrix(0, 11) = "Difference(DeliveredQty + OutOfSystemQty - CA_Qty)"
MSHFlexGrid1.TextMatrix(0, 12) = "BalanceQty(RRNQty - (DeliveredQty + Out Of System Qty))"
'MSHFlexGrid1.TextMatrix(0, 12) = "EntryDate"


MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1100
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 800
MSHFlexGrid1.ColWidth(9) = 800
MSHFlexGrid1.ColWidth(10) = 800
MSHFlexGrid1.ColWidth(11) = 2000
MSHFlexGrid1.ColWidth(12) = 1800
'MSHFlexGrid1.ColWidth(12) = 1000

MSHFlexGrid1.RowHeight(0) = 700

End Sub
Private Sub Grid_Head_CommTrackFormat()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 17


MSHFlexGrid1.TextMatrix(0, 0) = "RRN ID"
MSHFlexGrid1.TextMatrix(0, 1) = "RRN No"
MSHFlexGrid1.TextMatrix(0, 2) = "RRN Date"
MSHFlexGrid1.TextMatrix(0, 3) = "ISIN No"
MSHFlexGrid1.TextMatrix(0, 4) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 5) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 6) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 7) = "RRN Qty"
MSHFlexGrid1.TextMatrix(0, 8) = "Delivery Qty"
MSHFlexGrid1.TextMatrix(0, 9) = "Out Of System Qty"
MSHFlexGrid1.TextMatrix(0, 10) = "CA Qty"
MSHFlexGrid1.TextMatrix(0, 11) = "Difference(DeliveredQty + OutOfSystemQty - CA_Qty)"
MSHFlexGrid1.TextMatrix(0, 12) = "BalanceQty(RRNQty - (DeliveredQty + Out Of System Qty))"
MSHFlexGrid1.TextMatrix(0, 13) = "CMS E-Lot No" '---ISIN ID
MSHFlexGrid1.TextMatrix(0, 14) = "CMS P-Lot No" '----KR No from GSL
MSHFlexGrid1.TextMatrix(0, 15) = "RRN Qty"
MSHFlexGrid1.TextMatrix(0, 16) = "CITF Qty"

MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1100
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 800
MSHFlexGrid1.ColWidth(9) = 800
MSHFlexGrid1.ColWidth(10) = 800
MSHFlexGrid1.ColWidth(11) = 2000
MSHFlexGrid1.ColWidth(12) = 1800
MSHFlexGrid1.ColWidth(13) = 1800
MSHFlexGrid1.ColWidth(14) = 1800
MSHFlexGrid1.ColWidth(15) = 1800
MSHFlexGrid1.ColWidth(16) = 1800
MSHFlexGrid1.RowHeight(0) = 700
End Sub
Public Sub dispdata()
Dim x As Double
MsgBox ("Wait till Next Message")
x = 1
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 2) = Format(MSHFlexGrid1.TextMatrix(I, 2), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 11) = Round((Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 9))) - Val(MSHFlexGrid1.TextMatrix(I, 10)), 3)
       MSHFlexGrid1.TextMatrix(I, 12) = Round(Val(MSHFlexGrid1.TextMatrix(I, 7)) - (Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 9))), 3)
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   ProgressBar1.Value = ProgressBar1.Max
Else
   Call Grid_Head
End If
End Sub

Public Sub DisCommTrackData()
Dim x As Double
Dim mRRNNo As Variant
Dim mRRNinCITF As Boolean
Dim mRRNinPWR As Boolean
'MsgBox ("Wait till Next Message")
x = 1
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 1
If TmpRs.RecordCount > 0 Then
   mRRNNo = ""
   For I = 1 To TmpRs.RecordCount
       'x = MSHFlexGrid1.Rows
       mRRNinCITF = False
       mRRNinPWR = False
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
          TmpRs1.Find "RRNID = " & TmpRs!RRNID & " "
          If TmpRs1.EOF = False Then
             mRRNinCITF = True
          End If
       End If
       If TmpRs2.RecordCount > 0 Then
          TmpRs2.MoveFirst
          TmpRs2.Find "RRNID = " & TmpRs!RRNID & " "
          If TmpRs2.EOF = False Then
             mRRNinPWR = True
          End If
       End If
       If mRRNinCITF = False And mRRNinPWR = False Then
       
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          For C = 0 To TmpRs.Fields.Count - 1
              MSHFlexGrid1.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
          Next
          If mRRNNo <> MSHFlexGrid1.TextMatrix(x, 0) Then
             mRRNNo = MSHFlexGrid1.TextMatrix(x, 0)
          Else
             MSHFlexGrid1.TextMatrix(x, 7) = 0
             MSHFlexGrid1.TextMatrix(x, 8) = 0
             MSHFlexGrid1.TextMatrix(x, 9) = 0
             MSHFlexGrid1.TextMatrix(x, 10) = 0
          End If
          MSHFlexGrid1.TextMatrix(x, 2) = Format(MSHFlexGrid1.TextMatrix(x, 2), "dd-mmm-yyyy")
          MSHFlexGrid1.TextMatrix(x, 11) = Round((Val(MSHFlexGrid1.TextMatrix(x, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 9))) - Val(MSHFlexGrid1.TextMatrix(x, 10)), 3)
          MSHFlexGrid1.TextMatrix(x, 12) = Round(Val(MSHFlexGrid1.TextMatrix(x, 7)) - (Val(MSHFlexGrid1.TextMatrix(x, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 9))), 3)
          x = x + 1
       End If
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
End If

End Sub

