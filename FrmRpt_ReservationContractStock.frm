VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_ReservationContractStock 
   Caption         =   "Reservation Contract Stock Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox Check1 
         Caption         =   "Expired Reservations"
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
         Left            =   3405
         TabIndex        =   19
         Top             =   1020
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Having Zero Stock"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
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
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox TxtClientName 
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
         Left            =   10650
         TabIndex        =   16
         Top             =   450
         Width           =   4110
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
         Left            =   7065
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   450
         Width           =   3435
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
         Height          =   390
         Left            =   12585
         TabIndex        =   4
         Top             =   900
         Width           =   1095
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
         Height          =   390
         Left            =   11475
         TabIndex        =   3
         Top             =   900
         Width           =   1095
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
         Height          =   390
         Left            =   13680
         TabIndex        =   2
         Top             =   900
         Width           =   1095
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
         Left            =   3165
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   450
         Width           =   3765
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   195
         TabIndex        =   6
         Top             =   465
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
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
         Format          =   61079553
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1695
         TabIndex        =   7
         Top             =   465
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
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
         Format          =   61079553
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label3 
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
         Left            =   8055
         TabIndex        =   14
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label LblClient 
         Caption         =   "Client Name"
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
         Left            =   12120
         TabIndex        =   13
         Top             =   165
         Width           =   1290
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
         Left            =   9915
         TabIndex        =   12
         Top             =   975
         Width           =   90
      End
      Begin VB.Label Label4 
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
         Left            =   7950
         TabIndex        =   11
         Top             =   975
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "To  Date"
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
         Left            =   1995
         TabIndex        =   10
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "From  Date"
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
         Left            =   510
         TabIndex        =   9
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   4455
         TabIndex        =   8
         Top             =   150
         Width           =   840
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8295
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14631
      _Version        =   393216
      MergeCells      =   1
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
Attribute VB_Name = "FrmRpt_ReservationContractStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCMPID, mSM_Prefix, mCommodityID, mExchangeTypeID, mAgentName, mLocationName, mLocationID As Variant
Dim NoRecord As Boolean
Dim OpenBalWc, wc As Variant

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   mLocationID = Null
   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!LocationName
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Call ClearFrame
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ReservationContractStock.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ReservationContractStock.xls")
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

Private Sub CmdSearch_Click()

Dim wc As String
Call Grid_Head
LblTotalRecord.Caption = ""
MsgBox ("Wait till Next Message")

If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox "To date must be always greater than from date!!!"
   TxtToDate.SetFocus
   Exit Sub
End If


If CmbSCommodity.Text <> "" Then
   wc = wc & " And MC.Commodity='" & CmbSCommodity.Text & "'"
End If
If CmbSLocation.Text <> "" Then
   wc = wc & " And ML.LocationName='" & CmbSLocation.Text & "'"
End If

If TxtClientName.Text <> "" Then
   wc = wc & " And MCl.ClientName Like '" & TxtClientName.Text & "%'"
End If

tmp = ""
tmp = " Select MSC.ContractID,MSC.ContractDate,MSC.ContractName,MCl.ClientName,ML.LocationName,MC.Commodity,MSC.BookedFrom,MSC.BookedTo,MSC.BookedAsPer,"
tmp = tmp & " MSC.BookedRate , MSC.StorageCharges, MSC.BookedQty, RCS.BalanceBags, RCS.BalanceQty"
tmp = tmp & " From Mst_StorageContract MSC"
tmp = tmp & " left join Vw_ReservationContractStock RCS on MSC.ContractID=RCS.ContractID"
tmp = tmp & " inner join Mst_Location ML on MSC.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_Client MCl on MSC.ClientIDRow=MCl.ClientIDRow"
tmp = tmp & " left join Mst_Commodity MC on MSC.CommodityID=MC.CommodityID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where (MSC.ContractDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
Else
   tmp = tmp & " Where (MSC.ContractDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
End If

If wc <> "" Then
   tmp = tmp & wc
End If

If Option1.Value = True Then
   tmp = tmp & " And (Isnull(RCS.BalanceBags,0) <=0 or Isnull(RCS.BalanceQty,0) <=0) "
End If
If Check1.Value = 0 Then
   tmp = tmp & " And MSC.BookedTo > GetDate()"
End If

tmp = tmp & " Order By MSC.ContractID"

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If MSHFlexGrid1.TextMatrix(I, 13) = "" Then
       MSHFlexGrid1.TextMatrix(I, 13) = 0
    End If
    If MSHFlexGrid1.TextMatrix(I, 12) = "" Then
       MSHFlexGrid1.TextMatrix(I, 12) = 0
    End If
    MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(I, 6) = Format(MSHFlexGrid1.TextMatrix(I, 6), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(I, 7) = Format(MSHFlexGrid1.TextMatrix(I, 7), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(I, 9) = Format(MSHFlexGrid1.TextMatrix(I, 9), "#####0.00")
    MSHFlexGrid1.TextMatrix(I, 10) = Format(MSHFlexGrid1.TextMatrix(I, 10), "#####0.00")
    MSHFlexGrid1.TextMatrix(I, 13) = Format(MSHFlexGrid1.TextMatrix(I, 13), "#####0.000")
    TmpRs.MoveNext
Next

LblTotalRecord.Caption = TmpRs.RecordCount
MsgBox ("Done!!")

End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call ClearFrame
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call ClearFrame
Call Grid_Head
End Sub
Private Sub ClearFrame()
From_GenDate = Date
TxtFromDate.Value = CDate(From_GenDate) - 30 'Date
TxtToDate.Value = Date
CmbSLocation.Text = ""
CmbSCommodity.Text = ""
TxtClientName = ""
Option2.Value = True
LblTotalRecord.Caption = ""
Check1.Value = 0
End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 14
     .WordWrap = True
     .AllowUserResizing = flexResizeColumns
     
     .TextMatrix(0, 0) = "Reservation ID"
     .TextMatrix(0, 1) = "Reservation Date"
     .TextMatrix(0, 2) = "Reservation Description"
     .TextMatrix(0, 3) = "Client Name"
     .TextMatrix(0, 4) = "Location"
     .TextMatrix(0, 5) = "Commodity"
     .TextMatrix(0, 6) = "Booked From"
     .TextMatrix(0, 7) = "Booked To "
     .TextMatrix(0, 8) = "Booked As Per"
     .TextMatrix(0, 9) = "Booked Rate"
     .TextMatrix(0, 10) = "Storage Charges"
     .TextMatrix(0, 11) = "Booked Qty"
     .TextMatrix(0, 12) = "Balance Bags"
     .TextMatrix(0, 13) = "Balance Qty"
          
     .ColWidth(0) = 1000
     .ColWidth(1) = 1200
     .ColWidth(2) = 2200
     .ColWidth(3) = 2200
     .ColWidth(4) = 2200 '800
     .ColWidth(5) = 2200 '
     .ColWidth(6) = 1200
     .ColWidth(7) = 1200
     .ColWidth(8) = 1400
     .ColWidth(9) = 1400
     .ColWidth(10) = 1400
     .ColWidth(11) = 1200
     .ColWidth(12) = 1400
     .ColWidth(13) = 1400
     .RowHeight(0) = 600
End With
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
mCommodityID = RsMst_Commodity!CommodityID
End Sub

