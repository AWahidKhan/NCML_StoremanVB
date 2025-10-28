VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GSLDiscrepancies 
   Caption         =   "GSL Discrepancies"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8220
      Left            =   120
      TabIndex        =   22
      Top             =   1605
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   7755
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13150
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
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
         Left            =   12615
         TabIndex        =   8
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetReport 
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
         Height          =   375
         Left            =   11520
         TabIndex        =   7
         Top             =   1110
         Width           =   1095
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
         Left            =   13710
         TabIndex        =   9
         Top             =   1110
         Width           =   1095
      End
      Begin VB.ComboBox CmbStateID 
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
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   420
         Width           =   3555
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
         Left            =   3675
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   3480
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
         Left            =   7215
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   3135
      End
      Begin VB.ComboBox CmbStack 
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
         ItemData        =   "FrmRpt_GSLDiscrepancies.frx":0000
         Left            =   12291
         List            =   "FrmRpt_GSLDiscrepancies.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   1215
      End
      Begin VB.ComboBox CmbSExchangeType 
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
         Left            =   3660
         TabIndex        =   11
         Top             =   1125
         Visible         =   0   'False
         Width           =   2685
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
         Left            =   10395
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   1845
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1125
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.ComboBox CmbLot 
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
         ItemData        =   "FrmRpt_GSLDiscrepancies.frx":0004
         Left            =   13560
         List            =   "FrmRpt_GSLDiscrepancies.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   420
         Width           =   1215
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
         Left            =   7680
         TabIndex        =   21
         Top             =   1170
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
         Left            =   10080
         TabIndex        =   20
         Top             =   1170
         Width           =   210
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   1237
         TabIndex        =   19
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
         Caption         =   "Location"
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
         Left            =   4965
         TabIndex        =   18
         Top             =   120
         Width           =   900
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
         Left            =   8265
         TabIndex        =   17
         Top             =   120
         Width           =   1035
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
         Left            =   4282
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label21 
         Caption         =   "Stack No"
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
         Left            =   12478
         TabIndex        =   15
         Top             =   120
         Width           =   840
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
         Left            =   10897
         TabIndex        =   14
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
         Height          =   300
         Left            =   1298
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label22 
         Caption         =   "Lot No"
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
         Left            =   13860
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmRpt_GSLDiscrepancies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mSM_Prefix, mCMPID, mGodownID, mCommodityId, mExchangeTypeID As Variant

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
   CmbSGodown.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub
Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!"
   CmbCMPName.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
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
   CmbSGodown.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
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
End Sub

Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
CmbLocationID.Clear

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
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
   CmbCMPName.Text = ""
   CmbSGodown.Text = ""
   CmbStack.Text = ""
   CmbLot.Text = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub
Private Sub CmbSGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Please Select CMP  !!! "
   CmbSGodown.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   CmbStack.Text = ""
   CmbLot.Text = ""
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
Private Sub CmbStack_GotFocus()

If CmbSGodown.Text = "" Or mGodownID = "" Then
  MsgBox "Please Select Godown No!!!"
  CmbStack.SetFocus
  Exit Sub
End If
tmp = CmbStack.Text
Call TableMst_GSL(" Where CMPID = " & mCMPID & " And GodownID = " & mGodownID, "D")
CmbStack.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbStack.AddItem RsMst_GSL!StackNo
    RsMst_GSL.MoveNext
Next

CmbStack.Text = tmp
End Sub
'----
Private Sub CmbStack_LostFocus()
If CmbStack.Text = "" Then
   CmbLot.Text = ""
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "StackNo = '" & CmbStack.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbStack.SetFocus
   Exit Sub
End If

mCmbStack = CmbStack.Text
End Sub
Private Sub CmbSExchangeType_GotFocus()
tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeType.Text = tmp
End Sub

Private Sub CmbSExchangeType_LostFocus()
If CmbSExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeType.SetFocus
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
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!! "
   CmbCommodityID.SetFocus
   Exit Sub
End If

mCommodityId = RsMst_Commodity!CommodityID

End Sub
Private Sub CmbLot_GotFocus()
If CmbStack.Text = "" Then
  MsgBox "Please Select Stack No !!!"
  CmbLot.SetFocus
  Exit Sub
End If
tmp = CmbLot.Text
Call TableMst_GSL(" Where CMPID = " & mCMPID & " And GodownID = " & mGodownID & " And StackNo = '" & CmbStack.Text & "'")
CmbLot.Clear
If RsMst_GSL.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_GSL.RecordCount
    CmbLot.AddItem RsMst_GSL!LotNo
    RsMst_GSL.MoveNext
Next I
CmbLot.Text = tmp
End Sub
'----
Private Sub CmbLot_LostFocus()

If CmbLot.Text = "" Then
   Exit Sub
End If
RsMst_GSL.MoveFirst
RsMst_GSL.Find "LotNo = '" & CmbLot.Text & "'"
If RsMst_GSL.EOF Then
   MsgBox "Invalid selection "
   CmbLot.SetFocus
   Exit Sub
End If

End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 4
     .FixedRows = 3
     .Cols = 27 '23
     .Font.Size = 10
     .AllowUserResizing = flexResizeBoth
     .MergeCells = flexMergeRestrictRows
     
     .TextMatrix(2, 0) = "State"
     .TextMatrix(2, 1) = "Location"
     .TextMatrix(2, 2) = "CMP"
     .TextMatrix(2, 3) = "GSL ID"
     .TextMatrix(2, 4) = "Godown No"
     .TextMatrix(2, 5) = "Stack No"
     .TextMatrix(2, 6) = "Lot No"
     .TextMatrix(2, 7) = "GSL Flag"
     .TextMatrix(2, 8) = "Spillage Flag"
     .TextMatrix(2, 9) = "Madeup Lot Flag"
     
     .TextMatrix(1, 10) = "GSL Data"
     .TextMatrix(1, 11) = "GSL Data"
     .TextMatrix(1, 12) = "GSL Data"
     .TextMatrix(1, 13) = "GSL Data"
     .TextMatrix(1, 14) = "GSL Data"
     .TextMatrix(1, 15) = "GSL Data"

     .TextMatrix(2, 10) = "GSL Deposit bags"
     .TextMatrix(2, 11) = "GSL withdrawal bags"
     .TextMatrix(2, 12) = "GSL Balance bags"
     .TextMatrix(2, 13) = "GSL Deposit qty"
     .TextMatrix(2, 14) = "GSL withdrawal qty"
     .TextMatrix(2, 15) = "GSL Balance qty"
     
     .TextMatrix(1, 16) = "CxTF Details"
     .TextMatrix(1, 17) = "CxTF Details"
     .TextMatrix(1, 18) = "PWR Details"
     .TextMatrix(1, 19) = "PWR Details"
     .TextMatrix(1, 20) = "Balance Bags Details"
     .TextMatrix(1, 21) = "CxTF Details"
     .TextMatrix(1, 22) = "CxTF Details"
     .TextMatrix(1, 23) = "PWR Details"
     .TextMatrix(1, 24) = "PWR Details"
     .TextMatrix(1, 25) = "Balance Qty Details"

     .TextMatrix(2, 16) = "CxTF Deposit bags"
     .TextMatrix(2, 17) = "CxTF withdrawal bags"
     .TextMatrix(2, 18) = "PWR Deposit bags"
     .TextMatrix(2, 19) = "PWR withdrawal bags"
     .TextMatrix(2, 20) = "CxTF Balance bags"
     .TextMatrix(2, 21) = "CxTF Deposit qty"
     .TextMatrix(2, 22) = "CxTF withdrawal qty"
     .TextMatrix(2, 23) = "PWR Deposit qty"
     .TextMatrix(2, 24) = "PWR withdrawal qty"
     .TextMatrix(2, 25) = "CxTF Balance qty"
     
     
     
     
     .TextMatrix(2, 26) = "KEY FOR LOT"

     .MergeRow(1) = True
     
     
     .ColWidth(0) = 1800
     .ColWidth(1) = 1800
     .ColWidth(2) = 1800
     .ColWidth(3) = 1200
     .ColWidth(4) = 1200
     .ColWidth(5) = 1200
     .ColWidth(6) = 1200
     .ColWidth(7) = 1200
     .ColWidth(8) = 1000
     .ColWidth(9) = 1000
     
     .ColWidth(10) = 1000
     .ColWidth(11) = 1000
     .ColWidth(12) = 1000
     .ColWidth(13) = 1000
     .ColWidth(14) = 1000
     .ColWidth(15) = 1000
     .ColWidth(16) = 1000
     .ColWidth(17) = 1000
     .ColWidth(18) = 1000
     .ColWidth(19) = 1000
     .ColWidth(20) = 1000
     .ColWidth(21) = 1000
     
     .ColWidth(22) = 1000
     .ColWidth(23) = 1000
     .ColWidth(24) = 1000
     .ColWidth(25) = 1000
     .ColWidth(26) = 0
     
     
     .WordWrap = True
     .RowHeight(2) = 800
     
End With

End Sub

Private Sub CmdClear_Click()

Call Grid_Head

CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbSGodown.Text = ""
CmbStack.Text = ""
CmbLot.Text = ""
CmbCommodityID.Text = ""
CmbSExchangeType.Text = ""
LblRecordCount.Caption = "-"
ProgressBar1.Value = 0

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GSLDiscrepancies.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLDiscrepancies.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Max = MSHFlexGrid1.Rows + 5
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "mm/dd/yyyy")
           Else
              If c = 4 Or c = 5 Or c = 6 Then
                 oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid1.TextMatrix(I, c)
              Else
                 oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
              End If
           End If
        Else
           If c = 4 Or c = 5 Or c = 6 Then
              oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid1.TextMatrix(I, c)
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
Set oWB = Nothing
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

'On Error GoTo msgError

Dim Wc As String
Wc = ""

Gen_mTimeStamp = GetTimeStamp
Call Grid_Head

MsgBox "Wait till next message!!!"
ProgressBar1.Value = 0

If CmbStateID.Text = "" Then
   MsgBox "Please Select State !!!"
   CmbStateID.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" Then
   If CmbLocationID.Text = "" Then
      MsgBox "Please Select Location !!!"
      CmbLocationID.SetFocus
      Exit Sub
   End If
End If

If CmbStateID.Text <> "" Then
   If Wc = "" Then
      Wc = " Where MS.StateID=" & mStateID & ""
   Else
      Wc = Wc & " And MS.StateID=" & mStateID & ""
   End If
Else
   If Gen_WcState <> "" Then
      If Wc = "" Then
         Wc = " Where MS.StateID In " & Gen_WcState
      Else
         Wc = Wc & " And MS.StateID In " & Gen_WcState
      End If
   End If
End If

If CmbLocationID.Text <> "" Then
   If Wc = "" Then
      Wc = " Where ML.LocationID=" & mLocationID & ""
   Else
      Wc = Wc & " And ML.LocationID=" & mLocationID & ""
   End If
Else
   If Gen_WcLocation <> "" Then
      If Wc = "" Then
         Wc = " Where ML.LocationID In " & Gen_WcLocation
      Else
         Wc = Wc & " And ML.LocationID In " & Gen_WcLocation
      End If
   End If
End If

If CmbCMPName.Text <> "" Then
   If Wc = "" Then
      Wc = " Where MCMP.CMPID=" & mCMPID & ""
   Else
      Wc = Wc & " And MCMP.CMPID=" & mCMPID & ""
   End If
Else
   If Gen_WcCMP <> "" Then
      If Wc = "" Then
         Wc = " Where MCMP.CMPID In " & Gen_WcCMP
      Else
         Wc = Wc & " And MCMP.CMPID In " & Gen_WcCMP
      End If
   End If
End If

If CmbSGodown.Text <> "" Then
   If Wc = "" Then
      Wc = " Where MG.GodownID=" & mGodownID & ""
   Else
      Wc = Wc & " And MG.GodownID=" & mGodownID & ""
   End If
End If

If CmbStack.Text <> "" Then
   If Wc = "" Then
      Wc = " Where MGSL.StackNo='" & CmbStack.Text & "'"
   Else
      Wc = Wc & " And MGSL.StackNo='" & CmbStack.Text & "'"
   End If
End If

If CmbLot.Text <> "" Then
   If Wc = "" Then
      Wc = " Where MGSL.LotNo='" & CmbLot.Text & "'"
   Else
      Wc = Wc & " And MGSL.LotNo='" & CmbLot.Text & "'"
   End If
End If

'If CmbCommodityID.Text <> "" Then
'   WcMCOM = " And MC.CommodityID=" & mCommodityID & ""
'End If
'
'If CmbSExchangeType.Text <> "" Then
'   WcEX = " And ME.ExchangeTypeID=" & mExchangeTypeID & ""
'End If


tmp = "Select MGSL.gslid,ML.locationname ,MCMP.cmpname,MG.godownno, MGSL.stackno,MGSL.lotno ,MGSL.flag,MGSL.SpillageFlag,MGSL.MadeUpFlag,"
tmp = tmp & " MGSL.DepositeBags 'GSLDepositBags' ,MGSL.CDTFBags 'GSLWithdrawalBags',MGSL.BalanceBags 'GSLBalanceBags',"
tmp = tmp & " MGSL.DepositeWeight 'GSLDepositQty',MGSL.CDTFWeight 'GSLWithdrawalQty',MGSL.BalanceWeight 'GSLBalanceQty',"
tmp = tmp & " MS.statename,Convert(Varchar,MGSL.gslid)+'-'+ MGSL.SM_Prefix 'KEYFORLOT'"
tmp = tmp & " from mst_gsl MGSL"
tmp = tmp & " inner join mst_godown MG on MGSL.godownid = MG.godownid"
tmp = tmp & " inner join mst_cmp  MCMP on MG.cmpid = MCMP.cmpid"
tmp = tmp & " inner join mst_location ML on MCMP.locationid = ML.locationid"
tmp = tmp & " inner join mst_State ms on ML.Stateid = MS.stateid"
tmp = tmp & Wc
tmp = tmp & " Order By ML.locationname,MGSL.gslid"

Call TmpTable

x = 3
ProgressBar1.Max = TmpRs.RecordCount + 3

For I = 1 To TmpRs.RecordCount
    With MSHFlexGrid1
        .TextMatrix(x, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        .TextMatrix(x, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location"
        .TextMatrix(x, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) '"CMP"
        .TextMatrix(x, 3) = IIf(IsNull(TmpRs!gslid), "", TmpRs!gslid) '"GSL ID"
        .TextMatrix(x, 4) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo) '"Godown No"
        .TextMatrix(x, 5) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo) '"Stack No"
        .TextMatrix(x, 6) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo) '"Lot No"
        .TextMatrix(x, 7) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag) '"GSL Flag"
        .TextMatrix(x, 8) = IIf(IsNull(TmpRs!SpillageFlag), "", TmpRs!SpillageFlag) '"Spillage Flag"
        .TextMatrix(x, 9) = IIf(IsNull(TmpRs!MadeupFlag), "", TmpRs!MadeupFlag) '"Madeup Lot Flag"
        .TextMatrix(x, 10) = IIf(IsNull(TmpRs!GSLDepositBags), "", TmpRs!GSLDepositBags) '"GSL Deposit bags"
        .TextMatrix(x, 11) = IIf(IsNull(TmpRs!GSLwithdrawalbags), "", TmpRs!GSLwithdrawalbags) '"GSL withdrawal bags"
        .TextMatrix(x, 12) = IIf(IsNull(TmpRs!GSLBalanceBags), "", TmpRs!GSLBalanceBags) '"GSL Balance bags"
        .TextMatrix(x, 13) = IIf(IsNull(TmpRs!GSLDepositQty), "", TmpRs!GSLDepositQty) '"GSL Deposit qty"
        .TextMatrix(x, 14) = IIf(IsNull(TmpRs!GSLWithdrawalQty), "", TmpRs!GSLWithdrawalQty) '"GSL withdrawal qty"
        .TextMatrix(x, 15) = IIf(IsNull(TmpRs!GSLBalanceqty), "", TmpRs!GSLBalanceqty) '"GSL Balance qty"
        .TextMatrix(x, 26) = IIf(IsNull(TmpRs!KEYFORLOT), "", TmpRs!KEYFORLOT) '"KEY FOR LOT"
        If .TextMatrix(x, 8) = 0 Then
           .TextMatrix(x, 8) = "No"
        Else
           .TextMatrix(x, 8) = "Yes"
        End If
        If Val(.TextMatrix(x, 9)) = 0 Then
           .TextMatrix(x, 9) = "No"
        Else
           .TextMatrix(x, 9) = "Yes"
        End If
        
        For y = 10 To 15
            .TextMatrix(0, y) = Val(.TextMatrix(0, y)) + Val(.TextMatrix(x, y))
        Next
        
        x = x + 1
        .Rows = .Rows + 1
    End With
    ProgressBar1.Value = ProgressBar1.Value + 1
    TmpRs.MoveNext
Next

LblRecordCount.Caption = TmpRs.RecordCount

ProgressBar1.Value = 0



tmp = " Select MGSL.gslid,ML.locationname ,MCMP.cmpname,MG.godownno, MGSL.stackno,MGSL.lotno ,t.txntype,sum(t.NoOfBags) 'Bags' ,sum(t.GrossWeight) 'Gross Wt.',sum(t.DematWeight) 'DematWt',"
tmp = tmp & " MS.StateName,Convert(Varchar,MGSL.gslid)+'-'+ MGSL.SM_Prefix +'-'+ t.txntype 'KEYFORLOT'"
tmp = tmp & " From Txn_CxTF_GSL t"
tmp = tmp & " inner join mst_gsl MGSL on MGSL.gslid = t.gslid and MGSL.sm_prefix = t.sm_prefix"
tmp = tmp & " inner join mst_godown MG on MGSL.godownid = MG.godownid"
tmp = tmp & " inner join mst_cmp  MCMP on MG.cmpid = MCMP.cmpid"
tmp = tmp & " inner join mst_location ML on MCMP.locationid = ML.locationid"
tmp = tmp & " inner join mst_State MS on ML.stateid = MS.stateid"
tmp = tmp & Wc
tmp = tmp & " Group by MS.StateName,MGSL.gslid,ML.locationname ,MCMP.cmpname,MG.godownno, MGSL.stackno,MGSL.lotno ,t.txntype,MGSL.SM_Prefix,MS.statename"
tmp = tmp & " Order by ML.locationname,MGSL.GSLID "

Call TmpTable

tmp = "Select Convert(Varchar,MGSL.gslid)+'-'+ MGSL.SM_Prefix 'KEYFORLOT',Sum(AdjstBags) 'NewBags',Sum(AdjstQty) 'NewQty'"
tmp = tmp & " from Txn_RRNPWRGSL TRP"
tmp = tmp & " inner Join Mst_GSL MGSL On TRP.NewGSLID = MGSL.GSLID And TRP.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " inner join mst_godown MG on MGSL.godownid = MG.godownid"
tmp = tmp & " inner join mst_cmp  MCMP on MG.cmpid = MCMP.cmpid"
tmp = tmp & " inner join mst_location ML on MCMP.locationid = ML.locationid"
tmp = tmp & " inner join mst_State ms on ML.Stateid = MS.stateid"
tmp = tmp & Wc
tmp = tmp & " Group BY MGSL.gslid,MGSL.SM_Prefix"

Call Tmp1Table

tmp = "Select Convert(Varchar,MGSL.gslid)+'-'+ MGSL.SM_Prefix 'KEYFORLOT',Sum(AdjstBags) 'OldBags',Sum(AdjstQty) 'OldQty'"
tmp = tmp & " from Txn_RRNPWRGSL TRP"
tmp = tmp & " inner Join Mst_GSL MGSL On TRP.GSLID = MGSL.GSLID And TRP.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " inner join mst_godown MG on MGSL.godownid = MG.godownid"
tmp = tmp & " inner join mst_cmp  MCMP on MG.cmpid = MCMP.cmpid"
tmp = tmp & " inner join mst_location ML on MCMP.locationid = ML.locationid"
tmp = tmp & " inner join mst_State ms on ML.Stateid = MS.stateid"
tmp = tmp & Wc
tmp = tmp & " Group BY MGSL.gslid,MGSL.SM_Prefix"

Call Tmp2Table


If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = MSHFlexGrid1.Rows + 5
   For I = 3 To MSHFlexGrid1.Rows - 1
       If MSHFlexGrid1.TextMatrix(I, 26) = "" Then Exit For
       MSHFlexGrid1.TextMatrix(I, 16) = 0
       MSHFlexGrid1.TextMatrix(I, 17) = 0
       MSHFlexGrid1.TextMatrix(I, 18) = 0
       MSHFlexGrid1.TextMatrix(I, 19) = 0
       MSHFlexGrid1.TextMatrix(I, 20) = 0
       MSHFlexGrid1.TextMatrix(I, 21) = 0
       MSHFlexGrid1.TextMatrix(I, 22) = 0
       MSHFlexGrid1.TextMatrix(I, 23) = 0
       MSHFlexGrid1.TextMatrix(I, 24) = 0
       MSHFlexGrid1.TextMatrix(I, 25) = 0
       
       If TmpRs.RecordCount > 0 Then
          TmpRs.MoveFirst
          TmpRs.Find "KEYFORLOT = " & MSHFlexGrid1.TextMatrix(I, 26) & "-D"
          If TmpRs.EOF = False Then
             MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!Bags), 0, TmpRs!Bags)
             MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs!DematWt), 0, TmpRs!DematWt)
          End If
          TmpRs.MoveFirst
          TmpRs.Find "KEYFORLOT = " & MSHFlexGrid1.TextMatrix(I, 26) & "-W"
          If TmpRs.EOF = False Then
             MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs!Bags), 0, TmpRs!Bags)
             MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs!DematWt), 0, TmpRs!DematWt)
          End If
       End If
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
          TmpRs1.Find "KEYFORLOT = " & MSHFlexGrid1.TextMatrix(I, 26)
          If TmpRs1.EOF = False Then
             MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs1!NewBags), 0, TmpRs1!NewBags)
             MSHFlexGrid1.TextMatrix(I, 23) = IIf(IsNull(TmpRs1!NewQty), 0, TmpRs1!NewQty)
          End If
       End If
       
       If TmpRs2.RecordCount > 0 Then
          TmpRs2.MoveFirst
          TmpRs2.Find "KEYFORLOT = " & MSHFlexGrid1.TextMatrix(I, 26)
          If TmpRs2.EOF = False Then
             MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs2!OldBags), 0, TmpRs2!OldBags)
             MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs2!OldQty), 0, TmpRs2!OldQty)
          End If
       End If
       
       MSHFlexGrid1.TextMatrix(I, 20) = Val(Val(MSHFlexGrid1.TextMatrix(I, 16)) + Val(MSHFlexGrid1.TextMatrix(I, 18))) - Val(Val(MSHFlexGrid1.TextMatrix(I, 17)) + Val(MSHFlexGrid1.TextMatrix(I, 19)))
       MSHFlexGrid1.TextMatrix(I, 25) = Val(Val(MSHFlexGrid1.TextMatrix(I, 21)) + Val(MSHFlexGrid1.TextMatrix(I, 23))) - Val(Val(MSHFlexGrid1.TextMatrix(I, 22)) + Val(MSHFlexGrid1.TextMatrix(I, 24)))
       MSHFlexGrid1.TextMatrix(I, 25) = Format(MSHFlexGrid1.TextMatrix(I, 25), "#######0.000")
    
       For y = 16 To 25
           MSHFlexGrid1.TextMatrix(0, y) = Val(MSHFlexGrid1.TextMatrix(0, y)) + Val(MSHFlexGrid1.TextMatrix(I, y))
       Next
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
    
    ProgressBar1.Value = ProgressBar1.Max
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
   LblRecordCount.Caption = ""
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
LblRecordCount.Caption = ""
Call Grid_Head
End Sub
