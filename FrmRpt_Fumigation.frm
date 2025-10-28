VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_Fumigation 
   Caption         =   "Fumigation"
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
      Height          =   1590
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   15000
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
         ItemData        =   "FrmRpt_Fumigation.frx":0000
         Left            =   13560
         List            =   "FrmRpt_Fumigation.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   420
         Width           =   1215
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
         TabIndex        =   8
         Top             =   1125
         Width           =   3480
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
         Left            =   10875
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   1365
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
         TabIndex        =   9
         Top             =   1125
         Width           =   2685
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
         ItemData        =   "FrmRpt_Fumigation.frx":0004
         Left            =   12291
         List            =   "FrmRpt_Fumigation.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   420
         Width           =   1215
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
         Left            =   8289
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   2535
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
         Left            =   5598
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   2640
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
         Left            =   2952
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   2595
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
         TabIndex        =   12
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
         TabIndex        =   10
         Top             =   1110
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
         Height          =   375
         Left            =   12615
         TabIndex        =   11
         Top             =   1110
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
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
         Format          =   3342337
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1551
         TabIndex        =   1
         Top             =   420
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   635
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
         Format          =   3342337
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label1 
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
         Left            =   278
         TabIndex        =   27
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label3 
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
         Left            =   1791
         TabIndex        =   26
         Top             =   120
         Width           =   870
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
         TabIndex        =   25
         Top             =   120
         Width           =   615
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
         Left            =   1200
         TabIndex        =   24
         Top             =   840
         Width           =   1320
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
         Left            =   11137
         TabIndex        =   23
         Top             =   120
         Width           =   840
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
         TabIndex        =   22
         Top             =   120
         Width           =   840
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
         TabIndex        =   21
         Top             =   840
         Width           =   1440
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
         Left            =   9039
         TabIndex        =   20
         Top             =   120
         Width           =   1035
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
         Left            =   6540
         TabIndex        =   19
         Top             =   120
         Width           =   900
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
         Left            =   3720
         TabIndex        =   18
         Top             =   120
         Width           =   1230
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
         TabIndex        =   17
         Top             =   1170
         Width           =   210
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
         TabIndex        =   16
         Top             =   1170
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7935
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13361
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_Fumigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mSM_Prefix, mCMPID, mGodownID, mCommodityID, mExchangeTypeID As Variant


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

mCommodityID = RsMst_Commodity!CommodityID

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
     .Cols = 17
     .Font.Size = 10
     .AllowUserResizing = flexResizeBoth
     
     .MergeCells = flexMergeRestrictRows
     
     .TextMatrix(2, 0) = "State"
     .TextMatrix(2, 1) = "Location"
     .TextMatrix(2, 2) = "CMP"
     .TextMatrix(2, 3) = "Godown No"
     .TextMatrix(2, 4) = "Stack No"
     .TextMatrix(2, 5) = "Lot No"
     .TextMatrix(2, 6) = "Commodity"
     .TextMatrix(2, 7) = "Exchange Type"
     .TextMatrix(2, 8) = "Client  Name"
     
     .TextMatrix(1, 9) = "Opening Balance"
     .TextMatrix(1, 10) = "Opening Balance"
     
     .TextMatrix(2, 9) = "In Bags"
     .TextMatrix(2, 10) = "In MT"
     
     .TextMatrix(1, 11) = "Deposit Details"
     .TextMatrix(1, 12) = "Deposit Details"
     
     .TextMatrix(2, 11) = "In Bags"
     .TextMatrix(2, 12) = "In MT"
     
     .TextMatrix(1, 13) = "Withdrawal Details"
     .TextMatrix(1, 14) = "Withdrawal Details"
     
     .TextMatrix(2, 13) = "In Bags"
     .TextMatrix(2, 14) = "In MT"

     .TextMatrix(1, 15) = "Closing Balance"
     .TextMatrix(1, 16) = "Closing Balance"
     
     .TextMatrix(2, 15) = "In Bags"
     .TextMatrix(2, 16) = "In MT"
     .MergeRow(1) = True
     
'     .MergeCol(9) = True
'     .MergeCol(10) = True
'
'     .MergeCol(11) = True
'     .MergeCol(12) = True
'
'     .MergeCol(13) = True
'     .MergeCol(14) = True
'
'     .MergeCol(15) = True
'     .MergeCol(16) = True
     
     .ColWidth(0) = 1800
     .ColWidth(1) = 1800
     .ColWidth(2) = 1600
     .ColWidth(3) = 900
     .ColWidth(4) = 900
     .ColWidth(5) = 900
     .ColWidth(6) = 1800
     .ColWidth(7) = 1800
     .ColWidth(8) = 2000
     .ColWidth(9) = 1200
     .ColWidth(10) = 1000
     .ColWidth(11) = 1000
     .ColWidth(12) = 1000
     .ColWidth(13) = 1000
     .ColWidth(14) = 1000
     .ColWidth(15) = 1000
     .ColWidth(16) = 1000
     .WordWrap = True
     .RowHeight(2) = 500
End With
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbSGodown.Text = ""
CmbStack.Text = ""
CmbLot.Text = ""
CmbCommodityID.Text = ""
CmbSExchangeType.Text = ""
lblRecordCount.Caption = "-"
End Sub

Private Sub CmdGetReport_Click()

'On Error GoTo msgError

Dim wc As String
wc = ""
Dim WcMS, WcML, WcMCMP, WcMG, WcST, WcLT, WcMCOM, WcEX As Variant
WcMS = "": WcML = "": WcMCMP = "": WcMG = "": WcST = "": WcLT = "": WcMCOM = "": WcEX = ""

Gen_mTimeStamp = GetTimeStamp
Call Grid_Head

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date should be less than To Date !!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If CmbStateID.Text = "" Then
   MsgBox "Please Select State !!!"
   CmbStateID.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbStateID.Text <> "" Then
   WcMS = " And MS.StateID=" & mStateID & ""
Else
   If Gen_WcState <> "" Then
      WcMS = " And MS.StateID In " & Gen_WcState
   End If
End If

If CmbLocationID.Text <> "" Then
   WcML = " And ML.LocationID=" & mLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      WcML = " And ML.LocationID In " & Gen_WcLocation
   End If
End If

If CmbCMPName.Text <> "" Then
   WcMCMP = " And MCMP.CMPID=" & mCMPID & ""
Else
   If Gen_WcCMP <> "" Then
      WcMCMP = " And MCMP.CMPID In " & Gen_WcCMP
   End If
End If

If CmbSGodown.Text <> "" Then
   WcMG = " And MG.GodownID=" & mGodownID & ""
End If

If CmbStack.Text <> "" Then
   WcST = " And TCG.StackNo='" & CmbStack.Text & "'"
End If

If CmbLot.Text <> "" Then
   WcLT = " And TCG.LotNo='" & CmbLot.Text & "'"
End If

If CmbCommodityID.Text <> "" Then
   WcMCOM = " And MC.CommodityID=" & mCommodityID & ""
End If

If CmbSExchangeType.Text <> "" Then
   WcEX = " And ME.ExchangeTypeID=" & mExchangeTypeID & ""
End If


Call TableTxn_CreateTempDayWise

If CmbLocationID.Text <> "" Then
   Call TableTxn_CreateTempGSLForDayWise(TxtFromDate.Value, mSM_Prefix)
Else
   Call TableTxn_CreateTempGSLForDayWise(TxtFromDate.Value, "")
End If

'-- Find Distinct location for those location which are having Withdral txn.

tmp = " Select Distinct TCG.SM_Prefix"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
If CmbLocationID.Text <> "" Then
   tmp = tmp & " Where TCG.SM_Prefix = '" & mSM_Prefix & "' And TCG.TxnType = 'W' And  TCD.CxTFDate Between '" & TxtFromDate.Value & "' And  '" & Format(Date, Gen_DatFormat) & "'"
Else
   tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & TxtFromDate & "' " 'And  '" & Format(Date, Gen_DatFormat) & "'"
End If
tmp = tmp & " Order By TCG.SM_Prefix"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       
       tmp = "Select TCG.Mst_GSL_ID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
       tmp = tmp & " From Txn_CxTF_GSL  TCG"
       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & TxtFromDate.Value & "' "
       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "' " 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
       tmp = tmp & " Group By TCG.SM_Prefix,TCG.Mst_GSL_ID "
       tmp = tmp & " Order By TCG.SM_Prefix,TCG.Mst_GSL_ID "
       
       Call Tmp1Table
       
       '--Update temp table with Withdrawal transaction record to get opening stock
       Call TableTxn_TempDailyGSL(" Where SM_Prefix = '" & TmpRs!SM_Prefix & "'")
       
       If RsTxn_TempDailyGSL.RecordCount > 0 Then
          For j = 1 To TmpRs1.RecordCount
              If RsTxn_TempDailyGSL.RecordCount > 0 Then
                 RsTxn_TempDailyGSL.MoveFirst
              End If
              RsTxn_TempDailyGSL.Find "Mst_GSL_ID=" & TmpRs1!Mst_GSL_ID
              If Not RsTxn_TempDailyGSL.EOF Then
                 RsTxn_TempDailyGSL!BalanceBags = RsTxn_TempDailyGSL!BalanceBags + TmpRs1!NoofBags
                 RsTxn_TempDailyGSL!BalanceWeight = RsTxn_TempDailyGSL!BalanceWeight + TmpRs1!DematWeight
                 RsTxn_TempDailyGSL!CDTFBags = RsTxn_TempDailyGSL!CDTFBags - TmpRs1!NoofBags
                 RsTxn_TempDailyGSL!CDTFWeight = RsTxn_TempDailyGSL!CDTFWeight - TmpRs1!DematWeight
                 RsTxn_TempDailyGSL!Status = "O"
                 RsTxn_TempDailyGSL.Update
              End If
           TmpRs1.MoveNext
          Next
       End If
       TmpRs.MoveNext
   Next

End If
   '-- Delete All data from temp table having Flag='E'
   
   tmp = "Delete from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where Status = 'E' "
   Call TmpTable
   
   
   '-- Insert all Temp data into GSLInvoice table for billing
   
   tmp = " Insert into Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " "
   tmp = tmp & " (Mst_GSL_ID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientID,Status,Flag1,Tflag,GrossWt,CDTFDate,CDTFNo)"
   tmp = tmp & " Select Mst_GSL_ID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientID,Status,Flag1,'O',GrossWt,CDTFDate,CDTFNo"
   tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""

   Call TmpTable
   
   
   tmp = "Drop Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
   
   Call TmpTable
   
   ' Copy GSL from transaction table for selected month for generation of billing
   
   tmp = " Insert into Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " (Mst_GSL_ID,SM_Prefix,TFlag,CxTF_DetailsID,CxTFDate,CMPID,"
   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
   tmp = tmp & " ClientID,BalanceBags,BalanceWeight,Status,Flag1,GrossWt,"
   tmp = tmp & " CDTFNo,CDTFDate)"
   tmp = tmp & " Select TCG.Mst_GSL_ID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTF_DetailsID,TCD.CxTFDate,TCD.CMPID,MGSL.GodownID,MGSL.StackNo,"
   tmp = tmp & " MGSL.LotNo,TCD.CommodityID,TCD.ExchangeTypeID,"
   tmp = tmp & " MGSL.ClientID,TCG.NoOfBags,TCG.DematWeight,MGSL.Status,MGSL.Flag1,MGSL.GrossWt,"
   tmp = tmp & " MGSL.CDTFNo,MGSL.CDTFDate"
   tmp = tmp & " From Txn_CxTF_GSL  TCG"
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix"
   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
   tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
   'tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
   tmp = tmp & " Where MGSL.SpillageFlag = 0 "
   tmp = tmp & " and TCD.CxTFDate >= '" & TxtFromDate.Value & "' And TCD.CxTFDate <= '" & TxtToDate.Value & "'"
   
   Call TmpTable
   
   
   '''Display data for opening balance
   
   tmp = " Select  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName, "
   tmp = tmp & " Sum(TCG.BalanceBags)NoOfBags,Sum(TCG.BalanceWeight)Qty , TCG.Mst_GSL_ID"
   tmp = tmp & " From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " TCG"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCG.CommodityID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCG.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_Godown MG On TCG.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID = TCG.CMPID"
   tmp = tmp & " Inner Join Mst_Client MClient On TCG.ClientID = MClient.ClientID"
   tmp = tmp & " Where TCG.TFlag='O' " 'and  TCG.CxTFDate >= '2010-04-01' And TCG.CxTFDate <= '2010-04-30'
   tmp = tmp & WcMS & WcML & WcMCMP & WcMG & WcST & WcLT & WcMCOM & WcEX
   tmp = tmp & " Group By MS.StateName,ML.LocationName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MCMP.CMPName,"
   tmp = tmp & " MC.Commodity , Me.ExchangeType,TCG.Mst_GSL_ID,MClient.ClientName"
   tmp = tmp & " Order By  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName"
   Call TmpTable
   
   ''Deposite
   
   tmp = " Select  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName, "
   tmp = tmp & " Sum(TCG.BalanceBags)NoOfBags,Sum(TCG.BalanceWeight)Qty , TCG.Mst_GSL_ID"
   tmp = tmp & " From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " TCG"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCG.CommodityID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCG.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_Godown MG On TCG.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID = TCG.CMPID"
   tmp = tmp & " Inner Join Mst_Client MClient On TCG.ClientID = MClient.ClientID"
   tmp = tmp & " Where TCG.TFlag='D' and  TCG.CxTFDate >= '" & TxtFromDate.Value & "' And TCG.CxTFDate <= '" & TxtToDate.Value & "'"
   tmp = tmp & WcMS & WcML & WcMCMP & WcMG & WcST & WcLT & WcMCOM & WcEX
   tmp = tmp & " Group By MS.StateName,ML.LocationName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MCMP.CMPName,"
   tmp = tmp & " MC.Commodity , Me.ExchangeType,TCG.Mst_GSL_ID,MClient.ClientName"
   tmp = tmp & " Order By  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName"
   Call Tmp1Table


   ''Withdrawal
   
   tmp = " Select  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName, "
   tmp = tmp & " Sum(TCG.BalanceBags)NoOfBags,Sum(TCG.BalanceWeight)Qty , TCG.Mst_GSL_ID"
   tmp = tmp & " From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " TCG"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCG.CommodityID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCG.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_Godown MG On TCG.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID = TCG.CMPID"
   tmp = tmp & " Inner Join Mst_Client MClient On TCG.ClientID = MClient.ClientID"
   tmp = tmp & " Where TCG.TFlag='W' and  TCG.CxTFDate >= '" & TxtFromDate.Value & "' And TCG.CxTFDate <= '" & TxtToDate.Value & "'"
   tmp = tmp & WcMS & WcML & WcMCMP & WcMG & WcST & WcLT & WcMCOM & WcEX
   tmp = tmp & " Group By MS.StateName,ML.LocationName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MCMP.CMPName,"
   tmp = tmp & " MC.Commodity , Me.ExchangeType, TCG.Mst_GSL_ID,MClient.ClientName"
   tmp = tmp & " Order By  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName"
   Call Tmp2Table
   x = 3
   For I = 2 To TmpRs.RecordCount + 1
       For C = 0 To TmpRs.Fields.Count - 4
           MSHFlexGrid1.TextMatrix(x, C) = TmpRs.Fields(C)
       Next
       
       'MSHFlexGrid1.TextMatrix(x, 8) = Format(MSHFlexGrid1.TextMatrix(x, 8), "dd-MMM-yyyy")
       '''Openning Balance
       MSHFlexGrid1.TextMatrix(x, 9) = TmpRs!NoofBags
       MSHFlexGrid1.TextMatrix(x, 10) = TmpRs!Qty
       
       MSHFlexGrid1.TextMatrix(x, 11) = "0"
       MSHFlexGrid1.TextMatrix(x, 12) = "0"
       MSHFlexGrid1.TextMatrix(x, 13) = "0"
       MSHFlexGrid1.TextMatrix(x, 14) = "0"
       '''Deposit
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
       End If
       TmpRs1.Find "Mst_GSL_ID=" & TmpRs!Mst_GSL_ID & " " '-and CxTFDate='" & TmpRs!CxTFDate & "'"
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(x, 11) = TmpRs1!NoofBags
          MSHFlexGrid1.TextMatrix(x, 12) = TmpRs1!Qty
       End If
       ''Withdrawal
       
       If TmpRs2.RecordCount > 0 Then
          TmpRs2.MoveFirst
       End If
       TmpRs2.Find "Mst_GSL_ID=" & TmpRs!Mst_GSL_ID & " " '-and CxTFDate='" & TmpRs!CxTFDate & "'"
       If Not TmpRs2.EOF Then
          MSHFlexGrid1.TextMatrix(x, 13) = TmpRs2!NoofBags
          MSHFlexGrid1.TextMatrix(x, 14) = TmpRs2!Qty
       End If
       
       MSHFlexGrid1.TextMatrix(x, 15) = Val(MSHFlexGrid1.TextMatrix(x, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 11)) - Val(MSHFlexGrid1.TextMatrix(x, 13))
       MSHFlexGrid1.TextMatrix(x, 16) = Val(MSHFlexGrid1.TextMatrix(x, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 12)) - Val(MSHFlexGrid1.TextMatrix(x, 14))

       MSHFlexGrid1.TextMatrix(x, 15) = Format(MSHFlexGrid1.TextMatrix(x, 15), "#######0.00")
       MSHFlexGrid1.TextMatrix(x, 16) = Format(MSHFlexGrid1.TextMatrix(x, 16), "#######0.00")

       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1

       
       MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(x, 9))
       MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(x, 10))
       MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(x, 11))
       MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(x, 12))
       MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(x, 13))
       MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(x, 14))
       MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(x, 15))
       MSHFlexGrid1.TextMatrix(0, 16) = Val(MSHFlexGrid1.TextMatrix(0, 16)) + Val(MSHFlexGrid1.TextMatrix(x, 16))
       
       x = x + 1
       TmpRs.MoveNext
       
   Next
   
   ''''For Deposit
   tmp = " Select  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName, "
   tmp = tmp & " Sum(TCG.BalanceBags)NoOfBags,Sum(TCG.BalanceWeight)Qty , TCG.Mst_GSL_ID"
   tmp = tmp & " From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " TCG"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCG.CommodityID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCG.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_Godown MG On TCG.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID = TCG.CMPID"
   tmp = tmp & " Inner Join Mst_Client MClient On TCG.ClientID = MClient.ClientID"
   tmp = tmp & " Where TCG.TFlag='D' and  TCG.CxTFDate >= '" & TxtFromDate.Value & "' And TCG.CxTFDate <= '" & TxtToDate.Value & "'"
   tmp = tmp & " And TCG.Mst_GSL_ID Not in ( Select Distinct Mst_GSL_ID From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " Where TFlag='O')"
   tmp = tmp & WcMS & WcML & WcMCMP & WcMG & WcST & WcLT & WcMCOM & WcEX
   tmp = tmp & " Group By MS.StateName,ML.LocationName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MCMP.CMPName,"
   tmp = tmp & " MC.Commodity , Me.ExchangeType, TCG.Mst_GSL_ID,MClient.ClientName"
   tmp = tmp & " Order By MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName"
   Call Tmp1Table
   
   tmp = " Select  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName, "
   tmp = tmp & " Sum(TCG.BalanceBags)NoOfBags,Sum(TCG.BalanceWeight)Qty, TCG.Mst_GSL_ID"
   tmp = tmp & " From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " TCG"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCG.CommodityID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCG.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_Godown MG On TCG.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID = TCG.CMPID"
   tmp = tmp & " Inner Join Mst_Client MClient On TCG.ClientID = MClient.ClientID"
   tmp = tmp & " Where TCG.TFlag='W' and  TCG.CxTFDate >= '" & TxtFromDate.Value & "' And TCG.CxTFDate <= '" & TxtToDate.Value & "'"
   tmp = tmp & " And TCG.Mst_GSL_ID Not in ( Select Distinct Mst_GSL_ID From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " Where TFlag='O')"
   'tmp = tmp & " And TCG.Mst_GSL_ID Not in ( Select Distinct Mst_GSL_ID From Txn_DayWiseGSLDetail Where TFlag='D' And CxTFDate >= '" & TxtFromDate.Value & "' And CxTFDate <= '" & TxtToDate.Value & "')"
   tmp = tmp & WcMS & WcML & WcMCMP & WcMG & WcST & WcLT & WcMCOM & WcEX
   tmp = tmp & " Group By MS.StateName,ML.LocationName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MCMP.CMPName,"
   tmp = tmp & " MC.Commodity , Me.ExchangeType, TCG.Mst_GSL_ID,MClient.ClientName"
   tmp = tmp & " Order By MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName"
   Call Tmp2Table
   m = MSHFlexGrid1.Rows - 1
   For I = m To TmpRs1.RecordCount + m - 1
       MSHFlexGrid1.TextMatrix(I, 9) = "0"
       MSHFlexGrid1.TextMatrix(I, 10) = "0"
       MSHFlexGrid1.TextMatrix(I, 11) = "0"
       MSHFlexGrid1.TextMatrix(I, 12) = "0"
       MSHFlexGrid1.TextMatrix(I, 13) = "0"
       MSHFlexGrid1.TextMatrix(I, 14) = "0"
       MSHFlexGrid1.TextMatrix(I, 15) = "0"
       MSHFlexGrid1.TextMatrix(I, 16) = "0"
       
       
              
       For C = 0 To TmpRs1.Fields.Count - 4
           MSHFlexGrid1.TextMatrix(I, C) = TmpRs1.Fields(C)
       Next
       MSHFlexGrid1.TextMatrix(I, 11) = TmpRs1!NoofBags
       MSHFlexGrid1.TextMatrix(I, 12) = TmpRs1!Qty
       If TmpRs2.RecordCount > 0 Then
          TmpRs2.MoveFirst
       End If
       
       TmpRs2.Find "Mst_GSL_ID=" & TmpRs1!Mst_GSL_ID & ""
       If Not TmpRs2.EOF Then
          MSHFlexGrid1.TextMatrix(I, 13) = TmpRs2!NoofBags
          MSHFlexGrid1.TextMatrix(I, 14) = TmpRs2!Qty
       End If
       MSHFlexGrid1.TextMatrix(I, 15) = Val(MSHFlexGrid1.TextMatrix(I, 11)) - Val(MSHFlexGrid1.TextMatrix(I, 13))
       MSHFlexGrid1.TextMatrix(I, 16) = Val(MSHFlexGrid1.TextMatrix(I, 12)) - Val(MSHFlexGrid1.TextMatrix(I, 14))
       MSHFlexGrid1.TextMatrix(I, 15) = Format(MSHFlexGrid1.TextMatrix(I, 15), "#######0.00")
       MSHFlexGrid1.TextMatrix(I, 16) = Format(MSHFlexGrid1.TextMatrix(I, 16), "#######0.00")
       
       'MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(i, 9))
       'MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(i, 10))
       MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 11))
       MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 12))
       MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(I, 13))
       MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(I, 14))
       MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(I, 15))
       MSHFlexGrid1.TextMatrix(0, 16) = Val(MSHFlexGrid1.TextMatrix(0, 16)) + Val(MSHFlexGrid1.TextMatrix(I, 16))

       
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
            
       
       TmpRs1.MoveNext
   
   
   
   Next
   
   ''''For Withdrwal
   tmp = " Select  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName,"
   tmp = tmp & " Sum(TCG.BalanceBags)NoOfBags,Sum(TCG.BalanceWeight)Qty , TCG.Mst_GSL_ID"
   tmp = tmp & " From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " TCG"
   tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=TCG.CommodityID"
   tmp = tmp & " Inner Join Mst_Location ML On ML.SM_PreFix=TCG.SM_Prefix"
   tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
   tmp = tmp & " Inner Join Mst_ExchangeType ME On ME.ExchangeTypeID=TCG.ExchangeTypeID"
   tmp = tmp & " Inner Join Mst_Godown MG On TCG.GodownID = MG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP On MCMP.CMPID = TCG.CMPID"
   tmp = tmp & " Inner Join Mst_Client MClient On TCG.ClientID = MClient.ClientID"
   tmp = tmp & " Where TCG.TFlag='W' and  TCG.CxTFDate >= '" & TxtFromDate.Value & "' And TCG.CxTFDate <= '" & TxtToDate.Value & "'"
   tmp = tmp & " And TCG.Mst_GSL_ID Not in ( Select Distinct Mst_GSL_ID From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " Where TFlag='O')"
   tmp = tmp & " And TCG.Mst_GSL_ID Not in ( Select Distinct Mst_GSL_ID From Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & " Where TFlag='D' And CxTFDate >= '" & TxtFromDate.Value & "' And CxTFDate <= '" & TxtToDate.Value & "')"
   tmp = tmp & WcMS & WcML & WcMCMP & WcMG & WcST & WcLT & WcMCOM & WcEX
   tmp = tmp & " Group By MS.StateName,ML.LocationName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MCMP.CMPName,"
   tmp = tmp & " MC.Commodity , Me.ExchangeType,TCG.Mst_GSL_ID,MClient.ClientName"
   tmp = tmp & " Order By  MS.StateName,ML.LocationName,MCMP.CMPName,MG.GodownNo,TCG.StackNo,TCG.LotNo,MC.Commodity,ME.ExchangeType,MClient.ClientName"
   Call Tmp2Table
   m = MSHFlexGrid1.Rows - 1
   For I = m To TmpRs2.RecordCount + m - 1
       MSHFlexGrid1.TextMatrix(I, 9) = "0"
       MSHFlexGrid1.TextMatrix(I, 10) = "0"
       MSHFlexGrid1.TextMatrix(I, 11) = "0"
       MSHFlexGrid1.TextMatrix(I, 12) = "0"
       MSHFlexGrid1.TextMatrix(I, 13) = "0"
       MSHFlexGrid1.TextMatrix(I, 14) = "0"
       For C = 0 To TmpRs1.Fields.Count - 4
           MSHFlexGrid1.TextMatrix(I, C) = TmpRs2.Fields(C)
       Next
       MSHFlexGrid1.TextMatrix(I, 13) = TmpRs2!NoofBags
       MSHFlexGrid1.TextMatrix(I, 14) = TmpRs2!Qty
       
       MSHFlexGrid1.TextMatrix(I, 15) = Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 11)) '- Val(MSHFlexGrid1.TextMatrix(x, 13))
       MSHFlexGrid1.TextMatrix(I, 16) = Val(MSHFlexGrid1.TextMatrix(I, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 12)) ' - Val(MSHFlexGrid1.TextMatrix(x, 14))
       
       MSHFlexGrid1.TextMatrix(I, 15) = Format(MSHFlexGrid1.TextMatrix(I, 15), "#######0.00")
       MSHFlexGrid1.TextMatrix(I, 16) = Format(MSHFlexGrid1.TextMatrix(I, 16), "#######0.00")
       
       'MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(i, 9))
       'MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(i, 10))
       'MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(i, 11))
       'MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(i, 12))
       
       If Val(MSHFlexGrid1.TextMatrix(I, 9)) = 0 And Val(MSHFlexGrid1.TextMatrix(I, 11)) = 0 Then
          MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) ' + Val(MSHFlexGrid1.TextMatrix(I, 13))
          MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) ' + Val(MSHFlexGrid1.TextMatrix(I, 14))
       Else
          MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(I, 13))
          MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(I, 14))
       End If
       
       MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(I, 15))
       MSHFlexGrid1.TextMatrix(0, 16) = Val(MSHFlexGrid1.TextMatrix(0, 16)) + Val(MSHFlexGrid1.TextMatrix(I, 16))
       
       
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs2.MoveNext
   Next
    
   'tmp = "Delete Txn_DayWiseGSLDetail"
   tmp = "Drop Table Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & ""
   Call TmpTable
   
'End If

lblRecordCount.Caption = MSHFlexGrid1.Rows - 4
MsgBox "Done"

Exit Sub

msgError:
MsgBox "Please add more filter to generate report!!!"

tmp = "Drop Table Txn_DayWiseGSLDetail" & Gen_UserLoginID & Gen_mTimeStamp & ""
Call TmpTable


End Sub
Private Sub Get_Deposit_Withdrawal(mGSLID As Variant, mType As Variant, r As Variant, C As Variant)

tmp = "Select Sum(TCG.NoofBags)NoOfBags,Sum(TCG.GrossWeight)Qty ,MGSL.Mst_GSL_ID,TCG.TxnType"
tmp = tmp & " From Txn_CxTF_GSL TCG"
tmp = tmp & " Inner Join Mst_GSL MGSL  On TCG.Mst_GSL_ID=MGSL.Mst_GSL_ID And TCG.SM_Prefix=MGSL.SM_PreFix"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD on TCD.SM_Prefix=TCG.SM_Prefix and TCD.CxTF_DetailsID=TCG.CxTF_DetailsID and TCD.TxnType=TCG.TxnType"
tmp = tmp & " Where   "
'TCD.CxTFDate >= '7/19/2010' And TCD.CxTFDate <= '8/3/2010'"
If Gen_DBType = "MDB" Then
   tmp = tmp & " TCD.CxTFDate >= #" & TxtFromDate.Value & "# And TCD.CxTFDate <= #" & TxtFromDate.Value & "#"
Else
   tmp = tmp & " TCD.CxTFDate >= '" & TxtFromDate.Value & "' And TCD.CxTFDate <= '" & TxtToDate.Value & "'"
End If
tmp = tmp & " And TCG.TxnType='" & mType & "' And MGSL.Mst_GSL_ID=" & mGSLID & ""
tmp = tmp & " Group By TCG.TxnType,MGSL.Mst_GSL_ID"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(r, C) = TmpRs1!NoofBags
   MSHFlexGrid1.TextMatrix(r, C + 1) = TmpRs1!Qty
End If
End Sub


Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   lblRecordCount.Caption = ""
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
lblRecordCount.Caption = ""
Call Grid_Head
TxtFromDate.Value = Date - 365
TxtToDate.Value = Date
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_Fumigation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_Fumigation.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

