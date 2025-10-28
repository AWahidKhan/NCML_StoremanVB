VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GodownwiseStockSummary 
   Caption         =   "Godown wise Stock Summary"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   9855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   14895
      Begin VB.Frame Frame1 
         Height          =   9615
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   14655
         Begin VB.ComboBox CmbGodown 
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
            Left            =   12210
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   495
            Width           =   2220
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
            Left            =   5535
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   495
            Width           =   6585
         End
         Begin VB.ComboBox CmbStateName 
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
            TabIndex        =   1
            Top             =   495
            Width           =   2790
         End
         Begin VB.ComboBox CmbLocationName 
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
            Left            =   3000
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   495
            Width           =   2460
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
            Height          =   360
            Left            =   13230
            TabIndex        =   7
            Top             =   960
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
            Height          =   360
            Left            =   12015
            TabIndex        =   6
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton CmdGo 
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
            Height          =   360
            Left            =   10800
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   960
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7545
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1500
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   13309
            _Version        =   393216
            FixedCols       =   0
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   9120
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
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
            Left            =   8310
            TabIndex        =   21
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label LblGodown 
            AutoSize        =   -1  'True
            Caption         =   "Godown No"
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
            Left            =   12788
            TabIndex        =   20
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "State Name"
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
            Left            =   983
            TabIndex        =   19
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Location Name"
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
            Left            =   3548
            TabIndex        =   18
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label22 
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
            Left            =   8775
            TabIndex        =   17
            Top             =   1020
            Width           =   90
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   6375
            TabIndex        =   16
            Top             =   1020
            Width           =   2325
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Following Guards are not linked to any godowns and hence attendance can not be marked."
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
            Left            =   120
            TabIndex        =   15
            Top             =   5760
            Width           =   9360
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   10680
            TabIndex        =   14
            Top             =   5760
            Width           =   2325
         End
         Begin VB.Label Label24 
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
            Left            =   13080
            TabIndex        =   13
            Top             =   5760
            Width           =   90
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10680
      Width           =   10935
   End
End
Attribute VB_Name = "FrmRpt_GodownwiseStockSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mGodownID As Variant

Private Sub CmbLocationName_GotFocus()
If CmbStateName.Text = "" Then
   MsgBox "Plase Select State!!!"
   CmbStateName.SetFocus
   Exit Sub
End If

tmp = CmbLocationName.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location(" where StateID=" & mStateID)
Else
   Call TableMst_Location("where StateID=" & mStateID & " and " & Gen_WcLocation)
End If

CmbLocationName.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found!!!"
   CmbLocationName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocationName.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationName.Text = tmp

End Sub

Private Sub CmbLocationName_LostFocus()
If CmbLocationName.Text = "" Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If


RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationName.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid PreFix Selection!!!"
   CmbLocationName.SetFocus
   Exit Sub
End If
If tmp <> CmbLocationName.Text Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocationName.Text = "" Then
   MsgBox "Plase Select Location!!!"
   CmbLocationName.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text

If Gen_WcCMP = "" Then
   Call TableMst_CMP(" Where LocationID = " & mLocationID)
Else
   Call TableMst_CMP(" Where LocationID = " & mLocationID & " and " & Gen_WcCMP)
End If

CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbGodown.Text = ""
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
If tmp <> CmbCMPName.Text Then
   CmbGodown.Text = ""
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text

Call TableMst_Godown("Where CMPID = " & mCMPID)

CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
   CmbGodown.AddItem RsMst_Godown!GodownNO
   RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then Exit Sub

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!!"
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbStateName_GotFocus()
tmp = CmbStateName.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateName.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
   CmbStateName.AddItem RsMst_State!StateName
   RsMst_State.MoveNext
Next
CmbStateName.Text = tmp
End Sub

Private Sub CmbStateName_LostFocus()
If CmbStateName.Text = "" Then
   CmbLocationName.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If
If tmp <> CmbStateName.Text Then
   CmbLocationName.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateName.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection!!!"
   CmbStateName.SetFocus
   Exit Sub
End If

mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbStateName.Text = ""
CmbLocationName.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownwiseStockSummary.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownwiseStockSummary.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For c = 0 To MSHFlexGrid1.Cols - 2
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant

Call Grid_Head
MsgBox "Wait till Next Message!!!"

If CmbGodown.Text <> "" Then
   wc = " and MG.GodownID = " & mGodownID
ElseIf CmbCMPName.Text <> "" Then
   wc = " and MCMP.CMPID = " & mCMPID
ElseIf Gen_WcCMP <> "" Then
   wc = " and MCMP. " & Gen_WcCMP
End If

If CmbLocationName.Text <> "" Then
   wc = " and ML.LocationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = " and ML. " & Gen_WcLocation
End If

If CmbStateName.Text <> "" Then
   wc = wc & " and MS.StateID = " & mStateID
ElseIf Gen_WcState <> "" Then
   wc = wc & " and MS. " & Gen_WcState
End If

tmp = " Select MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo, MG.GodownCapacity, MUT.UnitType, MG.GodownID"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Left Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
tmp = tmp & " Where MG.Closedate is null"
tmp = tmp & wc
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Label22.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      ProgressBar1.Max = .Rows
      For I = 2 To TmpRs.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!GodownNO), "", TmpRs!GodownNO)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GodownCapacity), "", TmpRs!GodownCapacity)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
         .TextMatrix(I, 16) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
         
         '.TextMatrix(I, 7) = IIf(IsNull(TmpRs!ReserveStock), "", Format(TmpRs!ReserveStock, "#0.000"))
         '.TextMatrix(I, 8) = IIf(IsNull(TmpRs!ExtendedPeriodStock), "", Format(TmpRs!ExtendedPeriodStock, "#0.000"))
         '.TextMatrix(I, 9) = IIf(IsNull(TmpRs!SCMStock), "", Format(TmpRs!SCMStock, "#0.000"))
         '.TextMatrix(I, 10) = IIf(IsNull(TmpRs!ManagementStock), "", Format(TmpRs!ManagementStock, "#0.000"))
         '.TextMatrix(I, 11) = IIf(IsNull(TmpRs!RegularStock), "", Format(TmpRs!RegularStock, "#0.000"))
         '.TextMatrix(I, 12) = IIf(IsNull(TmpRs!NCDEXStock), "", Format(TmpRs!NCDEXStock, "#0.000"))
         '.TextMatrix(I, 13) = IIf(IsNull(TmpRs!SpotStock), "", Format(TmpRs!SpotStock, "#0.000"))
         '.TextMatrix(I, 14) = IIf(IsNull(TmpRs!TotalStock), "", Format(TmpRs!TotalStock, "#0.000"))
         
                  
         '.TextMatrix(I, 14) = Format(Val(.TextMatrix(I, 6)) + Val(.TextMatrix(I, 7)) + Val(.TextMatrix(I, 8)) + Val(.TextMatrix(I, 9)) + Val(.TextMatrix(I, 10)) + Val(.TextMatrix(I, 11)) + Val(.TextMatrix(I, 12)) + Val(.TextMatrix(I, 13)), "#0.000")
         
         '.TextMatrix(0, 7) = Format(Val(.TextMatrix(I, 7)) + Val(.TextMatrix(0, 7)), "#0.000")
         '.TextMatrix(0, 8) = Format(Val(.TextMatrix(I, 8)) + Val(.TextMatrix(0, 8)), "#0.000")
         '.TextMatrix(0, 9) = Format(Val(.TextMatrix(I, 9)) + Val(.TextMatrix(0, 9)), "#0.000")
         '.TextMatrix(0, 10) = Format(Val(.TextMatrix(I, 10)) + Val(.TextMatrix(0, 10)), "#0.000")
         '.TextMatrix(0, 11) = Format(Val(.TextMatrix(I, 11)) + Val(.TextMatrix(0, 11)), "#0.000")
         '.TextMatrix(0, 12) = Format(Val(.TextMatrix(I, 12)) + Val(.TextMatrix(0, 12)), "#0.000")
         '.TextMatrix(0, 13) = Format(Val(.TextMatrix(I, 13)) + Val(.TextMatrix(0, 13)), "#0.000")
         '.TextMatrix(0, 14) = Format(Val(.TextMatrix(I, 14)) + Val(.TextMatrix(0, 14)), "#0.000")
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
      Call GetReservationQty
      Call GetReservedStock
      Call GetExtendedPeriodStock
      Call GetSCMStock
      Call GetManagementStock
      Call GetRegularStock
      Call GetNCDEXStock
      Call GetSpotStock
      Call GetReserveSCMStock
      For I = 2 To MSHFlexGrid1.Rows - 2
         MSHFlexGrid1.TextMatrix(I, 15) = Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(MSHFlexGrid1.TextMatrix(I, 14)), "#0.000")
         MSHFlexGrid1.TextMatrix(0, 15) = Format(Val(MSHFlexGrid1.TextMatrix(I, 15)) + Val(MSHFlexGrid1.TextMatrix(0, 15)), "#0.000")
      Next
   End With
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
   MsgBox "No Record Found!!!"
End If
End Sub

Private Sub GetReservationQty()
tmp = " select GodownID, sum(ReservationQty) ReservationQty from Txn_SRGodownMapDetails group by GodownID"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!ReservationQty), "", Format(TmpRs1!ReservationQty, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 6) = Format(Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(MSHFlexGrid1.TextMatrix(0, 6)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 6) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetReservedStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " Where mgsl.balanceweight > 0 And mgsl.NContractID > 0 "
tmp = tmp & " and mgsl.NContractID not in ( select msc.contractid from mst_storagecontract msc inner join Mst_SContractBillingDetail mscb on mscb.contractid = msc.contractid where mscb.todate = msc.extendedperiod and '" & Format(Date, "yyyy-mm-dd") & "' between mscb.fromdate and mscb.todate)"
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') AND ISNULL(mgsl.spillageflag,0) = 0  and mg.unittypeid not in (4) "
tmp = tmp & " and mgsl.SCMFlag = 0 "
tmp = tmp & " group by mgsl.godownid"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 7) = Format(Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(MSHFlexGrid1.TextMatrix(0, 7)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 7) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetExtendedPeriodStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " where mgsl.balanceweight > 0 "
tmp = tmp & " and mgsl.NContractID in ( select msc.contractid from mst_storagecontract msc inner join Mst_SContractBillingDetail mscb on mscb.contractid = msc.contractid where mscb.todate = msc.extendedperiod and '" & Format(Date, "yyyy-mm-dd") & "' between mscb.fromdate and mscb.todate)"
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') AND ISNULL(mgsl.spillageflag,0) = 0 and mg.unittypeid not in (4)"
tmp = tmp & " and mgsl.SCMFlag = 0 "
tmp = tmp & " group by mgsl.godownid "
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 8) = Format(Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(0, 8)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 8) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetSCMStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " Where mgsl.balanceweight > 0 and mgsl.SCMFlag = 1 and isnull(mgsl.NContractID,0)= 0 "
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') AND ISNULL(mgsl.spillageflag,0) = 0  and mg.unittypeid not in (4)"
tmp = tmp & " group by mgsl.godownid"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 9) = Format(Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(0, 9)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 9) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetManagementStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " Where mgsl.balanceweight > 0 And mg.unittypeid = 4 AND ISNULL(mgsl.spillageflag,0) = 0 "
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') "
tmp = tmp & " group by mgsl.godownid"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 10) = Format(Val(MSHFlexGrid1.TextMatrix(I, 10)) + Val(MSHFlexGrid1.TextMatrix(0, 10)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 10) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetRegularStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " where mgsl.balanceweight > 0 and mgsl.SCMFlag = 0 and mgsl.exchangetypeid = 1 and mgsl.SCMFlag = 0 and isnull(mgsl.NContractID,0)= 0 "
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') AND ISNULL(mgsl.spillageflag,0) = 0  and mg.unittypeid not in (4)"
tmp = tmp & " group by mgsl.godownid"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 11) = Format(Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(MSHFlexGrid1.TextMatrix(0, 11)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 11) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetNCDEXStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " where mgsl.balanceweight > 0 and mgsl.SCMFlag = 0 and mgsl.exchangetypeid = 2 and mgsl.SCMFlag = 0 and isnull(mgsl.NContractID,0)= 0 "
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') AND ISNULL(mgsl.spillageflag,0) = 0  and mg.unittypeid not in (4)"
tmp = tmp & " group by mgsl.godownid"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(MSHFlexGrid1.TextMatrix(0, 12)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 12) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetSpotStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " where mgsl.balanceweight > 0 and mgsl.SCMFlag = 0 and mgsl.exchangetypeid = 3 and mgsl.SCMFlag = 0 and isnull(mgsl.NContractID,0)= 0 "
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') AND ISNULL(mgsl.spillageflag,0) = 0  and mg.unittypeid not in (4)"
tmp = tmp & " group by mgsl.godownid"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 13) = Format(Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(MSHFlexGrid1.TextMatrix(0, 13)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 13) = "0.000"
      End If
   Next
End If
End Sub

Private Sub GetReserveSCMStock()
tmp = " select mgsl.godownid, sum(mgsl.balanceweight) balanceweight from mst_gsl mgsl"
tmp = tmp & " inner join mst_godown mg on mg.godownid = mgsl.godownid"
tmp = tmp & " Where mgsl.balanceweight > 0 and mgsl.SCMFlag = 1 and isnull(mgsl.NContractID,0) > 0 "
tmp = tmp & " and mgsl.flag not in ('E', 'Z', 'B') AND ISNULL(mgsl.spillageflag,0) = 0  and mg.unittypeid not in (4)"
tmp = tmp & " group by mgsl.godownid"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID  = " & Val(MSHFlexGrid1.TextMatrix(I, 16))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs1!BalanceWeight), "", Format(TmpRs1!BalanceWeight, "#0.000"))
         MSHFlexGrid1.TextMatrix(0, 14) = Format(Val(MSHFlexGrid1.TextMatrix(I, 14)) + Val(MSHFlexGrid1.TextMatrix(0, 14)), "#0.000")
      Else
         MSHFlexGrid1.TextMatrix(I, 14) = "0.000"
      End If
   Next
End If
End Sub


Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub SetScrollBars()
'    Frame1.Width = 12000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame1.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True
If Frame1.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub VScroll1_Change()

   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * x) / 1200
   Frame4.Top = -L

End Sub

Private Sub VScroll1_Scroll()

   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * x) / 1200
   Frame4.Top = -L

End Sub

Private Sub HScroll1_Change()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * x) / 1200
   Frame4.Left = -L
End Sub

Private Sub HScroll1_Scroll()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * x) / 1200
   Frame4.Left = -L
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call CmdClear_Click
End Sub

Private Sub Grid_Head()
Label22.Caption = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 17
   .FixedCols = 1
   .AllowUserResizing = flexResizeBoth
   .WordWrap = True
    
   .TextMatrix(1, 0) = "State"
   .TextMatrix(1, 1) = "Location"
   .TextMatrix(1, 2) = "CMP Name"
   .TextMatrix(1, 3) = "Godown No"
   .TextMatrix(1, 4) = "Godown  Capacity (MT)"
   .TextMatrix(1, 5) = "Godown Type"
   .TextMatrix(1, 6) = "Reservation Qty"
   .TextMatrix(1, 7) = "Reserve Stock"
   .TextMatrix(1, 8) = "Extended period stock"
   .TextMatrix(1, 9) = "SCM stock"
   .TextMatrix(1, 10) = "Stock Management Stock"
   .TextMatrix(1, 11) = "Other Non-NCDEX stock"
   .TextMatrix(1, 12) = "NCDEX Stock"
   .TextMatrix(1, 13) = "Spot Stock"
   .TextMatrix(1, 14) = "Reserved SCM Stock"
   .TextMatrix(1, 15) = "Total Stock"
   .TextMatrix(1, 16) = "GodownID"
   
   .ColWidth(0) = 1600
   .ColWidth(1) = 1600
   .ColWidth(2) = 2200
   .ColWidth(3) = 1200
   .ColWidth(4) = 1200
   .ColWidth(5) = 1400
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 1200
   .ColWidth(16) = 0
   
   .RowHeight(1) = 600
End With
End Sub

