VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_ReservationInvoiceDetails 
   Caption         =   "Reservation Invoice Details "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   8280
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame0 
      Height          =   8190
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14850
      Begin VB.ComboBox CmbSearchBy 
         Height          =   315
         ItemData        =   "FrmRpt_ReservationInvoiceDetails.frx":0000
         Left            =   4005
         List            =   "FrmRpt_ReservationInvoiceDetails.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   3120
      End
      Begin VB.TextBox TxtSearchBy 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7200
         TabIndex        =   21
         Top             =   960
         Width           =   2340
      End
      Begin VB.ComboBox CmbStateID 
         Height          =   315
         Left            =   4500
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   3000
      End
      Begin VB.ComboBox CmbRegion 
         Height          =   315
         Left            =   2805
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   1665
      End
      Begin VB.ComboBox CmbNonNcdexClient 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   3870
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   7545
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   3045
      End
      Begin VB.ComboBox CmbSContract 
         Height          =   315
         Left            =   10605
         Sorted          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   4110
      End
      Begin VB.ComboBox CmbSStatus 
         Height          =   315
         ItemData        =   "FrmRpt_ReservationInvoiceDetails.frx":0033
         Left            =   11310
         List            =   "FrmRpt_ReservationInvoiceDetails.frx":003D
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   1875
      End
      Begin VB.CheckBox ChkBatchNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   13920
         TabIndex        =   4
         Top             =   990
         Width           =   225
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
         Left            =   13515
         TabIndex        =   3
         Top             =   1470
         Width           =   1000
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
         Left            =   12450
         TabIndex        =   2
         Top             =   1470
         Width           =   1000
      End
      Begin VB.TextBox TxtSInvoiceNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   9645
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6135
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   10821
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
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   48824321
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   300
         Left            =   1455
         TabIndex        =   28
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   48824321
         CurrentDate     =   39884
      End
      Begin VB.Label Label14 
         Caption         =   "To Date"
         Height          =   225
         Left            =   1807
         TabIndex        =   30
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
         Height          =   225
         Left            =   412
         TabIndex        =   29
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Search By"
         Height          =   255
         Left            =   5175
         TabIndex        =   26
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblSearchBy 
         Caption         =   "TI No (Voucher No)"
         Height          =   255
         Left            =   7650
         TabIndex        =   25
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
         Height          =   195
         Left            =   5580
         TabIndex        =   24
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Region"
         Height          =   195
         Left            =   3360
         TabIndex        =   23
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label37 
         Caption         =   "Client"
         Height          =   225
         Left            =   1808
         TabIndex        =   17
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label36 
         Caption         =   "Location"
         Height          =   300
         Left            =   8730
         TabIndex        =   16
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label11 
         Caption         =   "Contract"
         Height          =   225
         Left            =   12323
         TabIndex        =   15
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Status"
         Height          =   225
         Left            =   12000
         TabIndex        =   14
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label30 
         Caption         =   "Approved By FAG"
         Height          =   255
         Left            =   13440
         TabIndex        =   13
         Top             =   720
         Width           =   1320
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
         Left            =   6000
         TabIndex        =   12
         Top             =   1500
         Width           =   2415
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
         Left            =   8475
         TabIndex        =   11
         Top             =   1515
         Width           =   90
      End
      Begin VB.Label Label32 
         Caption         =   "PI No"
         Height          =   210
         Left            =   10185
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmRpt_ReservationInvoiceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mRegionID, mStateID, mLocationID As Variant

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame0.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame0.Enabled = True
Call Grid_Head3
ProgressBar1.Value = 0
Call ClearData
End Sub

Private Sub ClearData()
Dim mSMonth, mSYear, mSTotalDays As Variant
TxtFromDate.Value = Date
TxtToDate.Value = Date
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
Call TxtFromDate_LostFocus
Call TxtToDate_LostFocus
Label22.Caption = ""
CmbRegion.Text = ""
CmbStateID.Text = ""
CmbSLocation.Text = ""
CmbNonNcdexClient.Text = ""
CmbSContract.Text = ""
CmbSearchBy.Text = ""
TxtSearchBy = ""
TxtSInvoiceNo = ""
ChkBatchNo.Value = 1
CmbSStatus.Text = ""


End Sub

Private Sub CmbRegion_GotFocus()

tmp = CmbRegion.Text

If Gen_WcState <> "" Then
   Call TableMst_Region(" Where REgionID in (Select distinct RegionID from Mst_State Where " & Gen_WcState & " ")
Else
   Call TableMst_Region
End If
CmbRegion.Clear

If RsMst_Region.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Region.RecordCount
    CmbRegion.AddItem RsMst_Region!Region
    RsMst_Region.MoveNext
Next
CmbRegion.Text = tmp


End Sub

Private Sub CmbRegion_LostFocus()

If CmbRegion.Text = "" Then
   CmbStateID.Text = ""
   CmbSLocation.Text = ""
   Exit Sub
End If
RsMst_Region.MoveFirst
RsMst_Region.Find "Region = '" & CmbRegion.Text & "'"
If RsMst_Region.EOF Then
   MsgBox "Invalid selection "
   CmbRegion.SetFocus
   Exit Sub
End If

If RsMst_Region!Region <> CmbRegion.Text Then
   CmbStateID.Text = ""
   CmbSLocation.Text = ""
   Exit Sub
End If

mRegionID = RsMst_Region!RegionID

End Sub


Private Sub CmbStateID_GotFocus()
If CmbRegion.Text = "" Then
   MsgBox "Select Region!!!"
   CmbRegion.SetFocus
   Exit Sub
End If

tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where RegionID = " & mRegionID & " ANd " & Gen_WcState)
Else
   Call TableMst_State("Where RegionID = " & mRegionID & "")
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
   CmbSLocation.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If tmp <> CmbStateID.Text Then
   CmbSLocation.Text = ""
   Exit Sub
End If

mStateID = RsMst_State!StateID

End Sub


Private Sub CmbSearchBy_LostFocus()

If CmbSearchBy.Text = "" Then
   lblSearchBy.Caption = "TI No (Voucher No)"
   TxtSearchBy = ""
   TxtSearchBy.Locked = True
   Exit Sub
End If

If LCase(CmbSearchBy.Text) <> "ti no" And LCase(CmbSearchBy.Text) <> "post invoice no" And LCase(CmbSearchBy.Text) <> "entry no" Then
   MsgBox "Invalid Selection!!!"
   CmbSearchBy.SetFocus
   Exit Sub
End If

lblSearchBy.Caption = CmbSearchBy.Text
TxtSearchBy.Locked = False

End Sub


Private Sub CmbSLocation_GotFocus()

If CmbStateID.Text = "" Then
   MsgBox "Select State!!!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbSLocation.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where StateID = " & mStateID & " And  " & Gen_WcLocation)  '(" Where LocationId > 0 ")
Else
   Call TableMst_Location(" Where StateID = " & mStateID & "")
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!Locationname
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
mSLocationID = RsMst_Location!LocationID
End Sub
Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) = "approved" Then
   mSStatus = "A"
ElseIf LCase(CmbSStatus.Text) = "pending" Then
   mSStatus = "P"
Else
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmbSContract_GotFocus()
tmp = CmbSContract.Text
Call TableMst_StorageContract
CmbSContract.Clear
If RsMst_StorageContract.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_StorageContract.RecordCount
    CmbSContract.AddItem RsMst_StorageContract!ContractName
    RsMst_StorageContract.MoveNext
Next
CmbSContract.Text = tmp
End Sub
Private Sub CmbSContract_LostFocus()
If CmbSContract.Text = "" Then
   Exit Sub
End If
RsMst_StorageContract.MoveFirst
RsMst_StorageContract.Find "ContractName = '" & CmbSContract.Text & "'"
If RsMst_StorageContract.EOF Then
   MsgBox "Invalid selection "
   CmbSContract.SetFocus
   Exit Sub
End If
mContractIDRow = RsMst_StorageContract!ContractID

End Sub


Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub
Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID

End Sub

Private Sub CmdGo_Click()

Call Grid_Head3
ProgressBar1.Value = 0

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date Should Be Less Than To Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

MsgBox "Wait till next message!!!"

tmp = " Select ML.LocationName,TRI.Flag,TRI.InvoiceNo,TRI.TINO,TRI.TIAmount,TRI.TIDate,TRI.ActualAmount,MSC.ContractName,"
tmp = tmp & " TRI.BillingAmount ,TRI.BillingDate,TRI.BatchID , TRI.YearMonth, TRI.Remarks,TRI.FLag 'InvoiceFlag',"
tmp = tmp & " MC.Commodity , MCL.ClientNAme, " '",TRI.ApprovalRemarks ,TRI.ApprovalDate"
tmp = tmp & " TTD.TIID,TTD.TINo,TTD.TIDate,TTD.TIAmount,TTD.TISTAmount,TTD.EntryNo,TRI.ReverseFlag,ReverseRemark,ReverseDate,RBatchID,TRI.ServiceTaxNAVCode,isnull(TRID.FromDate,'')[FromDate],isnull(TRID.ToDate,'')[ToDate],"
tmp = tmp & " isnull(MBC.BillingCycle,'')[BillingCycle],isnull(TRID.Flag,'')[Flag],isnull(TRID.PeakValue,0)[PeakValue],"
tmp = tmp & " isnull(TRID.Rate,0)[Rate],isnull(TRID.Amount,0)[Amount],MS.STateName,MR.Region"
tmp = tmp & " from Txn_ReservationInvoice TRI"
tmp = tmp & " inner join Mst_StorageContract MSC on MSC.ContractID=TRI.ContractID"
tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
tmp = tmp & " inner join Mst_Client MCL on MCL.ClientID=MSC.ClientID"
tmp = tmp & " inner join mst_Location ML on ML.LocationID=TRI.LocationID"
tmp = tmp & " Left Join Txn_TIDetails TTD On TRI.InvoiceNo =  TTD.PIID And PIFlag = 'Reservation' "
tmp = tmp & " Left Join Txn_ReservationInvoiceDetail TRID On TRI.InvoiceNo = TRID.InvoiceNo"
tmp = tmp & " inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
tmp = tmp & " Left Join Mst_BillingCycle MBC On TRID.BillingCycleID = MBC.BillingCycleID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TRI.BillingDate >= #" & TxtFromDate.Value & "# And TRI.BillingDate <= #" & TxtToDate.Value & "# "  'tmp = tmp & " 'Where BillingDate = #" & mToDate & "# " 'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
Else
   tmp = tmp & " Where TRI.BillingDate >= '" & TxtFromDate.Value & "' And TRI.BillingDate <= '" & TxtToDate.Value & "' " 'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
End If

If CmbSContract.Text <> "" Then
   tmp = tmp & " And ContractName = '" & CmbSContract.Text & "'"
End If

If CmbRegion.Text <> "" Then
   tmp = tmp & " And MR.Region = '" & CmbRegion.Text & "'"
End If

If CmbStateID.Text <> "" Then
   tmp = tmp & " And MS.StateName = '" & CmbStateID.Text & "'"
Else
   If Gen_WcState <> "" Then
      tmp = tmp & " And MS." & Gen_WcState
   End If
End If

If CmbSLocation.Text <> "" Then
   tmp = tmp & " And LocationName = '" & CmbSLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

If CmbNonNcdexClient.Text <> "" Then
   tmp = tmp & " And MCL.ClientNAme = '" & CmbNonNcdexClient.Text & "'"
End If

'If Trim(CmbSForwardTo.Text) <> "" Then
'
'   strSplitString = Split(CmbSForwardTo, "~")
'   EmpName_ = Trim(strSplitString(0))
'   EmpNo_ = Trim(strSplitString(1))
'   tmp = tmp & " And ME.EmployeeNo = '" & EmpNo_ & "'"
'End If

mFlag = Left(CmbSStatus.Text, 1)
If CmbSStatus.Text <> "" Then
   tmp = tmp & " And TRI.Flag = '" & mFlag & "'"
End If

If ChkBatchNo.Value = 0 Then
   'tmp = tmp & " AND CS.BatchID = 0 "
   tmp = tmp & " And BatchID = 0 "
End If

If TxtSInvoiceNo <> "" Then
   tmp = tmp & " And TRI.InvoiceNo = " & TxtSInvoiceNo & " "
End If

If CmbSearchBy.Text <> "" Then
   If Left(LCase(CmbSearchBy.Text), 1) = "t" Then
      tmp = tmp & " And TTD.TINo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "p" Then
      tmp = tmp & " And TTD.PostInvoiceNo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "e" Then
      tmp = tmp & " And TTD.EntryNo like '%" & TxtSearchBy & "%'"
   End If
End If

tmp = tmp & " Order By TRI.InvoiceNo"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   ProgressBar1.Max = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs1!InvoiceNo
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs1!ContractName
       'mLocationId = TmpRs1!LocationID
       
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs1!Locationname
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
       MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs1!YearMonth, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 5) = Format(TmpRs1!BillingDate, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!BillingAmount), 0, TmpRs1!BillingAmount)
       MSHFlexGrid1.TextMatrix(I, 7) = TmpRs1!ClientName
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!ActualAmount), "", TmpRs1!ActualAmount) '
       MSHFlexGrid1.TextMatrix(I, 9) = GetServiceTaxDescFromNAVCode(TmpRs1!ServiceTaxNAVCode) 'IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo) 'TmpRs1!TINO
       MSHFlexGrid1.TextMatrix(I, 10) = "" 'IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-mm-yyyy")) '
'       MSHFlexGrid1.TextMatrix(I, 10) = Format(TmpRs1!TIDate, "dd-mm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 11) = "" 'IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
       MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!BatchID), 0, TmpRs1!BatchID) '
       MSHFlexGrid1.TextMatrix(I, 13) = TmpRs1!InvoiceFlag
       If LCase(MSHFlexGrid1.TextMatrix(I, 13)) = "p" Then
          MSHFlexGrid1.TextMatrix(I, 13) = "Pending"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 13)) = "a" Then
          MSHFlexGrid1.TextMatrix(I, 13) = "Approved"
       End If
       
       MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
       
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs1!TIID), "", TmpRs1!TIID)
       MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo)
       MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
       MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs1!TISTAmount), "", TmpRs1!TISTAmount)
       MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs1!EntryNo), "", TmpRs1!EntryNo)
       MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs1!ReverseFlag), "", TmpRs1!ReverseFlag)
       MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs1!ReverseRemark), "", TmpRs1!ReverseRemark)
       MSHFlexGrid1.TextMatrix(I, 23) = IIf(IsNull(TmpRs1!ReverseDate), "", Format(TmpRs1!ReverseDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs1!RBatchID), "", TmpRs1!RBatchID)
       
       MSHFlexGrid1.TextMatrix(I, 25) = IIf(IsNull(TmpRs1!FromDate), "", Format(TmpRs1!FromDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 26) = IIf(IsNull(TmpRs1!ToDate), "", Format(TmpRs1!ToDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 27) = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
       MSHFlexGrid1.TextMatrix(I, 28) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
       MSHFlexGrid1.TextMatrix(I, 29) = IIf(IsNull(TmpRs1!PeakValue), "", TmpRs1!PeakValue)
       MSHFlexGrid1.TextMatrix(I, 30) = IIf(IsNull(TmpRs1!Rate), "", TmpRs1!Rate)
       MSHFlexGrid1.TextMatrix(I, 31) = IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount)
       MSHFlexGrid1.TextMatrix(I, 32) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
       MSHFlexGrid1.TextMatrix(I, 33) = IIf(IsNull(TmpRs1!Region), "", TmpRs1!Region)

       If MSHFlexGrid1.TextMatrix(I, 28) = "Q" Then
          MSHFlexGrid1.TextMatrix(I, 28) = "Wt in MT"
       ElseIf MSHFlexGrid1.TextMatrix(I, 28) = "B" Then
          MSHFlexGrid1.TextMatrix(I, 28) = "No of Bags"
       ElseIf MSHFlexGrid1.TextMatrix(I, 28) = "F" Then
          MSHFlexGrid1.TextMatrix(I, 28) = "Fixed"
       End If


       If LCase(MSHFlexGrid1.TextMatrix(I, 21)) = "y" Then
          MSHFlexGrid1.TextMatrix(I, 21) = "Yes"
       Else
          MSHFlexGrid1.TextMatrix(I, 21) = "No"
       End If
       
       TmpRs1.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Label22.Caption = ""
Label22.Caption = TmpRs1.RecordCount
Label22.Refresh

MsgBox "Done!!!"

End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ReservationInvoice.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ReservationInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
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
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub Grid_Head3()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.Cols = 34 ' '25
MSHFlexGrid1.Font.Size = 10

MSHFlexGrid1.TextMatrix(0, 0) = "Invoice No"
MSHFlexGrid1.TextMatrix(0, 1) = "Contract Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Year Month"
MSHFlexGrid1.TextMatrix(0, 5) = "Billing Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Billing Amount"
MSHFlexGrid1.TextMatrix(0, 7) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Actual Amount"
MSHFlexGrid1.TextMatrix(0, 9) = "Service Tax" '"TI No"
MSHFlexGrid1.TextMatrix(0, 10) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 11) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 12) = "BatchID"
MSHFlexGrid1.TextMatrix(0, 13) = "Status"
MSHFlexGrid1.TextMatrix(0, 14) = "Remarks"

MSHFlexGrid1.TextMatrix(0, 15) = "TI ID"
MSHFlexGrid1.TextMatrix(0, 16) = "TI No"
MSHFlexGrid1.TextMatrix(0, 17) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 18) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 19) = "TI Service Tax Amount"
MSHFlexGrid1.TextMatrix(0, 20) = "TI Entry No"
MSHFlexGrid1.TextMatrix(0, 21) = "Reverse Flag"

MSHFlexGrid1.TextMatrix(0, 22) = "Reverse Remark"
MSHFlexGrid1.TextMatrix(0, 23) = "Reverse Date"
MSHFlexGrid1.TextMatrix(0, 24) = "RBatch ID"


MSHFlexGrid1.TextMatrix(0, 25) = "From Date"
MSHFlexGrid1.TextMatrix(0, 26) = "To Date"
MSHFlexGrid1.TextMatrix(0, 27) = "Billing Cycle"
MSHFlexGrid1.TextMatrix(0, 28) = "Flag"
MSHFlexGrid1.TextMatrix(0, 29) = "Peak Value"
MSHFlexGrid1.TextMatrix(0, 30) = "Rate"
MSHFlexGrid1.TextMatrix(0, 31) = "Amount"
 
MSHFlexGrid1.TextMatrix(0, 32) = "NCML Area"
MSHFlexGrid1.TextMatrix(0, 33) = "Region"
 



MSHFlexGrid1.ColWidth(0) = 720 '500
MSHFlexGrid1.ColWidth(1) = 2900
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1150
MSHFlexGrid1.ColWidth(4) = 1100
MSHFlexGrid1.ColWidth(5) = 1100

MSHFlexGrid1.ColWidth(6) = 900 '500
MSHFlexGrid1.ColWidth(7) = 3200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 0
MSHFlexGrid1.ColWidth(11) = 0
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 2500
MSHFlexGrid1.ColWidth(15) = 1600
MSHFlexGrid1.ColWidth(16) = 1600
MSHFlexGrid1.ColWidth(17) = 1600
MSHFlexGrid1.ColWidth(18) = 1600
MSHFlexGrid1.ColWidth(19) = 1600
MSHFlexGrid1.ColWidth(20) = 1600
'MSHFlexGrid1.ColWidth(15) = 2500
'MSHFlexGrid1.ColWidth(16) = 1100
'MSHFlexGrid1.ColWidth(17) = 2200

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True
End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
End Sub

Private Sub TxtToDate_LostFocus()
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

