VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_GSLCDTFvsTRFDate 
   Caption         =   "GSL wise CDTF vs TRF vs KR Date"
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
      Height          =   9900
      Left            =   45
      TabIndex        =   13
      Top             =   45
      Width           =   14745
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
         Left            =   13500
         TabIndex        =   26
         Top             =   1260
         Width           =   1000
      End
      Begin VB.ComboBox CmbStateID 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2685
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   60
         TabIndex        =   24
         Top             =   9480
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox TxtSCDTFNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8580
         TabIndex        =   7
         Top             =   735
         Width           =   1605
      End
      Begin VB.TextBox TxtSClientName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   765
         Width           =   4545
      End
      Begin VB.ComboBox CmbPreFix 
         Height          =   315
         Left            =   2835
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2445
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
         Left            =   12495
         TabIndex        =   12
         Top             =   1260
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   10665
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   4005
      End
      Begin VB.ComboBox CmbSGodown 
         Height          =   315
         Left            =   9120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   5325
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3765
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
         Height          =   350
         Left            =   11490
         TabIndex        =   11
         Top             =   1260
         Width           =   1000
      End
      Begin VB.TextBox TxtSGSLID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6585
         MaxLength       =   10
         TabIndex        =   6
         Top             =   750
         Width           =   1125
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "CDTF Date"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   10320
         TabIndex        =   8
         Top             =   705
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7740
         Left            =   60
         TabIndex        =   14
         Top             =   1680
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   13653
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
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
         Height          =   315
         Left            =   11700
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
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
         Format          =   22478849
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   315
         Left            =   13350
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22478849
         CurrentDate     =   39864
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
         Left            =   120
         TabIndex        =   28
         Top             =   1320
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
         Left            =   2520
         TabIndex        =   27
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label Label1 
         Caption         =   "State Name"
         Height          =   210
         Left            =   975
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "CDTF No"
         Height          =   225
         Left            =   7770
         TabIndex        =   23
         Top             =   780
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Client Name"
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label8 
         Caption         =   "Location"
         Height          =   210
         Left            =   3727
         TabIndex        =   21
         Top             =   135
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Godown"
         Height          =   225
         Left            =   9570
         TabIndex        =   20
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "CMP"
         Height          =   225
         Left            =   7027
         TabIndex        =   19
         Top             =   135
         Width           =   360
      End
      Begin VB.Label Label16 
         Caption         =   "GSL ID"
         Height          =   180
         Left            =   5955
         TabIndex        =   18
         Top             =   825
         Width           =   585
      End
      Begin VB.Label LblFrom 
         Caption         =   "From"
         Height          =   225
         Left            =   11280
         TabIndex        =   17
         Top             =   765
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label LblTo 
         Caption         =   "To"
         Height          =   225
         Left            =   13080
         TabIndex        =   16
         Top             =   765
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label3 
         Caption         =   "Commodity"
         Height          =   180
         Left            =   12247
         TabIndex        =   15
         Top             =   135
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmRpt_GSLCDTFvsTRFDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSStateID, mSLocationID, mSCMPID, mSCommodityID, mSGodownID, msSM_prefix As Variant
Private Sub CmdClear_Click()

Call Grid_Head
CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSGodown.Text = ""
CmbSCommodity.Text = ""
TxtSClientName.Text = ""
TxtSCDTFNo.Text = ""
TxtSGSLID.Text = ""
Check2.Value = 0
Check2_Click
lblRecordCount.Caption = ""
ProgressBar1.Value = 0

End Sub


Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text


'Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID = " & mSStateID & "")
Else
   Call TableMst_Location("where StateID = " & mSStateID & " And " & Gen_WcSMLocation)
End If

CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
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
   MsgBox "Invalid PreFix Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
msSM_prefix = RsMst_Location!SM_Prefix
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSClient_Change()

End Sub

Private Sub CmbSCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "" & tmp1)
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
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
mSCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
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

mSCommodityID = RsMst_Commodity!CommodityID

End Sub


Private Sub CmbSGodown_GotFocus()
If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mSCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
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
mSGodownID = RsMst_Godown!GodownID
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    LblFrom.Visible = True
    LblTo.Visible = True
    TxtFromDate.Visible = True
    TxtToDate.Visible = True
    TxtFromDate.Value = Date - 365
Else
    LblFrom.Visible = False
    LblTo.Visible = False
    TxtFromDate.Visible = False
    TxtToDate.Visible = False
    TxtToDate.Value = Date
End If
End Sub


Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 22
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(1, 0) = "GSLID"
MSHFlexGrid1.TextMatrix(1, 1) = "State Name"
MSHFlexGrid1.TextMatrix(1, 2) = "Location"
MSHFlexGrid1.TextMatrix(1, 3) = "CMP Name"
MSHFlexGrid1.TextMatrix(1, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(1, 5) = "Godown No"
MSHFlexGrid1.TextMatrix(1, 6) = "Stack No"
MSHFlexGrid1.TextMatrix(1, 7) = "Lot No"
MSHFlexGrid1.TextMatrix(1, 8) = "Deposited Bags"
MSHFlexGrid1.TextMatrix(1, 9) = "Deposited Weight"
MSHFlexGrid1.TextMatrix(1, 10) = "Client Name"
MSHFlexGrid1.TextMatrix(1, 11) = "Agent Name"
MSHFlexGrid1.TextMatrix(1, 12) = "CDTF No"
MSHFlexGrid1.TextMatrix(1, 13) = "CDTF Date"
MSHFlexGrid1.TextMatrix(1, 14) = "CDTF Created Date"
MSHFlexGrid1.TextMatrix(1, 15) = "TRF No"
MSHFlexGrid1.TextMatrix(1, 16) = "TRF Date"
MSHFlexGrid1.TextMatrix(1, 17) = "TRF Created Date"
MSHFlexGrid1.TextMatrix(1, 18) = "KR No"
MSHFlexGrid1.TextMatrix(1, 19) = "KR Date"
MSHFlexGrid1.TextMatrix(1, 20) = "CDTF Print Flag"
MSHFlexGrid1.TextMatrix(1, 21) = "GSL Flag"

MSHFlexGrid1.RowHeight(1) = 500

End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   qry = " where " & Gen_WcState & ""
Else
   qry = ""
End If
Call TableMst_State(qry)
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
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
mSStateID = RsMst_State!StateID
End Sub

Private Sub CmdExcel_Click()

Dim ans As Variant


Gen_mTimeStamp = GetTimeStamp


ans = MsgBox("Do you want to generate CSV for searched result? Click 'Yes' for CSV and 'No' for Excel generation!!!", vbYesNo)
If ans = vbYes Then
   Call GenerateCSV
Else
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GSLCDTFvsTRFDate.xls")
   Call Xls_Detail_Rpt
End If


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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLCDTFvsTRFDate.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MMM/dd/yyyy")
           Else
              If c >= 3 And c <= 7 Then
                  oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid1.TextMatrix(i, c)
              Else
                 oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
              End If
           End If
        Else
           If c >= 3 And c <= 7 Then
              oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid1.TextMatrix(i, c)
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdSearch_Click()

Dim Wc As String
Dim x As Double

ProgressBar1.Value = 0
Call Grid_Head

MsgBox "Wait till next message"

lblRecordCount.Caption = ""
Wc = ""
Wc = " Where Mst_GSL.ExchangeTypeID = 2 "

If CmbStateID.Text <> "" Then
   Wc = GenWc(Wc, mSStateID, "Mst_State.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp = Replace(Gen_WcState, "StateID", "Mst_State.StateID")
      Wc = GenWc(Wc, tmp, "", "N", "I")
   End If
End If

If CmbPreFix.Text <> "" Then
   Wc = GenWc(Wc, mSLocationID, "Mst_Location.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "Mst_Location.LocationID")
      Wc = GenWc(Wc, tmp, "", "N", "I")
   End If
End If

If CmbSCmp.Text <> "" Then
   Wc = GenWc(Wc, mSCMPID, "Mst_CMP.CMPID", "N", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp = Replace(Gen_WcCMP, "CMPID", "mst_CMP.CMPID")
      Wc = GenWc(Wc, tmp, "", "N", "I")
   End If
End If

If CmbSGodown.Text <> "" Then
   Wc = GenWc(Wc, mSGodownID, "Mst_Godown.GodownID", "N", "N")
End If

If CmbSCommodity.Text <> "" Then
   Wc = GenWc(Wc, mSCommodityID, "Mst_Commodity.CommodityID", "N", "N")
End If

'If CmbPreFix.Text <> "" Then
'   Wc = GenWc(Wc, msSM_prefix, "Mst_GSL.sm_prefix", "S", "N")
'End If

If TxtSClientName.Text <> "" Then
   Wc = GenWc(Wc, TxtSClientName, "Mst_Client.ClientName", "S", "L")
End If

If TxtSCDTFNo.Text <> "" Then
   Wc = GenWc(Wc, TxtSCDTFNo, "TCD.CxTFNo", "S", "L")
End If

If Check2.Value = 1 Then
    If TxtToDate < TxtFromDate Then
        MsgBox "'From Date' Must be less than 'To Date' !!!"
        TxtFromDate.SetFocus
        Exit Sub
    End If
    If Gen_DBType = "MDB" Then
       If Wc <> "" Then
          Wc = Wc & " And (TCD.CxTFDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
       Else
          Wc = " Where (TCD.CxTFDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
       End If
    Else
       If Wc <> "" Then
          Wc = Wc & " And (TCD.CxTFDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
       Else
          Wc = " Where (TCD.CxTFDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "') "
       End If
    End If
End If

If TxtSGSLID <> "" Then
   Wc = GenWc(Wc, TxtSGSLID, "Mst_GSL.GSLID", "N", "N")
End If





tmp = " SELECT Mst_GSL.GSLID, Mst_State.StateName,Mst_Location.LocationName , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo,"
tmp = tmp & " Mst_GSL.StackNo, Mst_GSL.LotNo , Mst_GSL.DepositeBags,Mst_GSL.DepositeWeight,"
tmp = tmp & " Replace(Mst_Client.ClientName,Char(10),'') 'ClientName', Replace(MA.AgentName,Char(10),'') 'AgentName',"
tmp = tmp & " Max(TCD.CxTFNo) as 'CDTFNo', Max(TCD.CxTFDate) as 'CDTFDate',MAx(TCD.CreatedDate) as 'CDTFCreatedDate',TRF.TRFNo,"
tmp = tmp & " TRF.SamplingDate,TRF.CreatedDate as 'TRFCreatedDate', Mst_GSL.KRNo, Mst_GSL.KRDate ,Mst_GSL.CDTFNo as 'GSLCDTFNo',Mst_GSL.CDTFDate as 'GSLCDTFDate',"
'tmp = tmp & " Mst_GSL.BGSLID , Mst_GSL.FGSLID, Mst_GSL.BSM_Prefix, Mst_GSL.FSM_Prefix"
tmp = tmp & " TCD.Flag as 'CDTFPrintFlag',Mst_GSL.Flag as 'GSLFlag'"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " Left JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
tmp = tmp & " Left JOIN Mst_Location on Mst_CMP.LocationID = Mst_Location.LocationID"
tmp = tmp & " Left JOIN Mst_State on Mst_Location.StateID = Mst_State.StateID"
tmp = tmp & " Left JOIN Mst_Godown on Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " Left JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
tmp = tmp & " Left JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
tmp = tmp & " Left Join Txn_CxTF_GSL TCG On TCG.GSLID  = Mst_GSL.GSLID And TCG.SM_Prefix = Mst_GSL.SM_Prefix And TCG.TxnType = 'D'"
tmp = tmp & " Left Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_PRefix = TCD.SM_Prefix And TCD.TxnType = 'D'"
tmp = tmp & " Left Join Mst_Agent MA On TCD.AgentID = MA.AgentID"
tmp = tmp & " Left Join Txn_TRFDetail TRFD On Mst_GSL.GSLID  = TRFD.GSLID And Mst_GSL.SM_Prefix = TRFD.SM_Prefix"
tmp = tmp & " Left Join Txn_TRF TRF On TRFD.TRFID = TRF.TRFID"
tmp = tmp & " " & Wc
tmp = tmp & " Group by Mst_GSL.GSLID, Mst_State.StateName,Mst_Location.LocationName , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo,"
tmp = tmp & " Mst_GSL.StackNo, Mst_GSL.LotNo , Mst_GSL.DepositeBags,Mst_GSL.DepositeWeight,Mst_GSL.CDTFNo,Mst_GSL.CDTFDate,"
tmp = tmp & " Mst_Client.ClientName,MA.AgentName,"
tmp = tmp & " TRF.TRFNo,TRF.SamplingDate,TRF.CreatedDate , Mst_GSL.KRNo, Mst_GSL.KRDate ,"
tmp = tmp & " TCD.Flag,Mst_GSL.Flag"
tmp = tmp & " ORDER BY MST_State.StateName,Mst_Location.LocationName,Mst_CMP.CMPName,Mst_GSL.GSLID"

Call TmpTable

''If Check3.Value = 1 Then
'tmp = "Select MG.GSLID,MG.SM_PRefix,MG.G_UID,Max(CDTFNo) as 'MaxCDTFNo' ,Max(TCD.CxTFDate) as  'MaxCDTFDate',Max(TCD.CreatedDate) as 'MaxCreatedDate' "
'tmp = tmp & " From Txn_CxTF_GSL TCG"
'tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_PRefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
'tmp = tmp & " Inner Join Mst_GSL MG On TCG.GSLID  = MG.GSLID And TCG.SM_Prefix = MG.SM_Prefix"
'tmp = tmp & " Where TCG.TxnType = 'D' "
'
'If CmbPreFix.Text <> "" Then
'   tmp = tmp & " And TCG.SM_PRefix = '" & msSM_prefix & "'"
'End If
'
'If TxtSGSLID <> "" Then
'   tmp = tmp & " And TCG.GSLID = " & Val(TxtSGSLID) & ""
'End If
'
'tmp = tmp & " Group By MG.GSLID,MG.SM_PRefix,MG.G_UID"
'
'Call Tmp1Table

MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
ProgressBar1.Max = TmpRs.RecordCount + 3
For i = 2 To TmpRs.RecordCount + 1
    'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1 'TmpRs.RecordCount + 2
'    For c = 0 To TmpRs.Fields.Count - 3
'        MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'    Next
  
    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs!GSLID), "", TmpRs!GSLID) '
    MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName) '"State Name"
    MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location"
    MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) '"CMP Name"
    MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity) '"Commodity"
    MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo) '"Godown No"
    MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo) '"Stack No"
    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo) '"Lot No"
    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(TmpRs!DepositeBags), "", TmpRs!DepositeBags) '"Deposited Bags"
    MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(TmpRs!DepositeWeight), "", TmpRs!DepositeWeight) '"Deposited Weight"
    MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) ' "Client Name"
    MSHFlexGrid1.TextMatrix(i, 11) = IIf(IsNull(TmpRs!AgentName), "", TmpRs!AgentName) '"Agent Name"
    MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs!CdtfNo), "", TmpRs!CdtfNo) '"CDTF No"
    MSHFlexGrid1.TextMatrix(i, 13) = IIf(IsNull(TmpRs!CDTFDate), "", TmpRs!CDTFDate) '"CDTF Date"
    MSHFlexGrid1.TextMatrix(i, 14) = IIf(IsNull(TmpRs!CDTFCreatedDate), "", TmpRs!CDTFCreatedDate) '"CDTF Created Date"
    MSHFlexGrid1.TextMatrix(i, 15) = IIf(IsNull(TmpRs!TRFNo), "", TmpRs!TRFNo) '"TRF No"
    MSHFlexGrid1.TextMatrix(i, 16) = IIf(IsNull(TmpRs!SamplingDate), "", TmpRs!SamplingDate) '"TRF Date"
    MSHFlexGrid1.TextMatrix(i, 17) = IIf(IsNull(TmpRs!TRFCreatedDate), "", TmpRs!TRFCreatedDate) '"TRF Created Date"
    MSHFlexGrid1.TextMatrix(i, 18) = IIf(IsNull(TmpRs!KRNo), "", TmpRs!KRNo) '"KR No"
    MSHFlexGrid1.TextMatrix(i, 19) = IIf(IsNull(TmpRs!KRDate), "", TmpRs!KRDate) '"KR Date"
    
    MSHFlexGrid1.TextMatrix(i, 20) = IIf(IsNull(TmpRs!CDTFPrintFlag), "", TmpRs!CDTFPrintFlag) '"KR Date"
    MSHFlexGrid1.TextMatrix(i, 21) = IIf(IsNull(TmpRs!GSLFlag), "", TmpRs!GSLFlag) '"KR Date"
    
    
    If MSHFlexGrid1.TextMatrix(i, 12) = "" Then
       MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs!GSLCdtfNo), "", TmpRs!GSLCdtfNo) '"CDTF No"
       MSHFlexGrid1.TextMatrix(i, 13) = IIf(IsNull(TmpRs!GSLCDTFDate), "", TmpRs!GSLCDTFDate) '"CDTF Date"
    End If

    MSHFlexGrid1.TextMatrix(i, 5) = Replace(Replace(MSHFlexGrid1.TextMatrix(i, 5), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(i, 6) = Replace(Replace(MSHFlexGrid1.TextMatrix(i, 6), vbCrLf, ""), ",", "~")
    MSHFlexGrid1.TextMatrix(i, 7) = Replace(Replace(MSHFlexGrid1.TextMatrix(i, 7), vbCrLf, ""), ",", "~")
     
    MSHFlexGrid1.TextMatrix(i, 8) = Format(MSHFlexGrid1.TextMatrix(i, 8), "########0")
    MSHFlexGrid1.TextMatrix(i, 9) = Format(MSHFlexGrid1.TextMatrix(i, 9), "########0.000")
    
    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(i, 8))
    MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(i, 9))
    
    MSHFlexGrid1.TextMatrix(i, 13) = Format(MSHFlexGrid1.TextMatrix(i, 13), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(i, 14) = Format(MSHFlexGrid1.TextMatrix(i, 14), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(i, 16) = Format(MSHFlexGrid1.TextMatrix(i, 16), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(i, 17) = Format(MSHFlexGrid1.TextMatrix(i, 17), "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(i, 19) = Format(MSHFlexGrid1.TextMatrix(i, 19), "dd-mmm-yyyy")
    
   TmpRs.MoveNext
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
lblRecordCount.Caption = TmpRs.RecordCount
MsgBox "Done !!! "

End Sub

Private Sub GenerateCSV()
Dim Path, strFilName, strFilLocation As Variant
Path = App.Path
strFilName = "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLCDTFvsTRFDate.csv"
strFilLocation = Path + strFilName
If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
Open strFilLocation For Output As #1

For i = 1 To MSHFlexGrid1.Rows - 1
    tmp = ""
    For c = 0 To MSHFlexGrid1.Cols - 1
        If tmp = "" Then
           tmp = tmp & MSHFlexGrid1.TextMatrix(i, c)
        Else
           tmp = tmp & "," & MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
    Print #1, tmp
Next
MsgBox "Done !!!"
Close #1
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame0.Enabled = False
   lblRecordCount.Caption = ""
   Exit Sub
End If
Frame0.Enabled = True
lblRecordCount.Caption = ""
Call CmdClear_Click

End Sub
