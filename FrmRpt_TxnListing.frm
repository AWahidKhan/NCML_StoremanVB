VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_TxnListing 
   Caption         =   "Transaction Detail Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   9480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbStateID 
         Height          =   315
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   345
         Width           =   2580
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3405
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   350
         Left            =   13935
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Go"
         Height          =   350
         Left            =   11865
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12960
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox CmbSExchangeType 
         Height          =   315
         Left            =   12270
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2640
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   870
         Width           =   5355
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   8940
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   3285
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   75
         TabIndex        =   0
         Top             =   345
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
         Format          =   109641729
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1515
         TabIndex        =   1
         Top             =   345
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
         Format          =   109641729
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.ComboBox CmbSAgent 
         Height          =   315
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   870
         Width           =   6030
      End
      Begin VB.Label LblStateID 
         Caption         =   "State Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3623
         TabIndex        =   23
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Location"
         Height          =   255
         Left            =   6855
         TabIndex        =   20
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "From  Date"
         Height          =   255
         Left            =   397
         TabIndex        =   19
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "To  Date"
         Height          =   225
         Left            =   1868
         TabIndex        =   18
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Total Record Found :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7650
         TabIndex        =   17
         Top             =   1230
         Width           =   1815
      End
      Begin VB.Label LblTotalRecord 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9615
         TabIndex        =   16
         Top             =   1230
         Width           =   75
      End
      Begin VB.Label LblAgentID 
         Caption         =   "Agent Name"
         Height          =   300
         Left            =   8070
         TabIndex        =   15
         Top             =   675
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Exchange Type"
         Height          =   300
         Left            =   12975
         TabIndex        =   14
         Top             =   150
         Width           =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   2317
         TabIndex        =   13
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "CMP"
         Height          =   255
         Left            =   10335
         TabIndex        =   12
         Top             =   150
         Width           =   480
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7695
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13573
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_TxnListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mCMPID, mSM_Prefix, mCommodityID, mExchangeTypeID, mAgentName, mLocationName, mLocationID As Variant
Dim NoRecord As Boolean
Dim OpenBalWc, wc As Variant

Private Sub CmbSLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Select State !!!"
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbSLocation.Text
Call TableMst_Location(" Where StateID = " & mStateID & "")
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbSCmp.Text = ""
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

If tmp <> CmbSLocation.Text Then
   CmbSCmp.Text = ""
   mCMPID = Null
End If
mLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!Locationname
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Call ClearFrame
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp

'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_TransactionDetails.xls")
'Call Xls_Detail_Rpt


tmp = ""
ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If ans = vbNo Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_TransactionDetails.xls")
   Call Xls_Detail_Rpt
Else
   Path = App.Path
   strFilName = "\Excel\" & Gen_mTimeStamp & "Rpt_TransactionDetails" & Gen_UserName & ".csv"
   strFilLocation = Path + strFilName
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
   Open strFilLocation For Output As #1
   For I = 0 To MSHFlexGrid1.Rows - 1
       tmp = ""
       For C = 0 To MSHFlexGrid1.Cols - 1
          If C = 0 Then
             tmp = tmp & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
'          ElseIf c = 6 Then
'             tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid1.TextMatrix(I, c), vbNewLine, ""), ",", "~")
          Else
             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          End If
       Next
       Print #1, tmp
   Next
   Close #1
   MsgBox "CSV Generated!!!"
End If


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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_TransactionDetails.xls")
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

Private Sub CmdSearch_Click()

Dim wc As String
Call Grid_Head
LblTotalRecord.Caption = ""
MsgBox ("Wait till Next Message")
ProgressBar1.Value = 0
If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox "To date must be always greater than from date!!!"
   TxtToDate.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
   If CmbStateID.Text = "" Then
      MsgBox "Blank state name not allowed!!!"
      CmbStateID.SetFocus
      Exit Sub
   End If
End If

'---Commented on 26 Oct 2010
'If CmbSLocation.Text <> "" Then
'   wc = GenWc(wc, mLocationName, "Mst_Location.LocationName", "S")
'End If



wc = ""

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, CmbStateID.Text, "Mst_State.StateName", "S")
End If

If CmbSLocation.Text <> "" Then
   wc = GenWc(wc, mSM_Prefix, "Txn_CxTf_Gsl.SM_Prefix", "S")
End If

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, mCMPID, "Mst_CMP.CMPID", "N")
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, mCommodityID, "Mst_Commodity.CommodityID", "N")
End If
If CmbSExchangeType.Text <> "" Then
   wc = GenWc(wc, mExchangeTypeID, "Mst_ExchangeType.ExchangeTypeID", "N")
End If


If CmbSAgent.Text <> "" Then
   wc = GenWc(wc, mAgentName, "Mst_Agent.AgentName", "S")
End If

tmp = ""

'---Commented on 26 Oct 2010
'tmp = " Select Mst_Location.LocationName,Mst_CMP.CMPName,Mst_Commodity.Commodity,Txn_CxTf_Details.TxnType,Txn_CxTf_Details.CxTFNo,Txn_CxTf_Details.ManualCDTFNo,"
'tmp = tmp & " Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Txn_CxTf_Details.Flag,Mst_Agent.AgentName ,"
'tmp = tmp & " Sum (Txn_CxTf_Gsl.NoofBags) 'NoOfBags',Sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight'"
'tmp = tmp & " From Txn_CxTf_Details"
'tmp = tmp & " Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And"
'tmp = tmp & " Txn_CxTf_Details.CxTFno = Txn_CxTf_Gsl.CxTFno And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType"
'tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
'tmp = tmp & " Inner Join Mst_Location On Mst_CMP.LocationID = Mst_Location.LocationID"
'tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
'tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
'tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
''tmp = tmp & " Where Txn_CxTf_Details.CxTFDate Between '2009/01/06' And '2010/01/06'"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Where Txn_CxTf_Details.CxTFDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
'Else
'   tmp = tmp & " Where Txn_CxTf_Details.CxTFDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "' "
'End If
'
'If wc <> "" Then
'   tmp = tmp & wc
'End If
'
'tmp = tmp & " Group by Txn_CxTf_Details.CxTFNo,Txn_CxTf_Details.TxnType,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,"
'tmp = tmp & " Txn_CxTf_Details.SM_Prefix, Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
'tmp = tmp & " Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo, Mst_Location.LocationName"
'tmp = tmp & " ORDER BY Mst_Location.LocationName,Mst_CMP.CMPName,Mst_Commodity.Commodity,Mst_ExchangeType.ExchangeType,"
'tmp = tmp & " Txn_CxTf_Details.CxTFDate , Txn_CxTf_Details.CxTFno"
'
'Call TmpTable
'-------------------------------------------------------------------------



'Parag
'tmp = " Select Mst_State.StateName,Mst_Location.LocationName,Mst_CMP.CMPName,Mst_Commodity.Commodity,Txn_CxTf_Gsl.TxnType,Txn_CxTf_Gsl.CxTFNo,Txn_CxTf_Details.ManualCDTFNo,"
'tmp = tmp & " Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Txn_CxTf_Details.Flag,Mst_Agent.AgentName ,MC.ClientName+'~'+MC.ClientID 'ClientName',"
'tmp = tmp & " Sum (Txn_CxTf_Gsl.NoofBags) 'NoOfBags',Sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight'"
'tmp = tmp & " From Txn_CxTf_Details"
'tmp = tmp & " Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And"
'tmp = tmp & " Txn_CxTf_Details.CxTFno = Txn_CxTf_Gsl.CxTFno And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType"
'tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
'tmp = tmp & " Inner Join Mst_Location On Mst_CMP.LocationID = Mst_Location.LocationID"
'tmp = tmp & " Inner Join Mst_State On Mst_Location.StateID = Mst_State.StateID"
'tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
'tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
'tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
'tmp = tmp & " Left Join Mst_GSL MGSL ON Txn_CxTf_Gsl.GSLID = MGSL.GSLID And Txn_CxTf_Gsl.SM_Prefix = MGSL.SM_Prefix "
'tmp = tmp & " Left Join Mst_Client MC ON MGSL.ClientIDRow = MC.ClientIDRow"
''tmp = tmp & " Where Txn_CxTf_Details.CxTFDate Between '2009/01/06' And '2010/01/06'"
'
''If Gen_DBType = "MDB" Then
''   tmp = tmp & " Where Txn_CxTf_Details.CxTFDate Between #" & Format(TxtFromDate, Gen_DatFormat) & "# And #" & Format(TxtToDate, Gen_DatFormat) & "#) "
''Else
''   tmp = tmp & " Where Txn_CxTf_Details.CxTFDate Between '" & Format(TxtFromDate, Gen_DatFormat) & "' And '" & Format(TxtToDate, Gen_DatFormat) & "' "
''End If
'
''tmp = tmp & " Where "
'
'If wc <> "" Then
'   tmp = tmp & " " & wc
'End If
'
'tmp = tmp & " Group by Txn_CxTf_Gsl.CxTFNo,Txn_CxTf_Gsl.TxnType,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,"
'tmp = tmp & " Txn_CxTf_Details.SM_Prefix, Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
'tmp = tmp & " Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo, Mst_Location.LocationName,Mst_State.StateName,MC.ClientName,MC.ClientID"
'tmp = tmp & " ORDER BY Mst_State.StateName,Mst_Location.LocationName,Mst_CMP.CMPName,Mst_Commodity.Commodity,Mst_ExchangeType.ExchangeType,"
'tmp = tmp & " Txn_CxTf_Details.CxTFDate , Txn_CxTf_Gsl.CxTFno"

tmp = " Select Mst_State.StateName,Mst_Location.LocationName,Mst_CMP.CMPName,Mst_Commodity.Commodity,Txn_CxTf_Gsl.TxnType,Txn_CxTf_Gsl.CxTF_DetailsID,Txn_CxTf_Details.ManualCDTFNo,"
tmp = tmp & " Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Txn_CxTf_Details.Flag,Mst_Agent.AgentName ,MC.ClientName+'~'+MC.ClientIDRow 'ClientName',"
tmp = tmp & " Sum (Txn_CxTf_Gsl.NoofBags) 'NoOfBags',Sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight'"
tmp = tmp & " From Txn_CxTf_Details"
tmp = tmp & " Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And"
tmp = tmp & " Txn_CxTf_Details.CxTF_DetailsID = Txn_CxTf_Gsl.CxTF_DetailsID And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType"
tmp = tmp & " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID"
tmp = tmp & " Inner Join Mst_Location On Mst_CMP.LocationID = Mst_Location.LocationID"
tmp = tmp & " Inner Join Mst_State On Mst_Location.StateID = Mst_State.StateID"
tmp = tmp & " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID"
tmp = tmp & " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID"
tmp = tmp & " Left Join Mst_GSL MGSL ON Txn_CxTf_Gsl.Mst_GSL_ID = MGSL.Mst_GSL_ID And Txn_CxTf_Gsl.SM_Prefix = MGSL.SM_Prefix "
tmp = tmp & " Left Join Mst_Client MC ON MGSL.ClientID = MC.ClientID"
If wc <> "" Then
   tmp = tmp & " " & wc
End If
tmp = tmp & " Group by Txn_CxTf_Gsl.CxTF_DetailsID,Txn_CxTf_Gsl.TxnType,Txn_CxTf_Details.CxTFDate,Mst_ExchangeType.ExchangeType,Mst_CMP.CMPName,"
tmp = tmp & " Txn_CxTf_Details.SM_Prefix, Mst_Commodity.Commodity,Txn_CxTf_Details.Flag,"
tmp = tmp & " Mst_Agent.AgentName , Txn_CxTf_Details.ManualCDTFNo, Mst_Location.LocationName,Mst_State.StateName,MC.ClientName,MC.ClientIDRow"
tmp = tmp & " ORDER BY Mst_State.StateName,Mst_Location.LocationName,Mst_CMP.CMPName,Mst_Commodity.Commodity,Mst_ExchangeType.ExchangeType,"
tmp = tmp & " Txn_CxTf_Details.CxTFDate , Txn_CxTf_Gsl.CxTF_DetailsID"

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 1

For ix = 1 To TmpRs.RecordCount
    If TmpRs!CxTFDate >= TxtFromDate And TmpRs!CxTFDate <= TxtToDate Then
       I = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 7) = Format(MSHFlexGrid1.TextMatrix(I, 7), "dd-mmm-yyyy")
       If LCase(MSHFlexGrid1.TextMatrix(I, 4)) = "d" Then
          MSHFlexGrid1.TextMatrix(I, 12) = Format(IIf(IsNull(TmpRs!NoofBags), "", TmpRs!NoofBags), "#######0")
          MSHFlexGrid1.TextMatrix(I, 13) = Format(IIf(IsNull(TmpRs!GrossWeight), "", TmpRs!GrossWeight), "#######0.000")
          MSHFlexGrid1.TextMatrix(I, 14) = "" 'IIf(IsNull(TmpRs.Fields(10)), "", TmpRs.Fields(10))
          MSHFlexGrid1.TextMatrix(I, 15) = "" 'IIf(IsNull(TmpRs.Fields(10)), "", TmpRs.Fields(10))
       Else
          MSHFlexGrid1.TextMatrix(I, 12) = "" 'IIf(IsNull(TmpRs.Fields(10)), "", TmpRs.Fields(10))
          MSHFlexGrid1.TextMatrix(I, 13) = "" 'IIf(IsNull(TmpRs.Fields(11)), "", TmpRs.Fields(11))
          MSHFlexGrid1.TextMatrix(I, 14) = Format(IIf(IsNull(TmpRs!NoofBags), "", TmpRs!NoofBags), "#######0")
          MSHFlexGrid1.TextMatrix(I, 15) = Format(IIf(IsNull(TmpRs!GrossWeight), "", TmpRs!GrossWeight), "#######0.000")
       End If
       
       MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 12)) 'Format(IIf(IsNull(TmpRs!NoofBags), "", TmpRs!NoofBags), "#######0.000")
       MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(I, 13)) 'Format(IIf(IsNull(TmpRs!NoofBags), "", TmpRs!NoofBags), "#######0.000")
       MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(I, 14)) 'Format(IIf(IsNull(TmpRs!NoofBags), "", TmpRs!NoofBags), "#######0.000")
       MSHFlexGrid1.TextMatrix(0, 15) = Val(MSHFlexGrid1.TextMatrix(0, 15)) + Val(MSHFlexGrid1.TextMatrix(I, 15)) 'Format(IIf(IsNull(TmpRs!NoofBags), "", TmpRs!NoofBags), "#######0.000")
       
    End If
    
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
    'If TmpRs.EOF Then Exit For
    
    'If TmpRs!CxtfDate >= TxtToDate Then
    '   Exit For
    'End If
Next

ProgressBar1.Value = ProgressBar1.Max

LblTotalRecord.Caption = MSHFlexGrid1.Rows - 2 'TmpRs.RecordCount
MsgBox ("Done")

End Sub
Function GenWc(wc_ As String, ID As Variant, Field_ As String, Type_ As String) As String
Dim Retval
Retval = wc_
If Retval = "" Then
   If Type_ = "S" Then
      Retval = "Where " & Field_ & "= '" & ID & "'"
   Else
      Retval = "Where " & Field_ & "= " & ID
   End If
Else
   If Type_ = "S" Then
      Retval = Retval & " And " & Field_ & "= '" & ID & "'"
   Else
      Retval = Retval & " And " & Field_ & "= " & ID
   End If
End If
GenWc = Retval
End Function

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
CmbSCmp.Text = ""
CmbSCommodity.Text = ""
CmbSExchangeType.Text = ""
CmbSAgent.Text = ""
LblTotalRecord.Caption = ""
ProgressBar1.Value = 0
'TxtFromDate.SetFocus

End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 3
     .Cols = 16 '15
     .Font.Size = 9
     .WordWrap = True
'     .FixedRows = 2
     .AllowUserResizing = flexResizeColumns
     
     .TextMatrix(1, 0) = "State"
     .TextMatrix(1, 1) = "Location"
     .TextMatrix(1, 2) = "CMP Name"
     .TextMatrix(1, 3) = "Commodity Name"
     .TextMatrix(1, 4) = "Txn Type"
     .TextMatrix(1, 5) = "CxTF No"
     .TextMatrix(1, 6) = "Manual CxTF No"
     .TextMatrix(1, 7) = "Date"
     .TextMatrix(1, 8) = "Exchange Type"
     .TextMatrix(1, 9) = "Print Flag"
     .TextMatrix(1, 10) = "Agent"
     .TextMatrix(1, 11) = "Client Name"
     .TextMatrix(1, 12) = "Deposit No of Bags"
     .TextMatrix(1, 13) = "Deposit Weight "
     .TextMatrix(1, 14) = "Withdrawal No Of Bags"
     .TextMatrix(1, 15) = "Withdrawal Weight"
          
     .ColWidth(0) = 1200
     .ColWidth(1) = 1200
     .ColWidth(2) = 1200
     .ColWidth(3) = 1200
     .ColWidth(4) = 400 '800
     .ColWidth(5) = 800 '
     .ColWidth(6) = 900
     .ColWidth(7) = 1200
     .ColWidth(8) = 1000
     .ColWidth(9) = 600
     .ColWidth(10) = 2000
     .ColWidth(11) = 1000
     .ColWidth(12) = 1000
     .ColWidth(13) = 1000
     .ColWidth(14) = 1000
     .ColWidth(15) = 1000
     
     .RowHeight(1) = 600
End With
End Sub
Private Sub CmbSCmp_GotFocus()
If CmbSLocation.Text = "" Then
    MsgBox "Select Location!!!"
    CmbSLocation.SetFocus
    Exit Sub
End If
tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mLocationID & " ")
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
If CmbSCmp.Text = "" Then Exit Sub

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
mSM_Prefix = GetPreFix(RsMst_CMP!CMPID, "P")
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
'CmbSAgent.Text = ""
'mAgentID = Null
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

Private Sub CmbSAgent_GotFocus()
'If CmbSExchangeType.Text = "" Then Exit Sub
tmp = CmbSAgent.Text
Call TableMst_Agent '("Where ExchangeTypeId = " & mExchangeTypeID & "")
CmbSAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Agent.RecordCount
    CmbSAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbSAgent.Text = tmp

End Sub

Private Sub CmbSAgent_LostFocus()

If CmbSAgent.Text = "" Then
   mAgentID = Null
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "AgentName = '" & CmbSAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection!!! "
   CmbSAgent.SetFocus
   Exit Sub
End If
mAgentName = RsMst_Agent!AgentName

End Sub


Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState = "" Then
   Call TableMst_State
Else
   Call TableMst_State(" Where " & Gen_WcState)
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
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbSLocation.Text = ""
   CmbSCmp.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If mStateID <> RsMst_State!StateID Then
   CmbSLocation.Text = ""
   CmbSCmp.Text = ""
End If

mStateID = RsMst_State!StateID

End Sub
