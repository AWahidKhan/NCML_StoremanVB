VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_DepositContractProcess 
   Caption         =   "Deposit Contract Auto Entry Process"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13230
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   15030
      Begin VB.CheckBox ChkFilter 
         Caption         =   "Apply Filter"
         Height          =   255
         Left            =   6960
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox CmbLocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   3750
      End
      Begin VB.ComboBox CmbState 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmTxn_DepositContractProcess.frx":0000
         Left            =   120
         List            =   "FrmTxn_DepositContractProcess.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3750
      End
      Begin VB.CommandButton CmbExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   13920
         TabIndex        =   7
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Process"
         Height          =   375
         Left            =   12960
         TabIndex        =   6
         Top             =   180
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3240
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   5715
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         HighLight       =   0
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   5520
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   9737
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         HighLight       =   0
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin VB.CommandButton CmdGetData 
         Caption         =   "Get Data"
         Height          =   375
         Left            =   12000
         TabIndex        =   5
         Top             =   180
         Width           =   975
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   111869953
         CurrentDate     =   39850
      End
      Begin MSComCtl2.DTPicker TxtEndDate 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   111869953
         CurrentDate     =   39850
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label30 
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
         Height          =   195
         Left            =   11055
         TabIndex        =   13
         Top             =   255
         Width           =   75
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
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
         Index           =   1
         Left            =   9120
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   10080
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "FrmTxn_DepositContractProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mWC As Variant
Dim mMaxYearMonth As Variant
Dim mSStateID As Double, mSLocationID As Double
Dim mStateID As Double, mLocationID As Double, mClientID As Double, mCommodityID As Double, mCommodityGroupId As Double
Dim mSMPrefix, mFumigationFlag, mMinBagSize, mMaxBagSize, mBagEquivalent As Variant
Dim mStartDate As Variant, mEndDate As Variant
Dim mLastRow As Double

Dim mCurrentConversionFactor As Variant, mDivisionFactor  As Variant, mFumigationFactorY  As Variant, mFumigationFactorN  As Variant
Dim mCMPMtoAMRMApprovalLimit As Variant, mKnownRentMinimumValue  As Variant, mDepositContractCMGGroupToBeApproved  As Variant, mDepositContractDefaultDaysRange  As Variant
Dim mOwnGodownSqrFeetRate As Variant
Dim mBillingOnePercentage, mDepositOnePercentage As Variant

Dim mBCycle As Variant, mDCycle As Variant
Dim mBillingCycleID As Double, mDepositBillingCycleID As Double
Dim mBillingRatio As Double, mDepositBillingRatio As Double
Dim mBillingCycleNoofDays As Double, mDepositBillingCycleNoofDays As Double
Dim mMinSize As Double, mMaxSize As Double, mRate As Double
Dim mLastPRow As Double, mLastGRow As Double

Public Sub Grid_Head()
With MSHFlexGrid1
     .Font.Size = 9
     .Clear
     .Rows = 2
     .Cols = 25
     .FixedCols = 0
     .TextMatrix(0, 0) = "State Name":
     .TextMatrix(0, 1) = "StateID":
     .TextMatrix(0, 2) = "Location Name":
     .TextMatrix(0, 3) = "LocationID":
     .TextMatrix(0, 4) = "SM Prefix":
     .TextMatrix(0, 5) = "Commodity":
     .TextMatrix(0, 6) = "CommodityID":
     .TextMatrix(0, 7) = "ClientName":
     .TextMatrix(0, 8) = "ClientRowID":
     .TextMatrix(0, 9) = "Balance Bags":
     .TextMatrix(0, 10) = "Balance Weight":
     .TextMatrix(0, 11) = "Fumigation Flag":
     .TextMatrix(0, 12) = "Minimum":
     .TextMatrix(0, 13) = "Maximum":
     .TextMatrix(0, 14) = "Bag Equivalent":
     .TextMatrix(0, 15) = "CommodityGroupID":
     .TextMatrix(0, 16) = "ContractOn":
     .TextMatrix(0, 17) = "BillingCycleID":
     .TextMatrix(0, 18) = "DWBillingCycleID":
     .TextMatrix(0, 19) = "RateCardType":
     .TextMatrix(0, 20) = "RateCardTypeID":
     .TextMatrix(0, 21) = "SrNo":
     .TextMatrix(0, 22) = "Billing Cycle":
     .TextMatrix(0, 23) = "DWBilling Cycle":
     .TextMatrix(0, 24) = "SysSrNo"
     
     .RowHeight(0) = 850
End With
End Sub

Private Sub GenerateCSV(Grid_ As MSHFlexGrid, mFile_ As Variant)
Dim Path, strFilName, strFilLocation As Variant
Dim mColorName As Variant
Path = App.Path
strFilName = "\excel\" & mFile_
strFilLocation = Path + strFilName
If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
Open strFilLocation For Output As #1

Select Case Grid_.Name
 Case "MSHFlexGrid1"
 
    For i = 0 To Grid_.Rows - 1
        tmp = ""
        For C = 0 To Grid_.Cols - 1
            If tmp = "" Then
               tmp = tmp & Grid_.TextMatrix(i, C)
            Else
               tmp = tmp & "," & Grid_.TextMatrix(i, C)
            End If
        Next
        Print #1, tmp
    Next
 
 Case "MSHFlexGrid2"
 
    For i = 0 To Grid_.Rows - 1
        tmp = ""
        mColorName = ""
        For C = 0 To Grid_.Cols - 1
            If tmp = "" Then
               tmp = tmp & Grid_.TextMatrix(i, C)
            Else
               If C = 3 Then tmp = tmp & ",""" & Grid_.TextMatrix(i, C) & """" Else tmp = tmp & "," & Grid_.TextMatrix(i, C)
               Grid_.row = i
               Grid_.Col = C
               If Grid_.CellBackColor = vbRed Then mColorName = mColorName & "Red"
               If Grid_.CellBackColor = vbGreen Then mColorName = mColorName & "Green"
               If Grid_.CellBackColor = vbYellow Then mColorName = mColorName & "Yellow"
            End If
        Next
        Print #1, tmp & "," & mColorName
    Next
 
End Select
Close #1
End Sub

Private Sub CmbExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Select Case MsgBox("Do you want to Export in CSV ?", vbYesNo + vbQuestion)
    Case vbNo
        MsgBox ("Wait till next message")
        Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_DepositContractProcess.xls")
        Call Xls_Detail_Rpt
    Case vbYes
        MsgBox ("Wait till next message")
        Call GenerateCSV(MSHFlexGrid1, Gen_mTimeStamp & Gen_UserName & "Txn_DepositContractProcess_Grid1.CSV")
        Call GenerateCSV(MSHFlexGrid2, Gen_mTimeStamp & Gen_UserName & "Txn_DepositContractProcess_Grid2.CSV")
        MsgBox ("2 CSV File is Exported in Excel Folder.")
End Select
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_DepositContractProcess.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
With MSHFlexGrid1
ProgressBar1.Max = .Rows - 1
For i = 0 To .Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To .Cols - 1
        If InStr(.TextMatrix(i, C), "/") > 0 Then
           If IsDate(.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = .TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = .TextMatrix(i, C)
        End If
    Next
ProgressBar1.Value = i
Next
End With

Set oWS = oWB.Worksheets("Sheet2")
oWS.ClearArrows
mRow = 1
With MSHFlexGrid2
ProgressBar1.Max = .Rows - 1
For i = 0 To .Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To .Cols - 1
        If InStr(.TextMatrix(i, C), "/") > 0 Then
           If IsDate(.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = .TextMatrix(i, C)
           End If
        Else
           If C = 2 Then
            oWS.Cells(mRow, C + 1) = "'" & .TextMatrix(i, C)
           Else
            oWS.Cells(mRow, C + 1) = .TextMatrix(i, C)
           End If
           If C = 10 Or C = 14 Then
              .row = i
              .Col = C
              If .CellBackColor = vbRed Then oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
              If .CellBackColor = vbGreen Then oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbGreen
              If .CellBackColor = vbYellow Then oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
           End If
        End If
    Next
ProgressBar1.Value = i
Next
End With

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbLocation_GotFocus()
If Trim(CmbState.Text) = "" Then
   MsgBox "Select State!!!"
   CmbState.SetFocus
   Exit Sub
End If
tmp1 = CmbLocation.Text
tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID =" & mSStateID & "")
Else
   Call TableMst_Location(" Where " & Gen_WcLocation & " And StateID = " & mSStateID & " ")
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
   CmbLocation.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
  Exit Sub
End If
If RsMst_Location.RecordCount > 0 Then RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
If RsMst_Location!LocationName <> CmbLocation Then
End If
mSLocationID = RsMst_Location.Fields("LocationID")
'mSM_Prefix = RsMst_Location.Fields("SM_Prefix")
End Sub

Private Sub CmbState_GotFocus()
tmp = CmbState.Text
tmp1 = tmp
Call TableMst_State(" Where Flag = 'A' " & IIf(Gen_WcState <> "", "And " & Gen_WcState, ""))
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record Found!"
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub

Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   CmbLocation.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If
If LCase(CmbState.Text) <> LCase(tmp1) Then
   CmbLocation.Text = ""
End If
mSStateID = RsMst_State!StateID
End Sub

Private Sub CmdGetData_Click()
Dim mR, mC As Double
Call Grid_Head
Call Grid_Head2

mWC = ""
If CmbLocation.Text <> "" Then mWC = " And MC.LocationID = " & mSLocationID & " "
mStartDate = Format(TxtStartDate.Value, "YYYY-MM-DD")
mEndDate = Format(TxtEndDate.Value, "YYYY-MM-DD")

Label30.Caption = "-"
tmp = " Select MS.StateName,ML.StateID,ML.LocationName,ML.LocationID,A.SM_Prefix,MC.Commodity,A.CommodityID,MC.CommodityGroupID,MC.FumigationFlag,MC.Minimum,MC.Maximum,Isnull(CommEql.Equivalent,1)Equivalent,Client.ClientName,Client.ClientID,A.ClientIDRow,A.Bags,A.Weight From ( "
tmp = tmp & " SELECT Distinct SM_Prefix,CommodityID,ClientIDRow,Sum(BalanceBags)Bags,Sum(BalanceWeight)Weight "
tmp = tmp & " From Mst_GSL GSL WHERE ExchangeTypeID = 1 And Flag not In ('E','Z','B') And (ContractID is Null or ContractID = 0) And BalanceBags > 0 "
tmp = tmp & " And SpillageFlag = '0' And MadeupFlag = '0' "

If ChkFilter.Value = vbChecked Then
    tmp = tmp & " And GSL.GodownID in ("
    tmp = tmp & " Select Distinct MG.GodownID From Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
    tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
    tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
    tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
    tmp = tmp & " Where MG.CloseDate is null And MG.StartDate<='" & mEndDate & "' And ( '" & mStartDate & "' <= TLA.ExpiryDate And '" & mEndDate & "' >= TLA.StartDate ) And  MU.UnitTypeID = 1 " & IIf(mWC = "", "", mWC) & " "
    tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)"
    tmp = tmp & " Union All"
    tmp = tmp & " Select MG.GodownID From Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
    tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
    tmp = tmp & " Where MG.CloseDate is null And MG.StartDate<='" & mEndDate & "' And '" & mEndDate & "' <= MG.AgreementEndDate And  MU.UnitTypeID <> 1 " & IIf(mWC = "", "", mWC) & " "
    tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)"
    tmp = tmp & " ) "
End If
tmp = tmp & " Group By SM_Prefix,CommodityID,ClientIDRow ) A"
tmp = tmp & " Inner Join Mst_Location ML On A.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_Commodity MC On A.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_Client Client On A.ClientIDRow = Client.ClientIDRow And Client.Status = 'A' "
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Left Join (Select CommodityID,MAX(Equivalent) 'Equivalent' from Mst_BagEquivalent Group By CommodityID) CommEql On CommEql.CommodityID = MC.CommodityID "
tmp = tmp & " Order By MS.StateName,ML.LocationName,MC.Commodity,Client.ClientName,A.Bags"
Call TmpTable

Label30.Caption = TmpRs.RecordCount
If TmpRs.RecordCount > 0 Then
With MSHFlexGrid1
   ProgressBar1.Max = TmpRs.RecordCount
   For mR = 1 To TmpRs.RecordCount
        .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!StateID), "", TmpRs!StateID)
        .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
        .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
        .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
        .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) & "#" & IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
        .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow)
        .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!Bags), "", TmpRs!Bags)
        .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!Weight), "", TmpRs!Weight)
        .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!FumigationFlag), "", TmpRs!FumigationFlag)
        .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!Minimum), "", TmpRs!Minimum)
        .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!Maximum), "", TmpRs!Maximum)
        .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!Equivalent), "", TmpRs!Equivalent)
        .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!CommodityGroupID), "", TmpRs!CommodityGroupID)
        .TextMatrix(mR, 21) = mR
        .Rows = .Rows + 1
        ProgressBar1.Value = mR
        TmpRs.MoveNext
   Next mR
End With

''--- Max Year Month
tmp = " Select MAX(YearMonth)[YearMonth] from Txn_InvMonGSLDetail "
Call TmpTable
If TmpRs.RecordCount > 0 Then mMaxYearMonth = IIf(IsNull(TmpRs!YearMonth), "", Format(TmpRs!YearMonth, "YYYY-MM-DD")) Else mMaxYearMonth = ""

If mMaxYearMonth = "" Then
    MsgBox "Max Year-Month not found."
    Exit Sub
End If
''---------------------------------------------------------

''---- Billing cycle / Deposite Cycle
''--1. First Special Rate Card
''--2. Second Base Rate Card if not in Special rate Card
tmp = " Select SRC.SPRCID,SRC.LocationID,SRC.CommodityGroupID,SRC.ClientIDRow,SRC.BillingUnit,SRC.BillingCycleID,SRC.DepoBillingCycleID,SRC.RateOnQty,SRC.Amount,MBC.BillingCycle,DMBC.BillingCycle[DBillingCycle]"
tmp = tmp & " From Mst_SpecialRateCard SRC"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On SRC.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Inner Join Mst_BillingCycle DMBC On SRC.DepoBillingCycleID = DMBC.BillingCycleID"
tmp = tmp & " Where SRC.Flag = 'A' And SRC.ExchangeTypeID  = 1 "
Call TmpTable
If TmpRs.RecordCount > 0 Then
    With MSHFlexGrid1
        For mR = 1 To .Rows - 1
            mLocationID = Val(.TextMatrix(mR, 3))
            mSMPrefix = Trim(.TextMatrix(mR, 4))
            mCommodityID = Val(.TextMatrix(mR, 6))
            mClientID = Val(.TextMatrix(mR, 8))
            mFumigationFlag = Trim(.TextMatrix(mR, 11))
            mBagEquivalent = .TextMatrix(mR, 14)
            mCommodityGroupId = Val(.TextMatrix(mR, 15))
                
            TmpRs.Filter = ""
            TmpRs.MoveFirst
            TmpRs.Filter = "LocationID = " & mLocationID & " And CommodityGroupID = " & mCommodityGroupId & " And ClientIDRow = " & mClientID
            If TmpRs.EOF = False Then
                .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!BillingUnit), "", TmpRs!BillingUnit)
                .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!BillingCycleID), "", TmpRs!BillingCycleID)
                .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!DepoBillingCycleID), "", TmpRs!DepoBillingCycleID)
                .TextMatrix(mR, 19) = "S"
                .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!SPRCID), "", TmpRs!SPRCID)
                .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)
                .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!DBillingCycle), "", TmpRs!DBillingCycle)
                
            End If
        Next mR
    End With
End If

tmp = " SELECT BRC.SRCID,BRC.LocationID,BRC.SM_Prefix,BRC.CommodityGroupID,BRC.Flag,BRC.BillingCycleID,BRC.DepoBillingCycleID,BRC.RateOnQty,MBC.BillingCycle,DMBC.BillingCycle[DBillingCycle]"
tmp = tmp & " FROM Mst_StorageRateCard BRC"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On BRC.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Inner Join Mst_BillingCycle DMBC On BRC.DepoBillingCycleID = DMBC.BillingCycleID"
tmp = tmp & " Where BRC.DefaultRate = 1 And BRC.ExchangeTypeID = 1"
Call TmpTable

If TmpRs.RecordCount > 0 Then
    With MSHFlexGrid1
        For mR = 1 To .Rows - 1
            mLocationID = Val(.TextMatrix(mR, 3))
            mSMPrefix = Trim(.TextMatrix(mR, 4))
            mCommodityID = Val(.TextMatrix(mR, 6))
            mClientID = Val(.TextMatrix(mR, 8))
            mFumigationFlag = Trim(.TextMatrix(mR, 11))
            mBagEquivalent = .TextMatrix(mR, 14)
            mCommodityGroupId = Val(.TextMatrix(mR, 15))
                
            If Trim(.TextMatrix(mR, 16)) = "" Then
                TmpRs.Filter = ""
                TmpRs.MoveFirst
                TmpRs.Filter = "LocationID = " & mLocationID & " And CommodityGroupID = " & mCommodityGroupId
                If TmpRs.EOF = False Then
                    .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag)
                    .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!BillingCycleID), "", TmpRs!BillingCycleID)
                    .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!DepoBillingCycleID), "", TmpRs!DepoBillingCycleID)
                    .TextMatrix(mR, 19) = "N"
                    .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!SRCID), "", TmpRs!SRCID)
                    .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)
                    .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!DBillingCycle), "", TmpRs!DBillingCycle)
                End If
            End If
        Next mR
    End With
End If
''---------------------------------------------------------

End If
If MSHFlexGrid1.Rows > 2 Then CmdProcess.Enabled = True Else CmdProcess.Enabled = False
End Sub


Private Sub CmdProcess_Click()
Dim mR, mRow, mC, mCol As Double
With MSHFlexGrid1
    Call Grid_Head2
    
    Call TableTxn_DepositContract               ''-- To Save DC
    Call TableTxn_DepositContractGodownDetail   ''-- To Save DC Details
    
    ''-- Billing Cycle Cache for Quantity Only
    Call TableMst_BillingCycle(" Where Flag <> 'R' Order By BillingCycleID ")
    ''-- Special Rate card Cache
    tmp = " Select SRC.SPRCID,SRC.LocationID,SRC.CommodityGroupID,SRC.ClientIDRow,SRC.BillingUnit,SRC.BillingCycleID,SRC.DepoBillingCycleID,SRC.RateOnQty,SRC.Amount"
    tmp = tmp & " From Mst_SpecialRateCard SRC Where SRC.Flag = 'A' And SRC.ExchangeTypeID  = 1 And SRC.BillingUnit = 'Q' "
    Call Tmp2Table
    ''-- Base Rate card Cache for Quantity Only
    tmp = " SELECT BRC.SRCID,BRC.LocationID,BRC.SM_Prefix,BRC.CommodityGroupID,BRC.Flag,BRC.BillingCycleID,BRC.DepoBillingCycleID,BRC.RateOnQty"
    tmp = tmp & " FROM Mst_StorageRateCard BRC Where BRC.DefaultRate = 1 And BRC.ExchangeTypeID = 1 ANd BRC.Flag = 'Q' "
    Call Tmp3Table
    
    ProgressBar1.Max = .Rows - 2
    For mR = 1 To .Rows - 2
    
        mLocationID = Val(.TextMatrix(mR, 3))
        mSMPrefix = Trim(.TextMatrix(mR, 4))
        mCommodityID = Val(.TextMatrix(mR, 6))
        mClientID = Val(.TextMatrix(mR, 8))
        mFumigationFlag = Trim(.TextMatrix(mR, 11))
        mMinBagSize = .TextMatrix(mR, 12)
        mMaxBagSize = .TextMatrix(mR, 13)
        mBagEquivalent = .TextMatrix(mR, 14)
        mCommodityGroupId = Val(.TextMatrix(mR, 15))
    
        mBillingCycleID = Val(.TextMatrix(mR, 17))
        mDepositBillingCycleID = Val(.TextMatrix(mR, 18))
    
        If RsMst_BillingCycle.RecordCount > 0 Then
           RsMst_BillingCycle.MoveFirst
           RsMst_BillingCycle.Find "BillingCycleID = " & mBillingCycleID
           If RsMst_BillingCycle.EOF = False Then
                mBCycle = IIf(IsNull(RsMst_BillingCycle!BillingCycle), "", RsMst_BillingCycle!BillingCycle)
                mBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
                mBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)
           End If
           
           RsMst_BillingCycle.MoveFirst
           RsMst_BillingCycle.Find "BillingCycleID = " & mDepositBillingCycleID
           If RsMst_BillingCycle.EOF = False Then
                mDCycle = IIf(IsNull(RsMst_BillingCycle!BillingCycle), "", RsMst_BillingCycle!BillingCycle)
                mDepositBillingRatio = IIf(IsNull(RsMst_BillingCycle!Ratio), 0, RsMst_BillingCycle!Ratio)
                mDepositBillingCycleNoofDays = IIf(IsNull(RsMst_BillingCycle!NoOfDays), 0, RsMst_BillingCycle!NoOfDays)
           End If
        End If
        
        mLastRow = MSHFlexGrid2.Rows - 2
        If Trim(.TextMatrix(mR, 16)) <> "" Then
        
        ''-------- Insert data in Primary table
            RsTxn_DepositContract.AddNew
            RsTxn_DepositContract!DepConID = GetMaxID("DepConID", "Txn_DepositContract")
            .TextMatrix(mR, 24) = RsTxn_DepositContract!DepConID
            RsTxn_DepositContract!Date = Date
            RsTxn_DepositContract!StateID = Val(.TextMatrix(mR, 1))
            RsTxn_DepositContract!LocationID = Val(.TextMatrix(mR, 3))
            RsTxn_DepositContract!ClientID = Val(.TextMatrix(mR, 8))
            RsTxn_DepositContract!CommodityID = Val(.TextMatrix(mR, 6))
            RsTxn_DepositContract!QuantumMT = 1
            RsTxn_DepositContract!StartDate = Format(mStartDate, "YYYY-MM-DD")
            RsTxn_DepositContract!EndDate = Format(mEndDate, "YYYY-MM-DD")
            RsTxn_DepositContract!BillingCycleID = Val(.TextMatrix(mR, 17))
            RsTxn_DepositContract!DepBillingCycleID = Val(.TextMatrix(mR, 18))
            RsTxn_DepositContract!WBillingCycleID = Val(.TextMatrix(mR, 18))
            RsTxn_DepositContract!RatePer = "E"
            RsTxn_DepositContract!TentativeDur = 12
            RsTxn_DepositContract!LockPeriod = 1
            RsTxn_DepositContract!EmployeeID = 2789
            RsTxn_DepositContract!InsuranceByID = 2
            ''RsTxn_DepositContract!DepMonMTRate = ""
            RsTxn_DepositContract!BagSizeFlag = IIf(LCase(.TextMatrix(mR, 16)) = "b", "Y", "N")
            RsTxn_DepositContract!SM_Prefix = .TextMatrix(mR, 4)
            RsTxn_DepositContract!ApprovalFlag = "F"
            RsTxn_DepositContract!ApprovalRemarks = "System Generated Contract, status as forwarded to Head CMG with master define rate card"
            ''RsTxn_DepositContract!ApproveOn = ""
            ''RsTxn_DepositContract!ApproveBy = ""
            ''RsTxn_DepositContract!ClosedRemark = ""
            ''RsTxn_DepositContract!ClosedOn = ""
            ''RsTxn_DepositContract!ClosedBy = ""
            RsTxn_DepositContract!G_UID = GetGUID
            ''RsTxn_DepositContract!Concurrency
            RsTxn_DepositContract!CreatedDate = Now
            RsTxn_DepositContract!CreatedBy = Gen_UserLoginID
            ''RsTxn_DepositContract!UpdatedDate
            ''RsTxn_DepositContract!UpdatedBy
            RsTxn_DepositContract.Update
            RsTxn_DepositContract.Requery
        ''--------------------
            mLastPRow = mR
            mLastGRow = MSHFlexGrid2.Rows - 1
        
            If UCase(Trim(.TextMatrix(mR, 16))) = "Q" Then ''-- Q
            
                tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,"
                tmp = tmp & " MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
                tmp = tmp & " From Mst_Godown MG"
                tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
                tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
                tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
                tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
                tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & mEndDate & "' And ( '" & mStartDate & "' <= TLA.ExpiryDate And '" & mEndDate & "' >= TLA.StartDate )"
                tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1 And MG.GodownID In ( SELECT GodownID From Mst_GSL WHERE ExchangeTypeID = 1 And Flag not In ('E','Z','B') And (ContractID is Null or ContractID = 0) And BalanceBags > 0 And ClientIDRow = " & mClientID & " And SM_Prefix = '" & mSMPrefix & "' Group By GodownID,ClientIDRow ) "
                tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
                tmp = tmp & " Union All"
                tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
                tmp = tmp & " From Mst_Godown MG"
                tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
                tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
                tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & mEndDate & "' And '" & mEndDate & "' <= MG.AgreementEndDate "
                tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1 And MG.GodownID In ( SELECT GodownID From Mst_GSL WHERE ExchangeTypeID = 1 And Flag not In ('E','Z','B') And (ContractID is Null or ContractID = 0) And BalanceBags > 0 And ClientIDRow = " & mClientID & " And SM_Prefix = '" & mSMPrefix & "' Group By GodownID,ClientIDRow ) "
                tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
                Call TmpTable
            
                If TmpRs.RecordCount > 0 Then
                    
                    For mRow = 1 To TmpRs.RecordCount
                        mOwnGodownSqrFeetRate = GetOwnGodownPerSqtFeet(mLocationID)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 0) = .TextMatrix(mR, 21)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 1) = IIf(IsNull(TmpRs!CMPName), 0, TmpRs!CMPName)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 2) = IIf(IsNull(TmpRs!GodownNo), 0, TmpRs!GodownNo)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 3) = IIf(IsNull(TmpRs!Address), 0, TmpRs!Address)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 4) = IIf(IsNull(TmpRs!UnitType), 0, TmpRs!UnitType)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5) = Format(IIf(IsNull(TmpRs!Rent), 0, TmpRs!Rent), "#0.00")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6) = Format(IIf(IsNull(TmpRs!Area), 0, TmpRs!Area), "#0.000")
                        'MSHFlexGrid2.TextMatrix(mRow + mLastRow, 7) = Format(mMinBagSize, "#0.000")
                        'MSHFlexGrid2.TextMatrix(mRow + mLastRow, 8) = Format(mMaxBagSize, "#0.000")
                        
                        If TmpRs!UnitTypeID = 7 And Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5)) = 0 Then
                             MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6)) * mOwnGodownSqrFeetRate, "#0.00")
                        End If
        
                        If TmpRs!UnitTypeID = 1 Or TmpRs!UnitTypeID = 7 Then
                           If LCase(mFumigationFlag) = "y" Then
                              MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9) = Format((((Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5)) / Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                              MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13) = Format((((Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5)) / Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                              MSHFlexGrid2.TextMatrix(mRow + mLastRow, 18) = Format((((Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5)) / Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                           Else
                              MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9) = Format((((Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5)) / Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                              MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13) = Format((((Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5)) / Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                              MSHFlexGrid2.TextMatrix(mRow + mLastRow, 18) = Format((((Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 5)) / Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                           End If
                        Else
                           MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9) = "0.0000"
                           MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13) = "0.0000"
                           MSHFlexGrid2.TextMatrix(mRow + mLastRow, 18) = "0.0000"
                        End If
        
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 11) = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 12) = ""
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 15) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)), "#0.0000")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 16) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)), "#0.0000")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 17) = IIf(IsNull(TmpRs!UnitTypeID), 0, TmpRs!UnitTypeID)
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 19) = Format(IIf(Val(TmpRs!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs!ApprovalLimit), "#0.00")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 20) = UCase(Trim(.TextMatrix(mR, 16)))
                        mBillingOnePercentage = Format(IIf(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) > 0, Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) / 100, 0), "#0.0000")
                        mDepositOnePercentage = Format(IIf(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) > 0, Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) / 100, 0), "#0.0000")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 10) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) + (mBillingOnePercentage * MSHFlexGrid2.TextMatrix(mRow + mLastRow, 19)), "#0.0000")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 14) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) + (mDepositOnePercentage * MSHFlexGrid2.TextMatrix(mRow + mLastRow, 19)), "#0.0000")
                        
                        '''-- Billing Cycle / Deposite Biling Cycle Wise Rate Calculation
                                        
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9) = Format(CalculateRate(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 15), mBillingRatio, mBCycle), "#0.0000")
                        mBillingOnePercentage = Format(IIf(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) > 0, Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) / 100, 0), "#0.0000")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 10) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) + (mBillingOnePercentage * MSHFlexGrid2.TextMatrix(mRow + mLastRow, 19)), "#0.0000")
                    
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13) = Format(CalculateRate(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 16), mDepositBillingRatio, mDCycle), "#0.0000")
                        mDepositOnePercentage = Format(IIf(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) > 0, Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) / 100, 0), "#0.0000")
                        MSHFlexGrid2.TextMatrix(mRow + mLastRow, 14) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) + (mDepositOnePercentage * MSHFlexGrid2.TextMatrix(mRow + mLastRow, 19)), "#0.0000")
                           
                        ''-- Mapped the Rate
                        If UCase(.TextMatrix(mR, 19)) = "S" Then
                             TmpRs2.MoveFirst
                             TmpRs2.Find "SPRCID = " & Val(.TextMatrix(mR, 20))
                             If TmpRs2.EOF = False Then
                                 MSHFlexGrid2.TextMatrix(mRow + mLastRow, 10) = Format(IIf(IsNull(TmpRs2!Amount), 0, TmpRs2!Amount), "#0.0000")
                                 MSHFlexGrid2.TextMatrix(mRow + mLastRow, 14) = Format(IIf(IsNull(TmpRs2!Amount), 0, TmpRs2!Amount), "#0.0000")
                             End If
                        ElseIf UCase(.TextMatrix(mR, 19)) = "N" Then
                             TmpRs3.MoveFirst
                             TmpRs3.Find "SRCID = " & Val(.TextMatrix(mR, 20))
                             If TmpRs3.EOF = False Then
                                 MSHFlexGrid2.TextMatrix(mRow + mLastRow, 10) = Format(IIf(IsNull(TmpRs3!RateOnQty), 0, TmpRs3!RateOnQty), "#0.0000")
                                 MSHFlexGrid2.TextMatrix(mRow + mLastRow, 14) = Format(IIf(IsNull(TmpRs3!RateOnQty), 0, TmpRs3!RateOnQty), "#0.0000")
                             End If
                        End If
                        ''------------------
                        Call CheckRateByColor(mRow + mLastRow)   ''-- Color Effect
                           
                    TmpRs.MoveNext
                    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                    Next mRow
                         
                End If
            
            ElseIf UCase(Trim(.TextMatrix(mR, 16))) = "B" Then
            ''--
                tmp = " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,sum(M1.RentAmount) 'Rent',MG.Area,"
                tmp = tmp & " MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
                tmp = tmp & " From Mst_Godown MG"
                tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
                tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
                tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail M1 On MG.GodownID=M1.GodownID"
                tmp = tmp & " inner Join Txn_LessorAgreement TLA on M1.LAID = TLA.LAID And TLA.Flag = 'A'"
                tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & mEndDate & "' And ( '" & mStartDate & "' <= TLA.ExpiryDate And '" & mEndDate & "' >= TLA.StartDate )"
                tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID = 1 And MG.GodownID In ( SELECT GodownID From Mst_GSL WHERE ExchangeTypeID = 1 And Flag not In ('E','Z','B') And (ContractID is Null or ContractID = 0) And BalanceBags > 0 And ClientIDRow = " & mClientID & " And SM_Prefix = '" & mSMPrefix & "' Group By GodownID,ClientIDRow ) "
                tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
                tmp = tmp & " Union All"
                tmp = tmp & " Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,0 'Rent',MG.Area,MG.GodownID,cast('' As Varchar) DepConDID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0)ApprovalLimit "
                tmp = tmp & " From Mst_Godown MG"
                tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
                tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
                tmp = tmp & " Where MG.CloseDate is null  And MG.StartDate<='" & mEndDate & "' And '" & mEndDate & "' <= MG.AgreementEndDate "
                tmp = tmp & " And MC.LocationID = " & mLocationID & " And  MU.UnitTypeID <> 1 And MG.GodownID In ( SELECT GodownID From Mst_GSL WHERE ExchangeTypeID = 1 And Flag not In ('E','Z','B') And (ContractID is Null or ContractID = 0) And BalanceBags > 0 And ClientIDRow = " & mClientID & " And SM_Prefix = '" & mSMPrefix & "' Group By GodownID,ClientIDRow ) "
                tmp = tmp & " Group by MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,MU.UnitTypeID,Isnull(MG.ApprovalLimit,0) "
                Call TmpTable

                If TmpRs.RecordCount > 0 Then
                
                    For mRow = 1 To TmpRs.RecordCount
                        mOwnGodownSqrFeetRate = GetOwnGodownPerSqtFeet(mLocationID)
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 0) = mR
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 1) = IIf(IsNull(TmpRs!CMPName), 0, TmpRs!CMPName)
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 2) = IIf(IsNull(TmpRs!GodownNo), 0, TmpRs!GodownNo)
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 3) = IIf(IsNull(TmpRs!Address), 0, TmpRs!Address)
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 4) = IIf(IsNull(TmpRs!UnitType), 0, TmpRs!UnitType)
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5) = Format(IIf(IsNull(TmpRs!Rent), 0, TmpRs!Rent), "#0.00")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6) = Format(IIf(IsNull(TmpRs!Area), 0, TmpRs!Area), "#0.000")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 7) = Format(mMinBagSize, "#0.000")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 8) = Format(mMaxBagSize, "#0.000")
                        
                        If TmpRs!UnitTypeID = 7 And Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5)) = 0 Then
                             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5) = Format(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6)) * mOwnGodownSqrFeetRate, "#0.00")
                        End If
        
                        If TmpRs!UnitTypeID = 1 Or TmpRs!UnitTypeID = 7 Then
                           If LCase(mFumigationFlag) = "y" Then
                              MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9) = Format((((Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5)) / Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                              MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 13) = Format((((Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5)) / Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                              MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 18) = Format((((Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5)) / Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorY, "#0.0000")
                           Else
                              MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9) = Format((((Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5)) / Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                              MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 13) = Format((((Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5)) / Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                              MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 18) = Format((((Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 5)) / Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 6))) * Val(mCurrentConversionFactor) * Val(mBagEquivalent)) / mDivisionFactor) * 100 + mFumigationFactorN, "#0.0000")
                           End If
                        Else
                           MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9) = "0.0000"
                           MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 13) = "0.0000"
                           MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 18) = "0.0000"
                        End If
        
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 11) = IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 12) = ""
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 15) = Format(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9)), "#0.0000")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 16) = Format(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9)), "#0.0000")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 17) = IIf(IsNull(TmpRs!UnitTypeID), 0, TmpRs!UnitTypeID)
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 19) = Format(IIf(Val(TmpRs!ApprovalLimit) = 0, mCMPMtoAMRMApprovalLimit, TmpRs!ApprovalLimit), "#0.00")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 20) = UCase(Trim(.TextMatrix(mR, 16)))
                        mBillingOnePercentage = Format(IIf(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9)) > 0, Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9)) / 100, 0), "#0.0000")
                        mDepositOnePercentage = Format(IIf(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 13)) > 0, Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 13)) / 100, 0), "#0.0000")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 10) = Format(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 9)) + (mBillingOnePercentage * MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 19)), "#0.0000")
                        MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 14) = Format(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 13)) + (mDepositOnePercentage * MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 1, 19)), "#0.0000")
                        
                        Call AddBagRowAndUpdateBagSizeWithRate(MSHFlexGrid2.Rows - 1, UCase(.TextMatrix(mR, 19)), CDbl(.TextMatrix(mR, 20)))
                        
                        '''-- Billing Cycle / Deposite Biling Cycle Wise Rate Calculation
                        'MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9) = Format(CalculateRate(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 15), mBillingRatio, mBCycle) * Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 8)) / 1000, "#0.0000")
                        'mBillingOnePercentage = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) / 100, "#0.0000")
                        'MSHFlexGrid2.TextMatrix(mRow + mLastRow, 10) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 9)) + (mBillingOnePercentage * MSHFlexGrid2.TextMatrix(mRow + mLastRow, 19)), "#0.0000")
    
                        'MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13) = Format(CalculateRate(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 16), mDepositBillingRatio, mDCycle) * Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 8)) / 1000, "#0.0000")
                        'mDepositOnePercentage = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) / 100, "#0.0000")
                        'MSHFlexGrid2.TextMatrix(mRow + mLastRow, 14) = Format(Val(MSHFlexGrid2.TextMatrix(mRow + mLastRow, 13)) + (mDepositOnePercentage * MSHFlexGrid2.TextMatrix(mRow + mLastRow, 19)), "#0.0000")
    
                        'Call CheckBagRateByColor(mRow + mLastRow)    ''-- Color Effect
                        'mLastRow = MSHFlexGrid2.Rows - 2
                        
                    TmpRs.MoveNext
                    'MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                    Next mRow
                End If
            End If
            ''---- Data Insert in Secondry Table for godown details
            
            For i = mLastGRow To MSHFlexGrid2.Rows - 2
                 
                RsTxn_DepositContractGodownDetail.AddNew
                RsTxn_DepositContractGodownDetail!DepConDID = GetMaxID("DepConDID", "Txn_DepositContractGodownDetail")
                MSHFlexGrid2.TextMatrix(i, 12) = RsTxn_DepositContractGodownDetail!DepConDID
                RsTxn_DepositContractGodownDetail!DepConID = .TextMatrix(mR, 24)
                RsTxn_DepositContractGodownDetail!GodownID = MSHFlexGrid2.TextMatrix(i, 11)
                RsTxn_DepositContractGodownDetail!FinalPrice = MSHFlexGrid2.TextMatrix(i, 10)
                RsTxn_DepositContractGodownDetail!DepositFinalPrice = MSHFlexGrid2.TextMatrix(i, 14)
                RsTxn_DepositContractGodownDetail!WithdrawalFinalPrice = MSHFlexGrid2.TextMatrix(i, 14)
                RsTxn_DepositContractGodownDetail!BaseRate = MSHFlexGrid2.TextMatrix(i, 15)
                RsTxn_DepositContractGodownDetail!DepositBaseRate = MSHFlexGrid2.TextMatrix(i, 16)
                RsTxn_DepositContractGodownDetail!WithdrawalBaseRate = MSHFlexGrid2.TextMatrix(i, 16)
                RsTxn_DepositContractGodownDetail!OldBillingRate = MSHFlexGrid2.TextMatrix(i, 9)
                RsTxn_DepositContractGodownDetail!OldDepositBillingRate = MSHFlexGrid2.TextMatrix(i, 13)
                RsTxn_DepositContractGodownDetail!OldWithdrawalBillingRate = MSHFlexGrid2.TextMatrix(i, 13)
                If LCase(Trim(.TextMatrix(mR, 16))) = "b" Then
                    RsTxn_DepositContractGodownDetail!MinBagSize = MSHFlexGrid2.TextMatrix(i, 7)
                    RsTxn_DepositContractGodownDetail!MaxBagSize = MSHFlexGrid2.TextMatrix(i, 8)
                End If
                
                MSHFlexGrid2.row = i: MSHFlexGrid2.Col = 10
                If MSHFlexGrid2.CellBackColor = vbRed Then
                    RsTxn_DepositContractGodownDetail!BaseRateColor = "R"
                ElseIf MSHFlexGrid2.CellBackColor = vbGreen Then
                    RsTxn_DepositContractGodownDetail!BaseRateColor = "G"
                ElseIf MSHFlexGrid2.CellBackColor = vbYellow Then
                    RsTxn_DepositContractGodownDetail!BaseRateColor = "Y"
                Else
                    RsTxn_DepositContractGodownDetail!BaseRateColor = "W"
                End If
                MSHFlexGrid2.row = i: MSHFlexGrid2.Col = 14
                If MSHFlexGrid2.CellBackColor = vbRed Then
                    RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "R"
                    RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "R"
                ElseIf MSHFlexGrid2.CellBackColor = vbGreen Then
                    RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "G"
                    RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "G"
                ElseIf MSHFlexGrid2.CellBackColor = vbYellow Then
                    RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "Y"
                    RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "Y"
                Else
                    RsTxn_DepositContractGodownDetail!DepositBaseRateColor = "W"
                    RsTxn_DepositContractGodownDetail!WithdrawalBaseRateColor = "W"
                End If
                    
                RsTxn_DepositContractGodownDetail!SM_Prefix = mSMPrefix
                RsTxn_DepositContractGodownDetail!RatePerMonthPerMT = MSHFlexGrid2.TextMatrix(i, 18)
                RsTxn_DepositContractGodownDetail!Rent = MSHFlexGrid2.TextMatrix(i, 5)
                RsTxn_DepositContractGodownDetail!G_UID = GetGUID
                'RsTxn_DepositContractGodownDetail!Concurrency = ""
                RsTxn_DepositContractGodownDetail!CreatedBy = Gen_UserLoginID
                RsTxn_DepositContractGodownDetail!CreatedDate = Now
                ''RsTxn_DepositContractGodownDetail!UpdatedBy
                ''RsTxn_DepositContractGodownDetail!UpdatedDate
                RsTxn_DepositContractGodownDetail.Update
                RsTxn_DepositContractGodownDetail.Requery
            Next i
            ''------------------------------------------------------
        End If
        
      ProgressBar1.Value = mR
      ProgressBar1.Refresh
    Next mR
End With
MsgBox "Done."
End Sub

Private Sub AddBagRowAndUpdateBagSizeWithRate(Index_ As Double, Type_ As Variant, Id_ As Double)
Dim mR  As Double, mRow As Double, mC As Double
Dim TmpCell(20) As Variant
Dim mMinSize() As Double, mMaxSize() As Double, mBRate() As Double, mDRate() As Double

With MSHFlexGrid2
    ''--- Read The Current Cell Values in TmpCell array to refill in new Rows
    For mC = 0 To .Cols - 1
        TmpCell(mC) = .TextMatrix(Index_, mC)
    Next mC
    
    If UCase(Type_) = "S" Then
        tmp = " Select * from Mst_SpecialRateCardDetail Where SPRCID = " & Id_ & " Order By SPRCID,ToBagSize   "
        Call Tmp4Table
        tmp = " Select * from Mst_SpecialRateCardDepoDetail Where SPRCID = " & Id_ & " Order By SPRCID,ToBagSize  "
        Call Tmp5Table
    ElseIf UCase(Type_) = "N" Then
        tmp = " Select * from Mst_StorageRateCardDetail Where SRCID = " & Id_ & " Order By SRCID,ToBagSize   "
        Call Tmp4Table
        tmp = " Select * from Mst_StorageRateCardDepoDetail Where SRCID = " & Id_ & " Order By SRCID,ToBagSize   "
        Call Tmp5Table
    End If
    
    If TmpRs4.RecordCount > 0 And TmpRs5.RecordCount > 0 Then
        If TmpRs4.RecordCount = TmpRs5.RecordCount Then
            mRow = TmpRs4.RecordCount
        ElseIf TmpRs4.RecordCount > TmpRs5.RecordCount Then
            mRow = TmpRs4.RecordCount
        ElseIf TmpRs4.RecordCount < TmpRs5.RecordCount Then
            mRow = TmpRs5.RecordCount
        End If
    Else
        mRow = 1
    End If
   
    ReDim mMinSize(mRow) As Double
    ReDim mMaxSize(mRow) As Double
    ReDim mBRate(mRow) As Double
    ReDim mDRate(mRow) As Double
    
    If TmpRs4.RecordCount > 0 And TmpRs5.RecordCount > 0 Then
        TmpRs4.MoveFirst
        TmpRs5.MoveFirst
        For mR = 0 To mRow - 1
            If TmpRs4.RecordCount = TmpRs5.RecordCount Then
                mMinSize(mR) = IIf(IsNull(TmpRs4!FromBagSize), 0, TmpRs4!FromBagSize)
                mMaxSize(mR) = IIf(IsNull(TmpRs4!ToBagSize), 0, TmpRs4!ToBagSize)
                mBRate(mR) = IIf(IsNull(TmpRs4!RatePerBag), 0, TmpRs4!RatePerBag)
                mDRate(mR) = IIf(IsNull(TmpRs5!RatePerBag), 0, TmpRs5!RatePerBag)
                TmpRs4.MoveNext
                TmpRs5.MoveNext
            ElseIf TmpRs4.RecordCount > TmpRs5.RecordCount Then
                mMinSize(mR) = IIf(IsNull(TmpRs4!FromBagSize), 0, TmpRs4!FromBagSize)
                mMaxSize(mR) = IIf(IsNull(TmpRs4!ToBagSize), 0, TmpRs4!ToBagSize)
                mBRate(mR) = IIf(IsNull(TmpRs4!RatePerBag), 0, TmpRs4!RatePerBag)
                mDRate(mR) = IIf(IsNull(TmpRs5!RatePerBag), 0, TmpRs5!RatePerBag)
                If mR < mRow Then TmpRs4.MoveNext
                If mR < mRow Then TmpRs5.MoveNext
            ElseIf TmpRs4.RecordCount < TmpRs5.RecordCount Then
                mMinSize(mR) = IIf(IsNull(TmpRs5!FromBagSize), 0, TmpRs5!FromBagSize)
                mMaxSize(mR) = IIf(IsNull(TmpRs5!ToBagSize), 0, TmpRs5!ToBagSize)
                mBRate(mR) = IIf(IsNull(TmpRs5!RatePerBag), 0, TmpRs5!RatePerBag)
                mDRate(mR) = IIf(IsNull(TmpRs4!RatePerBag), 0, TmpRs4!RatePerBag)
                If mR < mRow Then TmpRs4.MoveNext
                If mR < mRow Then TmpRs5.MoveNext
            End If
        Next mR
    End If
    
    For mR = Index_ To Index_ + IIf(mRow = 1, 0, mRow - 1)
        
        .TextMatrix(mR, 0) = TmpCell(0)
        .TextMatrix(mR, 1) = TmpCell(1)
        .TextMatrix(mR, 2) = TmpCell(2)
        .TextMatrix(mR, 3) = TmpCell(3)
        .TextMatrix(mR, 4) = TmpCell(4)
        .TextMatrix(mR, 5) = TmpCell(5)
        .TextMatrix(mR, 6) = TmpCell(6)
        .TextMatrix(mR, 7) = mMinSize(mR - Index_)  ''-- TmpCell(7) ''-- Min Bag
        .TextMatrix(mR, 8) = mMaxSize(mR - Index_)  ''-- TmpCell(8) ''-- Max Bag
        .TextMatrix(mR, 9) = TmpCell(9)
        .TextMatrix(mR, 10) = TmpCell(10) ''-- Billing Rate
        .TextMatrix(mR, 11) = TmpCell(11)
        .TextMatrix(mR, 12) = TmpCell(12)
        .TextMatrix(mR, 13) = TmpCell(13)
        .TextMatrix(mR, 14) = TmpCell(14) ''-- Deposite Rate
        .TextMatrix(mR, 15) = TmpCell(15)
        .TextMatrix(mR, 16) = TmpCell(16)
        .TextMatrix(mR, 17) = TmpCell(17)
        .TextMatrix(mR, 18) = TmpCell(18)
        .TextMatrix(mR, 19) = TmpCell(19)
        .TextMatrix(mR, 20) = TmpCell(20)
        
        '''-- Billing Cycle / Deposite Biling Cycle Wise Rate Calculation
        .TextMatrix(mR, 9) = Format(CalculateRate(.TextMatrix(mR, 15), mBillingRatio, mBCycle) * Val(.TextMatrix(mR, 8)) / 1000, "#0.0000")
        mBillingOnePercentage = Format(Val(.TextMatrix(mR, 9)) / 100, "#0.0000")
        .TextMatrix(mR, 10) = Format(Val(.TextMatrix(mR, 9)) + (mBillingOnePercentage * .TextMatrix(mR, 19)), "#0.0000")

        .TextMatrix(mR, 13) = Format(CalculateRate(.TextMatrix(mR, 16), mDepositBillingRatio, mDCycle) * Val(.TextMatrix(mR, 8)) / 1000, "#0.0000")
        mDepositOnePercentage = Format(Val(.TextMatrix(mR, 13)) / 100, "#0.0000")
        .TextMatrix(mR, 14) = Format(Val(.TextMatrix(mR, 13)) + (mDepositOnePercentage * .TextMatrix(mR, 19)), "#0.0000")
        
        .TextMatrix(mR, 10) = Format(mBRate(mR - Index_), "#0.0000")   ''-- Billing Rate
        .TextMatrix(mR, 14) = Format(mDRate(mR - Index_), "#0.0000")   ''-- Deposite Rate

        Call CheckBagRateByColor(CInt(mR))    ''-- Color Effect
        
        .Rows = .Rows + 1
    Next mR
        
End With

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call Grid_Head2

TxtStartDate.Value = Date
TxtEndDate.Value = Date + 30
CmdProcess.Enabled = False

''-- Calculation Factor Values From Email Config File
mCurrentConversionFactor = Val(Gen_DepositContractCalculationFactor("DepositContractCurrentConversionFactor-"))
mDivisionFactor = Val(Gen_DepositContractCalculationFactor("DepositContractDivisionFactor-"))                        ''-- Percentage %
mFumigationFactorY = Val(Gen_DepositContractCalculationFactor("DepositContractFumigationFactorY-"))
mFumigationFactorN = Val(Gen_DepositContractCalculationFactor("DepositContractFumigationFactorN-"))
mCMPMtoAMRMApprovalLimit = Val(Gen_DepositContractCalculationFactor("DepositContractCMPMtoAMRMApprovalLimit-"))      ''-- Percentage %
''-- if Godown Type is Non Leased then (Rent is not available) then PAN India Mini Value From Email Config File
mKnownRentMinimumValue = Val(Gen_DepositContractCalculationFactor("DepositContractNonKnownRentMinimumValue-"))
''-- CMG Group ID to be List in Forward to Additionlly and Approved for Same
mDepositContractCMGGroupToBeApproved = Gen_DepositContractCalculationFactor("DepositContractCMGGroupToBeApproved-")
mDepositContractDefaultDaysRange = Gen_DepositContractCalculationFactor("DepositContractDefaultDaysRange-")

End Sub

Public Sub Grid_Head2()
With MSHFlexGrid2
     .Font.Size = 9
     .Clear
     .Rows = 2
     .Cols = 21
     .FixedCols = 0
     
    .TextMatrix(0, 0) = "Select":                                              .ColWidth(0) = 500
    .TextMatrix(0, 1) = "CMP Name":                                            .ColWidth(1) = 1300
    .TextMatrix(0, 2) = "Godown No":                                           .ColWidth(2) = 900
    .TextMatrix(0, 3) = "Address":                                             .ColWidth(3) = 1200
    .TextMatrix(0, 4) = "Godown Type":                                         .ColWidth(4) = 1000
    .TextMatrix(0, 5) = "Rent":                                                .ColWidth(5) = 1000
    .TextMatrix(0, 6) = "Area in Sq. Ft.":                                     .ColWidth(6) = 1000
    .TextMatrix(0, 7) = "Bag Size (Minimum)":                                  .ColWidth(7) = 900
    .TextMatrix(0, 8) = "Bag Size (Maximum)":                                  .ColWidth(8) = 900
    .TextMatrix(0, 9) = "Billing Rate/MT (As per System)":                    .ColWidth(9) = 1000
    .TextMatrix(0, 10) = "Billing Rate/MT (Proposed/ Approved)":              .ColWidth(10) = 1200
    .TextMatrix(0, 11) = "GodownID":                                           .ColWidth(11) = 1000
    .TextMatrix(0, 12) = "ROWID":                                              .ColWidth(12) = 1000
    .TextMatrix(0, 13) = "Deposit/ Withdrawal Rate/MT (As per System)":        .ColWidth(13) = 1200
    .TextMatrix(0, 14) = "Deposit/ Withdrawal Rate/MT (Proposed/ Approved)":   .ColWidth(14) = 1200
    .TextMatrix(0, 15) = "Billing Rate":                                       .ColWidth(15) = 1000
    .TextMatrix(0, 16) = "Deposit Billing Rate":                               .ColWidth(16) = 1000
    .TextMatrix(0, 17) = "UnitTypeID":                                         .ColWidth(17) = 1000
    .TextMatrix(0, 18) = "Billing Rate/MT (Monthly Cycle)":                   .ColWidth(18) = 1200
    .TextMatrix(0, 19) = "ApprovalLimit":                                      .ColWidth(19) = 1000
    .TextMatrix(0, 20) = "BagFlag":                                            .ColWidth(20) = 1000
    .RowHeight(0) = 850
End With
End Sub

Public Function GetOwnGodownPerSqtFeet(mLocation_ As Variant) As Double
tmp = " SELECT LNE.LocationID,LNE.NotionalExp,Sum(MG.Area)[TotalArea],LNE.NotionalExp/Sum(MG.Area)[SqrFeetRate] FROM Txn_LocationwiseNotionalExpense LNE"
tmp = tmp & " Inner Join Mst_CMP MC On LNE.LocationID = MC.LocationID"
tmp = tmp & " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID"
tmp = tmp & " Where LNE.LocationID = " & Val(mLocation_) & " Group By LNE.LocationID,LNE.NotionalExp"
Call Tmp1Table
If TmpRs1.EOF = False Then
    GetOwnGodownPerSqtFeet = Format(IIf(IsNull(TmpRs1!SqrFeetRate), "0.00", TmpRs1!SqrFeetRate), "#0.0000")
Else
    GetOwnGodownPerSqtFeet = 0
End If
End Function

Public Function CalculateRate(mmRateOnQty, mmRatio, mmBillingCycle) As Double
Dim Retval As Variant

If LCase(mmBillingCycle) = "daily" Then
   Retval = "30"   'According to Mayur sir, it will be 30 days
ElseIf LCase(mmBillingCycle) = "weekly" Then
   Retval = "4"
ElseIf LCase(mmBillingCycle) = "fortnightly" Then
   Retval = "2"
ElseIf LCase(mmBillingCycle) = "monthly" Then
   Retval = "1"
End If

CalculateRate = Format((Val(mmRateOnQty) / Val(Retval)) * Val(mmRatio) / 100, "########0.0000")
CalculateRate = Format((Val(mmRateOnQty) / Val(Retval)) + Val(CalculateRate), "########0.0000")
End Function


Sub CheckRateByColor(Index_ As Integer)
Dim mBillingOnePercentage, mDepositOnePercentage As Variant
Dim Retval, mGetResult, mBResult, mDResult As Variant
Dim mColor As ColorConstants
Retval = ""
With MSHFlexGrid2
    If .TextMatrix(Index_, 1) <> "" And .TextMatrix(Index_, 0) <> "" Then
        mBillingOnePercentage = Format(Val(.TextMatrix(Index_, 9)) / 100, "#0.0000")
        mDepositOnePercentage = Format(Val(.TextMatrix(Index_, 13)) / 100, "#0.0000")
        ''-- Lease / Own Godown
        If .TextMatrix(Index_, 17) = 1 Or .TextMatrix(Index_, 17) = 7 Then
           
           If Val(.TextMatrix(Index_, 10)) < Val(.TextMatrix(Index_, 9)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 10: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 10)) >= Val(.TextMatrix(Index_, 9)) And Val(.TextMatrix(Index_, 10)) <= Format(Val(.TextMatrix(Index_, 9)) + (mBillingOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 10: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 10)) > Format(Val(.TextMatrix(Index_, 9)) - (mBillingOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 10: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 10: .row = Index_: .CellBackColor = mColor
           End If
           
           If Val(.TextMatrix(Index_, 14)) < Val(.TextMatrix(Index_, 13)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 14: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 14)) >= Val(.TextMatrix(Index_, 13)) And Val(.TextMatrix(Index_, 14)) <= Format(Val(.TextMatrix(Index_, 13)) + (mDepositOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 14: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 14)) > Format(Val(.TextMatrix(Index_, 13)) - (mDepositOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 14: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 14: .row = Index_: .CellBackColor = mColor
           End If
           
        Else ''--Un-Lease Godown
           If .TextMatrix(Index_, 17) = 2 Then ''-- Franchisee
               mGetResult = GetFranchiseeRate(mLocationID, Val(.TextMatrix(Index_, 11)), mCommodityID, Val(.TextMatrix(Index_, 10)), Val(.TextMatrix(Index_, 14)), CDbl(.TextMatrix(Index_, 19)))
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 10)) > 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 10: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 10: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 14)) > 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 14: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 14: .row = Index_: .CellBackColor = mColor
               End If

           Else
               mGetResult = "CC"      ''-- for Any rate Approvel from CMG
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 10)) > 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 10: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 10: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 14)) > 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 14: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 14: .row = Index_: .CellBackColor = mColor
               End If
               
           End If

        End If
    End If
End With
End Sub

Sub CheckBagRateByColor(Index_ As Integer)
Dim mBillingOnePercentage, mDepositOnePercentage As Variant
Dim Retval, mGetResult, mBResult, mDResult As Variant
Dim mColor As ColorConstants
Retval = ""
With MSHFlexGrid2
    If .TextMatrix(Index_, 1) <> "" And .TextMatrix(Index_, 0) <> "" Then
        mBillingOnePercentage = Format(Val(.TextMatrix(Index_, 9)) / 100, "#0.0000")
        mDepositOnePercentage = Format(Val(.TextMatrix(Index_, 13)) / 100, "#0.0000")
        
        ''-- Lease / Own Godown
        If .TextMatrix(Index_, 17) = 1 Or .TextMatrix(Index_, 17) = 7 Then
        
           If Val(.TextMatrix(Index_, 10)) < Val(.TextMatrix(Index_, 9)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 10: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 10)) >= Val(.TextMatrix(Index_, 9)) And Val(.TextMatrix(Index_, 10)) <= Format(Val(.TextMatrix(Index_, 9)) + (mBillingOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 10: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 10)) > Format(Val(.TextMatrix(Index_, 9)) - (mBillingOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 10: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 10: .row = Index_: .CellBackColor = mColor
           End If
           
           If Val(.TextMatrix(Index_, 14)) < Val(.TextMatrix(Index_, 13)) Then
           ''-- Rate is less then System rate (CMG Approval)
              mColor = vbRed
              .Col = 14: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 14)) >= Val(.TextMatrix(Index_, 13)) And Val(.TextMatrix(Index_, 14)) <= Format(Val(.TextMatrix(Index_, 13)) + (mDepositOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is less then System rate (RM / Upper Approval)
              mColor = vbYellow
              .Col = 14: .row = Index_: .CellBackColor = mColor
           ElseIf Val(.TextMatrix(Index_, 14)) > Format(Val(.TextMatrix(Index_, 13)) - (mDepositOnePercentage * Val(.TextMatrix(Index_, 19))), "#0.0000") Then
           ''-- Rate is more then System rate + Approval Limit (System Approval)
              mColor = vbGreen
              .Col = 14: .row = Index_: .CellBackColor = mColor
           Else
              mColor = vbWhite
              .Col = 14: .row = Index_: .CellBackColor = mColor
           End If
           
        Else ''--Un-Lease Godown
           If .TextMatrix(Index_, 17) = 2 Then ''-- Franchisee
               mGetResult = GetFranchiseeRate(mLocationID, Val(.TextMatrix(Index_, 11)), mCommodityID, Val(.TextMatrix(Index_, 10)), Val(.TextMatrix(Index_, 14)), CDbl(.TextMatrix(Index_, 19)))
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 10)) > 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 10: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 10: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 14)) > 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 14: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 14: .row = Index_: .CellBackColor = mColor
               End If

           Else
               mGetResult = "CC"      ''-- for Any rate Approvel from CMG
               mBResult = Left(mGetResult, 1)
               mDResult = Right(mGetResult, 1)
               
               If Val(.TextMatrix(Index_, 10)) > 0 Then
                    If LCase(mBResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mBResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 10: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 10: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 10: .row = Index_: .CellBackColor = mColor
               End If
               
               If Val(.TextMatrix(Index_, 14)) > 0 Then
                    If LCase(mDResult) = LCase("C") Then
                       mColor = vbRed
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("U") Then
                       mColor = vbYellow
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    ElseIf LCase(mDResult) = LCase("S") Then
                       mColor = vbGreen
                       .Col = 14: .row = Index_: .CellBackColor = mColor
                    Else
                        mColor = vbWhite
                        .Col = 14: .row = Index_: .CellBackColor = mColor
                    End If
               Else
                    mColor = vbWhite
                    .Col = 14: .row = Index_: .CellBackColor = mColor
               End If
               
           End If

        End If
    End If
End With
End Sub

Public Function GetFranchiseeRate(mLocation_ As Variant, mGodownID_ As Variant, mCommodityID_ As Variant, mmRate As Double, mmDRate As Double, mApprovalLimit_ As Double) As String
Dim mGFRCycleID, mGFRRate, mGFRUnit, mGFRRatio, mGFRCycle As Variant
Dim mGFRBRatePerMonth, mGFRDRatePerMonth, mIncomePart, mDepositIncomePart, mGFRBRatio, mGFRDRatio As Variant
Dim mReturn, mDReturn, mBKnownRentMinimumValue, mDKnownRentMinimumValue, mApprovalLimit As Variant
Dim mBOnePercentage, mDOnePercentage, mBRate, mDRate As Variant
Dim mGodownMinimumRent As Variant

''-- Get Godown wise minimum charges monthly otherwise from config
tmp = " Select RentMinimumValue From Mst_Godown Where CloseDate Is Null And GodownID = " & mGodownID_ & " "
Call Tmp1Table
If Not TmpRs1.EOF Then
    mGodownMinimumRent = IIf(IsNull(TmpRs1!RentMinimumValue), mKnownRentMinimumValue, RentMinimumValue)
Else
    mGodownMinimumRent = mKnownRentMinimumValue
End If

''-- Billing Cycle Wise Minimum Godown Rent Value
tmp = " Select * From Mst_BillingCycle Where Flag<>'R' And BillingCycleID = " & mBillingCycleID & " "
Call Tmp1Table
If Not TmpRs1.EOF Then
    mGFRBRatio = IIf(IsNull(TmpRs1!Ratio), 0, TmpRs1!Ratio)
    mBKnownRentMinimumValue = CalculateRate(mGodownMinimumRent, mGFRBRatio, mBCycle)
Else
    mGFRBRatio = 0
    mBKnownRentMinimumValue = mGodownMinimumRent
End If
''-- Deposite Billing Cycle Wise Minimum Godown Rent Value
tmp = " Select * From Mst_BillingCycle Where Flag<>'R' And BillingCycleID = " & mDepositBillingCycleID & " "
Call Tmp1Table
If Not TmpRs1.EOF Then
    mGFRDRatio = IIf(IsNull(TmpRs1!Ratio), 0, TmpRs1!Ratio)
    mDKnownRentMinimumValue = CalculateRate(mGodownMinimumRent, mGFRDRatio, mDCycle)
Else
    mGFRDRatio = 0
    mDKnownRentMinimumValue = mGodownMinimumRent
End If

mBOnePercentage = mmRate / 100
mDOnePercentage = mmDRate / 100

tmp = " SELECT FFRCG.BillingCycleID,Max(FFRCG.RateOnQty) [Rate],FFRCG.BillingUnit,MBC.Ratio,MBC.BillingCycle"
tmp = tmp & " FROM Mst_FF_RateCard_Godown FFRCG"
tmp = tmp & " Inner Join Mst_FF_RateCard FFRC On FFRCG.FF_RateCardID = FFRC.FF_RateCardID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On FFRC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_Commodity MC On MCG.CommodityGroupID = MC.CommodityGroupID"
tmp = tmp & " Inner Join Mst_BillingCycle MBC On FFRCG.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Where FFRC.Flag = 'A' And FFRC.ExchangeTypeID = 1"
tmp = tmp & " And FFRC.LocationID = " & mLocation_ & " And FFRCG.GodownID = " & mGodownID_ & " And MC.CommodityID = " & mCommodityID_ & ""
tmp = tmp & " Group By FFRCG.GodownID,FFRCG.BillingCycleID,FFRCG.BillingUnit,MBC.Ratio,MBC.BillingCycle"
Call Tmp1Table
If TmpRs1.EOF = False Then
    
    mGFRCycleID = IIf(IsNull(TmpRs1!BillingCycleID), "", TmpRs1!BillingCycleID)
    mGFRRate = Format(IIf(IsNull(TmpRs1!Rate), 0#, TmpRs1!Rate), "#0.0000")
    mGFRUnit = IIf(IsNull(TmpRs1!BillingUnit), "", TmpRs1!BillingUnit)
    mGFRRatio = Format(IIf(IsNull(TmpRs1!Ratio), 0#, TmpRs1!Ratio), "#0.0000")
    mGFRCycle = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
    
    Select Case LCase(mGFRUnit)
        Case "p"
            mIncomePart = mmRate - ((mmRate * mGFRRate) / 100)
            mDepositIncomePart = mmDRate - ((mmDRate * mGFRRate) / 100)
            
            ''-- Rate is less then System rate (CMG Approval)
            If mIncomePart < mBKnownRentMinimumValue Then
                mReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mIncomePart >= mBKnownRentMinimumValue) And (mIncomePart < mBKnownRentMinimumValue + (mBOnePercentage * mApprovalLimit_)) Then
                mReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mIncomePart > (mBKnownRentMinimumValue + (mBOnePercentage * mApprovalLimit_)) Then
               mReturn = "S"
            End If
               
            ''-- Rate is less then System rate (CMG Approval)
            If mDepositIncomePart < mDKnownRentMinimumValue Then
                mDReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mDepositIncomePart >= mDKnownRentMinimumValue) And (mDepositIncomePart < mDKnownRentMinimumValue + (mDOnePercentage * mApprovalLimit_)) Then
                mDReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mDepositIncomePart > (mBKnownRentMinimumValue + (mDOnePercentage * mApprovalLimit_)) Then
               mDReturn = "S"
            End If
            
        Case "b"
            mReturn = "C"
            mDReturn = "C"
        Case "f"
            mReturn = "C"
            mDReturn = "C"
        Case "q"
            If mGFRCycleID = 4 Then
                mGFRBRatePerMonth = mGFRRate
            ElseIf mGFRCycleID = 3 Then
                mGFRBRatePerMonth = mGFRRate * 2
            ElseIf mGFRCycleID = 2 Then
                mGFRBRatePerMonth = mGFRRate * 4
            ElseIf mGFRCycleID = 1 Then
                mGFRBRatePerMonth = mGFRRate * 30
            End If
            mGFRDRatePerMonth = mGFRBRatePerMonth
            
            mBRate = CalculateRate(mGFRBRatePerMonth, mGFRBRatio, mBCycle)
            mDRate = CalculateRate(mGFRDRatePerMonth, mGFRDRatio, mDCycle)
        
            ''-- Rate is less then System rate (CMG Approval)
            If mmRate < mBRate Then
                mReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mmRate >= mBRate) And (mmRate < mBRate + (mBOnePercentage * mApprovalLimit_)) Then
                mReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mmRate > (mBRate + (mBOnePercentage * mApprovalLimit_)) Then
               mReturn = "S"
            End If
            
            ''-- Rate is less then System rate (CMG Approval)
            If mmDRate < mDRate Then
                mDReturn = "C"
            ''-- Rate is less then System rate (RM / Upper Approval)
            ElseIf (mmDRate >= mDRate) And (mmDRate < mDRate + (mBOnePercentage * mApprovalLimit_)) Then
                mDReturn = "U"
            ''-- Rate is more then System rate + Approval Limit (System Approval)
            ElseIf mmDRate > (mDRate + (mBOnePercentage * mApprovalLimit_)) Then
               mDReturn = "S"
            End If
            
            'mReturn = "C"
            'mDReturn = "C"
    End Select
Else    ''-- Rate for Found
    mReturn = "C"
    mDReturn = "C"
End If
GetFranchiseeRate = mReturn & "" & mDReturn
End Function


