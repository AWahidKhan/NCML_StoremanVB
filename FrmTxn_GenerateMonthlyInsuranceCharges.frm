VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_GenerateMonthlyInsuranceCharges 
   Caption         =   "Generate Monthly Insurance Charges"
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
      Height          =   660
      Left            =   120
      TabIndex        =   14
      Top             =   8805
      Width           =   15090
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   75
         TabIndex        =   15
         Top             =   210
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8820
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15090
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   6945
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   12250
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Left            =   10095
         TabIndex        =   17
         Top             =   8100
         Width           =   1440
      End
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Delete processed data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   765
         Width           =   2925
      End
      Begin VB.ComboBox CmbMonth 
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
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton CmdShowLocation 
         Caption         =   "Show Locations"
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
         Left            =   6360
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox CmbPreFix 
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
         TabIndex        =   4
         Top             =   8760
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox CmbYear 
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
         Left            =   2595
         TabIndex        =   3
         Top             =   375
         Width           =   1620
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   11535
         TabIndex        =   2
         Top             =   8100
         Width           =   1440
      End
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Process"
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
         Left            =   4500
         TabIndex        =   1
         Top             =   368
         Width           =   1440
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6945
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   12250
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         Left            =   6135
         TabIndex        =   13
         Top             =   780
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label7 
         Caption         =   "Total Location Found :"
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
         Left            =   3735
         TabIndex        =   12
         Top             =   780
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Month"
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
         Left            =   1020
         TabIndex        =   11
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label36 
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
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   8400
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
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
         Left            =   3090
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmTxn_GenerateMonthlyInsuranceCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mTotalDays, mToDate, mNoofdays As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mGenInvoiceLocation As Variant
Dim mCountLocation As Variant

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdProcess_Click()
ProgressBar1.Value = 0


Call Grid_Head1

If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

'If MSHFlexGrid1.Rows < 3 Then
'   MsgBox " No Location populated in Grid  !!!"
'   Exit Sub
'End If
'
'If ChkSelection = False Then
'   MsgBox " No Location selected from Grid  !!!"
'   Exit Sub
'End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
mNoofdays = ""

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If



'tmp = "Select * from Txn_MonthlyInsuranceCharges Where YearMonth = '" & mFromDate & "'"
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   MsgBox "Insurance Charges already exist for selected Month!!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

Dim ans As Variant

'CmdProcess.Enabled = False
'mGenInvoiceLocation = ""
'
'If MSHFlexGrid1.Rows > 2 Then
'   mGenInvoiceLocation = "SM_Prefix in ("
'   For i = 1 To MSHFlexGrid1.Rows - 1
'       If MSHFlexGrid1.TextMatrix(i, 1) = "" Then Exit For
'          If MSHFlexGrid1.TextMatrix(i, 0) = strChecked Then
'             If mGenInvoiceLocation = "SM_Prefix in (" Then
'                mGenInvoiceLocation = mGenInvoiceLocation & "'" & MSHFlexGrid1.TextMatrix(i, 2) & "'"
'             Else
'                mGenInvoiceLocation = mGenInvoiceLocation & ",'" & MSHFlexGrid1.TextMatrix(i, 2) & "' "
'             End If
'          End If
'      Next
'      mGenInvoiceLocation = mGenInvoiceLocation & " )"
'End If

'tmp = "Select Distinct SM_Prefix, LocationID From Mst_Location Where " & mGenInvoiceLocation


tmp = " Select LocationID,SM_Prefix,GodownID,CommodityID,Sum(Bags) 'Bags',Sum(Wt) 'Wt','','','' From"
tmp = tmp & " ((Select ML.LocationID,ML.SM_Prefix,TIM.GodownID,TIM.CommodityID,Sum(BalanceBags) 'Bags' ,Sum(BalanceWeight)'Wt' From Txn_InvMonGSLDetail TIM"
tmp = tmp & " Inner Join Mst_Location ML On TIM.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Where TIM.YearMonth = '" & mFromDate & "' And TIM.TFlag in ('O','D') "
tmp = tmp & " Group By ML.LocationID,ML.SM_Prefix,TIM.CommodityID,TIM.GodownID)"
tmp = tmp & " Union All"
tmp = tmp & " (Select LocationID,SM_Prefix,GodownID,CommodityID,-Sum(Bags) 'Bags',-Sum(Wt) 'Wt' From"
tmp = tmp & " ((Select ML.LocationID,ML.SM_Prefix,TIM.GodownID,TIM.CommodityID,Sum(BalanceBags) 'Bags' ,Sum(BalanceWeight)'Wt' From Txn_InvMonGSLDetail TIM"
tmp = tmp & " Inner Join Mst_Location ML On TIM.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Txn_InsuranceExceptionMapping  TIEM On ML.LocationID = TIEM.LocationID And TIM.ClientIDRow = TIEM.ClientIDRow"
tmp = tmp & " Where TIM.YearMonth = '" & mFromDate & "' And TIM.TFlag in ('O','D') And TIEM.ClientIDRow is not null And '" & mFromDate & "'  Between TIEM.StartDate And TIEM.EndDate "
tmp = tmp & " Group By ML.LocationID,ML.SM_Prefix,TIM.CommodityID,TIM.GodownID)"
tmp = tmp & " Union All"
tmp = tmp & " (Select ML.LocationID,ML.SM_Prefix,TIM.GodownID,TIM.CommodityID,Sum(BalanceBags) 'Bags' ,Sum(BalanceWeight)'Wt' From Txn_InvMonGSLDetail TIM"
tmp = tmp & " Inner Join Mst_Location ML On TIM.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Txn_InsuranceExceptionMapping  TIEM On ML.LocationID = TIEM.LocationID And TIM.CommodityID = TIEM.CommodityID"
tmp = tmp & " Where TIM.YearMonth = '" & mFromDate & "' And TIM.TFlag in ('O','D') And '" & mFromDate & "'  Between TIEM.StartDate And TIEM.EndDate And TIEM.CommodityID is not null "
tmp = tmp & " And Convert(Varchar,ML.LocationID)+'-'+Convert(Varchar,TIM.ClientIDRow) Not in"
tmp = tmp & " (Select Convert(Varchar,LocationID)+'-'+Convert(Varchar,ClientIDRow) From Txn_InsuranceExceptionMapping Where ClientIDRow is not null And '" & mFromDate & "' Between TIEM.StartDate And TIEM.EndDate)"
tmp = tmp & " Group By ML.LocationID,ML.SM_Prefix,TIM.CommodityID,TIM.GodownID))  A"
tmp = tmp & " Group By LocationID,SM_Prefix,GodownID,CommodityID)) B"
tmp = tmp & " Group By LocationID,SM_Prefix,GodownID,CommodityID"
tmp = tmp & " Order By SM_Prefix,GodownID,CommodityID"

Call TmpTable

tmp = " Select Convert(Varchar,ML.LocationID)+'-'+Convert(Varchar,MCP.CommodityID) 'Key',ML.LocationID,MCP.CommodityID,MCP.YearMonth,Day28,Day29,Day30,Day31 From Mst_CommodityPrice MCP"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
tmp = tmp & " Where YearMonth = (Select Max(YearMonth) From Mst_CommodityPrice MCP"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
tmp = tmp & " Where YearMonth <= '" & mFromDate & "')"

Call Tmp1Table

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 5
MSHFlexGrid2.Rows = TmpRs.RecordCount + 3

x = 2
For i = 1 To TmpRs.RecordCount
    For C = 0 To MSHFlexGrid2.Cols - 1
        MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    
    If TmpRs1.RecordCount > 0 Then
       TmpRs1.MoveFirst
       'TmpRs1.Find "" & TmpRs1!LocationID & "-" & TmpRs1!CommodityID & " =  " & TmpRs!LocationID & "-" & TmpRs!CommodityID & " "
       TmpRs1.Find "Key =  '" & TmpRs!LocationID & "-" & TmpRs!CommodityID & "'"
       If TmpRs1.EOF = False Then
          If IsNull(TmpRs1!Day31) = False Then
             MSHFlexGrid2.TextMatrix(x, 6) = ((IIf(IsNull(TmpRs1!Day31), 0, Val(TmpRs1!Day31)))) ' * TmpRs1!Day31)
          ElseIf IsNull(TmpRs1!Day30) = False Then
             MSHFlexGrid2.TextMatrix(x, 6) = ((IIf(IsNull(TmpRs1!Day30), 0, Val(TmpRs1!Day30)))) ' ((IIf(IsNull(MSHFlexGrid2.TextMatrix(x, 5)), 0, Val(MSHFlexGrid2.TextMatrix(x, 5)))) * TmpRs1!Day30)
          ElseIf IsNull(TmpRs1!Day29) = False Then
             MSHFlexGrid2.TextMatrix(x, 6) = ((IIf(IsNull(TmpRs1!Day29), 0, Val(TmpRs1!Day29)))) '((IIf(IsNull(MSHFlexGrid2.TextMatrix(x, 5)), 0, Val(MSHFlexGrid2.TextMatrix(x, 5)))) * TmpRs1!Day29)
          ElseIf IsNull(TmpRs1!Day28) = False Then
             MSHFlexGrid2.TextMatrix(x, 6) = ((IIf(IsNull(TmpRs1!Day28), 0, Val(TmpRs1!Day28)))) '((IIf(IsNull(MSHFlexGrid2.TextMatrix(x, 5)), 0, Val(MSHFlexGrid2.TextMatrix(x, 5)))) * TmpRs1!Day28)
          End If
       Else
          MSHFlexGrid2.TextMatrix(x, 6) = "0"
       End If
       MSHFlexGrid2.TextMatrix(x, 7) = Val(MSHFlexGrid2.TextMatrix(x, 5)) * Val(MSHFlexGrid2.TextMatrix(x, 6))
       MSHFlexGrid2.TextMatrix(x, 7) = Format(MSHFlexGrid2.TextMatrix(x, 7), "#######0.00")
       If Format(mFromDate, "yyyy-mm") >= "2012-04" And Format(mFromDate, "yyyy-mm") <= "2013-03" Then
          MSHFlexGrid2.TextMatrix(x, 8) = Val(MSHFlexGrid2.TextMatrix(x, 7)) * 0.0003 / 12
       Else
          If Val(MSHFlexGrid2.TextMatrix(x, 3)) <> "43" Then
             MSHFlexGrid2.TextMatrix(x, 8) = Val(MSHFlexGrid2.TextMatrix(x, 7)) * 0.0006 / 12
          Else
             MSHFlexGrid2.TextMatrix(x, 8) = Val(MSHFlexGrid2.TextMatrix(x, 7)) * 0.005 / 12
          End If
       End If
       If Val(MSHFlexGrid2.TextMatrix(x, 7)) <= 0 Then
          MSHFlexGrid2.TextMatrix(x, 7) = 0
       End If
       
       'If Bags are zero don't calculate Insurance Charges
       If Val(MSHFlexGrid2.TextMatrix(x, 4)) <= 0 Then
          MSHFlexGrid2.TextMatrix(x, 8) = 0
       End If
       
       If Val(MSHFlexGrid2.TextMatrix(x, 8)) <= 0 Then
          MSHFlexGrid2.TextMatrix(x, 8) = 0
       End If
       
       
       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(x, 5))
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(x, 7))
       MSHFlexGrid2.TextMatrix(x, 8) = Format(MSHFlexGrid2.TextMatrix(x, 8), "#######0.00")
       MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(x, 8))
    End If
    x = x + 1
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max

If MSHFlexGrid2.Rows > 3 Then
   CmdSave.Enabled = True
End If


End Sub

Private Sub CmbMonth_GotFocus()

tmp = CmbMonth.Text
CmbMonth.Clear
For i = 1 To 12
    CmbMonth.AddItem MonthName(i), i - 1
Next

CmbMonth.Text = tmp

End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For i = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(i) Then
       CmbMonth.Text = CmbMonth.List(i)
       mMonthListindex = i '+ 1
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If

End Sub


Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For i = 2010 To Year(Now)
    CmbYear.AddItem i
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For i = 2010 To Year(Now)
    If CmbYear.Text = i Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

End Sub


Private Sub CmdSave_Click()

If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 3 Then
   MsgBox " No data in Grid  !!!"
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
mNoofdays = ""

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If


If ChkCheckAll.Value = 0 Then
   tmp = "Select * from Txn_MonthlyInsuranceCharges Where YearMonth = '" & mFromDate & "'"
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Insurance Charges already exist for selected Month!!!"
      CmbMonth.SetFocus
      Exit Sub
   End If
Else
   tmp = "Delete from Txn_MonthlyInsuranceChargesDetail Where TxnID in (Select TXNID from Txn_MonthlyInsuranceCharges Where YearMonth = '" & mFromDate & "')"
   Call TmpTable

   tmp = "Delete from Txn_MonthlyInsuranceCharges Where YearMonth = '" & mFromDate & "'"
   Call TmpTable
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

CmdSave.Enabled = False

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid2.Rows + 5
tmp = "Select * from Txn_MonthlyInsuranceCharges Where YearMonth = '" & mFromDate & "'"
Call TmpTable

If TmpRs.RecordCount = 0 Then
   TmpRs.AddNew
   TmpRs!TxnID = GetMaxTxnID
   mTxnID = TmpRs!TxnID
   TmpRs!TxnDate = Date
   TmpRs!YearMonth = CDate(mFromDate)
   TmpRs!InsuranceAmount = Val(MSHFlexGrid2.TextMatrix(0, 8))
   TmpRs!ActualAmount = 0
   TmpRs!Remarks = "Generated by Process"
   TmpRs!G_UID = GetGUID
   TmpRs!CreatedBy = Gen_UserLoginID
   TmpRs!CreatedDate = Now
   TmpRs.Update
   TmpRs.Requery
   TmpRs.Close
   
   tmp = "Delete from Txn_MonthlyInsuranceChargesDetail Where TxnID = " & mTxnID & ""
   Call TmpTable
   
   tmp = "Select * from Txn_MonthlyInsuranceChargesDetail Where TxnID = " & mTxnID & " "
   Call TmpTable
   If TmpRs.RecordCount = 0 Then
      For i = 2 To MSHFlexGrid2.Rows - 2
          If Val(MSHFlexGrid2.TextMatrix(i, 8)) > 0 Then
             TmpRs.AddNew
             TmpRs!DetailID = GetMaxDetailID
             TmpRs!TxnID = mTxnID 'GetMaxTxnID
             TmpRs!SM_Prefix = MSHFlexGrid2.TextMatrix(i, 1)
             TmpRs!GodownID = Val(MSHFlexGrid2.TextMatrix(i, 2))
             TmpRs!CommodityID = Val(MSHFlexGrid2.TextMatrix(i, 3))
             TmpRs!BalanceBags = Val(MSHFlexGrid2.TextMatrix(i, 4))
             TmpRs!BalanceWeight = Val(MSHFlexGrid2.TextMatrix(i, 5))
             TmpRs!CommodityRate = Val(MSHFlexGrid2.TextMatrix(i, 6))
             TmpRs!AUM = Val(MSHFlexGrid2.TextMatrix(i, 7))
             TmpRs!InsuranceAmount = Val(MSHFlexGrid2.TextMatrix(i, 8))
             TmpRs!G_UID = GetGUID
             TmpRs!CreatedBy = Gen_UserLoginID
             TmpRs!CreatedDate = Now
             TmpRs.Update
          'TmpRs.Requery
          End If
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
      ProgressBar1.Value = ProgressBar1.Max
      MsgBox "Done!!!"
      Call Grid_Head1
   End If
End If



End Sub
Private Function GetMaxTxnID() As Double
Dim Retval As Double
tmp = "Select max(TxnID) from Txn_MonthlyInsuranceCharges"
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
GetMaxTxnID = Retval
End Function
Private Function GetMaxDetailID() As Double
Dim Retval As Double
tmp = "Select max(DetailID) from Txn_MonthlyInsuranceChargesDetail"
Call Tmp1Table
Retval = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
Retval = Retval + 1
GetMaxDetailID = Retval
End Function


Private Sub CmdShowLocation_Click()

Call Grid_Head
ChkCheckAll.Value = 0
CmdProcess.Enabled = True
ProgressBar1.Value = 0

If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

'-------Commented By Sagar on 26-08-211

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

Dim mChooseLocation As Variant

mChooseLocation = ""


If Gen_DBType = "MDB" Then
    tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
    tmp = tmp & " Where TIG.YearMonth = #" & mFromDate & "# "
    tmp = tmp & " And TIG.SM_Prefix not in ("
    tmp = tmp & " Select Distinct TIG.SM_Prefix From Txn_MonthlyInsuranceCharges TIG"
    tmp = tmp & " Where TIG.YearMonth = #" & mFromDate & "# "
    
Else
    tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
    tmp = tmp & " Where TIG.YearMonth = '" & mFromDate & "' "
    tmp = tmp & " And TIG.SM_Prefix not in ("
    tmp = tmp & " Select Distinct TIG.SM_Prefix From Txn_MonthlyInsuranceCharges TIG"
    tmp = tmp & " Where TIG.YearMonth = '" & mFromDate & "' "
End If

Call TmpTable

If TmpRs.RecordCount = 0 Then Exit Sub

For i = 1 To TmpRs.RecordCount
    If mChooseLocation = "" Then
       mChooseLocation = "(" & "'" & TmpRs!SM_Prefix & "'"
    Else
       mChooseLocation = mChooseLocation & "," & "'" & TmpRs!SM_Prefix & "'"
    End If
    
    If i = TmpRs.RecordCount Then
       mChooseLocation = mChooseLocation & ")"
    End If
    TmpRs.MoveNext
Next

tmp = "Select Distinct ML.LocationName,ML.SM_Prefix  "
tmp = tmp & " From Mst_Location ML "
tmp = tmp & " Left Join Mst_CMP  MC On ML.LocationID = MC.LocationID "
tmp = tmp & " Left Join Mst_GSL MG On MC.CMPID = MG.CMPID "

If mChooseLocation <> "" Then
   tmp = tmp & " Where ML.SM_Prefix in " & mChooseLocation & ""
End If

If Gen_WcLocation <> "" Then
   If mChooseLocation = "" Then
      tmp = tmp & " Where ML." & Gen_WcLocation
   Else
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

tmp = tmp & " Order By ML.LocationName "

Call TmpTable

ProgressBar1.Value = 0

ProgressBar1.Max = TmpRs.RecordCount + 1
MSHFlexGrid1.Rows = TmpRs.RecordCount + 2

With MSHFlexGrid1
    For mR = 1 To TmpRs.RecordCount
        .TextMatrix(mR, 1) = TmpRs!LocationName
        .TextMatrix(mR, 2) = TmpRs!SM_Prefix
        
        .row = mR
        .Col = 0
        .CellFontName = "Wingdings"
        .CellFontSize = 12
        .CellAlignment = flexAlignCenterCenter
        .Text = strUnChecked
         
        TmpRs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Next
End With

ProgressBar1.Value = ProgressBar1.Max
LblTotalRecord.Refresh
LblTotalRecord.Caption = TmpRs.RecordCount
mCountLocation = 0

End Sub

Private Sub ChkCheckAll_Click()
'If MSHFlexGrid2.Rows = 2 Then
'   ChkCheckAll.Value = 0
'   Exit Sub
'End If
'
'If ChkCheckAll.Value = vbChecked Then
'   For i = 1 To MSHFlexGrid1.Rows - 2
'       MSHFlexGrid1.Row = i
'       MSHFlexGrid1.Col = 0
'       MSHFlexGrid1.TextMatrix(i, 0) = strChecked
'   Next
'Else
'   For i = 1 To MSHFlexGrid1.Rows - 2
'       MSHFlexGrid1.Row = i
'       MSHFlexGrid1.TextMatrix(i, 0) = strUnChecked
'   Next
'End If
End Sub

Private Function ChkSelection() As Boolean
Dim Retval As Boolean
Retval = False

For k = 1 To MSHFlexGrid1.Rows
    If MSHFlexGrid1.TextMatrix(k, 1) = "" Then Exit For
    If MSHFlexGrid1.TextMatrix(k, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
ChkSelection = Retval
End Function


Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 3 '26 '27 '23 '25
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(0) = 500

     .TextMatrix(0, 0) = "Select"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "SM_Prefix"
          
     .ColWidth(0) = 1000
     .ColWidth(1) = 4600
     .ColWidth(2) = 1000
          
End With
End Sub

Private Sub Grid_Head1()
With MSHFlexGrid2
     .Clear
     .Rows = 3
     .Cols = 9
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(1) = 500

     .TextMatrix(1, 0) = "Location ID"
     .TextMatrix(1, 1) = "SM_Prefix"
     .TextMatrix(1, 2) = "Godown ID"
     .TextMatrix(1, 3) = "Commodity ID"
     .TextMatrix(1, 4) = "Balance Bags"
     .TextMatrix(1, 5) = "Balance Qty"
     .TextMatrix(1, 6) = "Commodity Rate"
     .TextMatrix(1, 7) = "AUM"
     .TextMatrix(1, 8) = "Insurance Charges"
     
     .ColWidth(0) = 1000
     .ColWidth(1) = 1000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1000
     .ColWidth(5) = 1000
     .ColWidth(6) = 1000
     .ColWidth(7) = 1000
     .ColWidth(8) = 2000
          
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

'Call Grid_Head
Call Grid_Head1

Frame1.Enabled = True
'CmdDelete.Visible = False
'ChkCheckAll.Visible = False
'If LCase(Gen_UserRole) = "power" Then
'   CmdDelete.Visible = True
   ChkCheckAll.Visible = True
'End If
CmdProcess.Enabled = True
CmdSave.Enabled = False
ProgressBar1.Value = 0
End Sub
