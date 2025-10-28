VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_ReservationStockStatement 
   Caption         =   "Reservation Stock Statement"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12945
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   12945
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9060
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15090
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
         Left            =   7740
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   375
         Width           =   3255
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
         Left            =   3915
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   375
         Width           =   3765
      End
      Begin VB.CommandButton CmdStockExcel 
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
         Left            =   12345
         TabIndex        =   7
         Top             =   765
         Width           =   1320
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   7755
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   13679
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
      Begin VB.ComboBox CmbMSC 
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
         Left            =   2340
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   795
         Width           =   8640
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
         Left            =   11025
         TabIndex        =   6
         Top             =   765
         Width           =   1320
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
         Left            =   13665
         TabIndex        =   8
         Top             =   765
         Width           =   1200
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
         ItemData        =   "FrmRpt_ReservationStockStatement.frx":0000
         Left            =   2355
         List            =   "FrmRpt_ReservationStockStatement.frx":0002
         TabIndex        =   2
         Top             =   375
         Width           =   1500
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
         Left            =   15960
         TabIndex        =   11
         Top             =   360
         Width           =   1800
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
         ItemData        =   "FrmRpt_ReservationStockStatement.frx":0004
         Left            =   120
         List            =   "FrmRpt_ReservationStockStatement.frx":0006
         TabIndex        =   1
         Top             =   375
         Width           =   2175
      End
      Begin VB.Label Label39 
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
         Height          =   210
         Left            =   8640
         TabIndex        =   17
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label42 
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
         Height          =   225
         Left            =   5280
         TabIndex        =   16
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label12 
         Caption         =   "Contract Name"
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
         TabIndex        =   14
         Top             =   825
         Width           =   1815
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
         Left            =   2850
         TabIndex        =   13
         Top             =   135
         Width           =   495
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
         Left            =   960
         TabIndex        =   12
         Top             =   135
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   9045
      Width           =   15090
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   75
         TabIndex        =   9
         Top             =   210
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_ReservationStockStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mTotalDays, mToDate, mNoofDays As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mGenInvoiceLocation As Variant
Dim mLocationID, mSLocationID, mContractID, mCountLocation As Variant
Dim mShowDTD As Boolean
Dim mCommodityID, mSpilBags, mSpilQty As Variant

Private Sub CmbMonth_GotFocus()

tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I), I - 1
Next

CmbMonth.Text = tmp

End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mMonthListindex = I '+ 1
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
For I = 2011 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For I = 2010 To Year(Now)
    If CmbYear.Text = I Then
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

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdProcess_Click()

'On Error GoTo msgError

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

If CmbMSC.Text = "" Then
   MsgBox " Blank contract Name not allowed  !!!"
   CmbMSC.SetFocus
   Exit Sub
End If

''If MSHFlexGrid1.Rows < 3 Then
''   MsgBox " No Location populated in Grid  !!!"
''   Exit Sub
''End If

''If ChkSelection = False Then
''   MsgBox " No Location selected from Grid  !!!"
''   Exit Sub
''End If
''
mMonth = mMonthListindex + 1

If mMonth < 10 Then
   mMonth = "0" & mMonth
End If


mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

mNoofDays = ""

If CDate(Format(mFromDate, "yyyymm")) > Format(Date, "yyyymm") Then
   MsgBox "Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

CmdProcess.Enabled = False


mGenInvoiceLocation = "SM_Prefix in ('" & mSM_Prefix & "')"

'If CmbPreFix.Text <> "" Then
'   tmp = "Select Count(*) as 'TotalRows' from Txn_InvMonGSLDetail Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & mFromDate & "' And TFlag = 'O' "
''Else
''   tmp = "Select Count(*) as 'TotalRows' from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And TFlag = 'O' "
'
'  Call TmpTable
'
'  If TmpRs!TotalRows > 0 Then
'     If CmbPreFix.Text <> "" Then
'        MsgBox "Opening balance process for " & CmbPreFix.Text & " " & CmbMonth.Text & "-" & CmbYear.Text & "  month is already executed!!!"
''     Else
''        MsgBox "Opening balance process for this month is already executed!!!"
'     End If
'     Exit Sub
'  End If
'
'
'End If

''If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" Then
''   If mShowDTD = False Then
''      MsgBox "Please Complete all check points before Processing Opening Balance!!!"
'''      CmdProcess.Enabled = True
''      Exit Sub
''   End If
''   If TxtConfirmation.Text = "" Then
''      MsgBox "Blank Confirmation message not allowed!!!"
''      TxtConfirmation.SetFocus
'''      CmdProcess.Enabled = True
''      Exit Sub
''   End If
''End If

Gen_mTimeStamp = GetTimeStamp

'-- Create Temp Table with record of GSL
MsgBox "Wait till next message!!!"

Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)


tmp = " Alter table TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " Add Mst_GSL_ID_New int null;"
Call TmpTable

tmp = " Update TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " Set Mst_GSL_ID_New = Mst_GSL_ID"
Call TmpTable

tmp = " Alter table TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " drop column Mst_GSL_ID;"
Call TmpTable

tmp = " EXEC sp_rename 'TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & ".Mst_GSL_ID_New', 'Mst_GSL_ID', 'COLUMN';"
Call TmpTable


'-- Find Distinct location for those location which are having Withdral txn.

tmp = " Select Distinct TCG.SM_Prefix"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Where TCD.ExchangeTypeID=1 And TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And " & mGenInvoiceLocation
tmp = tmp & " And TCG." & mGenInvoiceLocation
tmp = tmp & " Order By TCG.SM_Prefix"

Call TmpTable

'If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 1
   For I = 1 To TmpRs.RecordCount
       
       '---fetch Withdrawal record for adding back to opening stock. which is > [To Date]
'
'       tmp = "Select TCG.GSLID,TCG.SM_Prefix,TCG.CxTFNo,TCD.CxTFDate,TCG.NoOfBags,TCG.DematWeight,TCG.TxnType"
'       tmp = tmp & " From Txn_CxTF_GSL  TCG"
'       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
'       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
'       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '" & mFromDate & "' And  '" & Format(Date, Gen_DatFormat) & "' "
'       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
'       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID,TCD.CxTFDate"
'

       tmp = "Select TCG.Mst_GSL_ID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
       tmp = tmp & " From Txn_CxTF_GSL  TCG"
       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate > '" & mToDate & "' "
       tmp = tmp & " And TCD.ExchangeTypeID=1 And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'" 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
       tmp = tmp & " Group By TCG.SM_Prefix,TCG.Mst_GSL_ID "
       tmp = tmp & " Order By TCG.SM_Prefix,TCG.Mst_GSL_ID "
       
       
       Call Tmp1Table
       
       '--Update temp table with Withdrawal transaction record to get opening stock
       
      
       Call TableTxn_TempGSL(" Where SM_Prefix = '" & TmpRs!SM_Prefix & "'")
       
      
       If RsTxn_TempGSL.RecordCount > 0 Then
          For j = 1 To TmpRs1.RecordCount
              If RsTxn_TempGSL.RecordCount > 0 Then
                 RsTxn_TempGSL.MoveFirst
              End If
              RsTxn_TempGSL.Find "Mst_GSL_ID=" & TmpRs1!Mst_GSL_ID
              If Not RsTxn_TempGSL.EOF Then
                 RsTxn_TempGSL!BalanceBags = RsTxn_TempGSL!BalanceBags + TmpRs1!NoofBags
                 RsTxn_TempGSL!BalanceWeight = RsTxn_TempGSL!BalanceWeight + TmpRs1!DematWeight
                 RsTxn_TempGSL!CDTFBags = RsTxn_TempGSL!CDTFBags - TmpRs1!NoofBags
                 RsTxn_TempGSL!CDTFWeight = RsTxn_TempGSL!CDTFWeight - TmpRs1!DematWeight
                 RsTxn_TempGSL!Flag = "O"
                 RsTxn_TempGSL.Update
              End If
           TmpRs1.MoveNext
          Next
       End If
       
              '--------Added to add Opening balance of Full Withdrawal GSLs for Reservation with 'P' TFlag
       
       tmp = "Select TCG.Mst_GSL_ID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
       tmp = tmp & " From Txn_CxTF_GSL  TCG"
       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' And  TCD.CxTFDate <= '" & mToDate & "' "
       tmp = tmp & " And TCD.ExchangeTypeID=1 And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "' And isnull(MGSL.NContractID,0)>0" 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
       tmp = tmp & " Group By TCG.SM_Prefix,TCG.Mst_GSL_ID "
       tmp = tmp & " Order By TCG.SM_Prefix,TCG.Mst_GSL_ID "
       
       Call Tmp1Table
       
       '--Update temp table with Withdrawal transaction record to get opening stock
             
       Call TableTxn_TempGSL(" Where SM_Prefix = '" & TmpRs!SM_Prefix & "'")
       
       If RsTxn_TempGSL.RecordCount > 0 Then
          For j = 1 To TmpRs1.RecordCount
              If RsTxn_TempGSL.RecordCount > 0 Then
                 RsTxn_TempGSL.MoveFirst
              End If
              
              RsTxn_TempGSL.Find "Mst_GSL_ID=" & TmpRs1!Mst_GSL_ID
              If Not RsTxn_TempGSL.EOF Then
                 RsTxn_TempGSL!BalanceBags = RsTxn_TempGSL!BalanceBags + TmpRs1!NoofBags
                 RsTxn_TempGSL!BalanceWeight = RsTxn_TempGSL!BalanceWeight + TmpRs1!DematWeight
                 RsTxn_TempGSL!CDTFBags = RsTxn_TempGSL!CDTFBags - TmpRs1!NoofBags
                 RsTxn_TempGSL!CDTFWeight = RsTxn_TempGSL!CDTFWeight - TmpRs1!DematWeight
                 RsTxn_TempGSL!Flag = "P"
                 RsTxn_TempGSL.Update
              End If
           TmpRs1.MoveNext
          Next
       End If

       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   
   
   ProgressBar1.Value = 0
   
   If TmpRs.RecordCount < 50 Then
      ProgressBar1.Max = 1000
   End If
   
   '-- Delete All data from temp table having Flag='E'
   
   tmp = "Delete from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where Flag = 'E' And isnull(NContractID,0)=0"
   Call TmpTable
   
   ProgressBar1.Value = ProgressBar1.Value + 1
   
   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Drop Column BillingCycleID "
   Call TmpTable
   
   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Drop Column BillingUnit "
   Call TmpTable
   
   
   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Drop Column Rate "
   Call TmpTable
   
   
   
   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Add YearMonth DateTime,BillingCycleID decimal(18,0),BillingUnit Varchar(1),Drno decimal(18,0),Rate decimal(18,4), "
   tmp = tmp & " InvoiceAmount decimal(18,2),TFLag Varchar(1),CxTFDate DateTime,InvoiceDate DateTime,DateOfWithdrawal Datetime ,MonG_UID Varchar(36) "
   'CxTF_DetailsID Decimal(18,0),
   Call TmpTable
      
   If Gen_DBType = "MDB" Then
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = #" & mFromDate & "#"
   Else
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = '" & mFromDate & "'"
   End If
     
   Call TmpTable
   
   tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set MonG_UID = newID(),YearMonth = '" & mFromDate & "'"
   Call TmpTable
   
   ProgressBar1.Value = ProgressBar1.Value + 1
      
   tmp = " Insert into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " (Mst_GSL_ID,SM_Prefix,TFlag,CxTF_DetailsID,CxTFDate,CMPID,"
   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
   tmp = tmp & " ClientID,BalanceBags,BalanceWeight,Flag,Flag1,GrossWt,"
   tmp = tmp & " BagSize,CDTFNo,CDTFDate,KRDate,NContractID,YearMonth,G_UID,CreatedBy,CreatedDate,"
   tmp = tmp & " UpdatedDate , UpdatedBy,KRNo,MonG_UID,GSLID,SpillageFlag,MadeupFLag)"
   tmp = tmp & " Select TCG.Mst_GSL_ID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTF_DetailsID,TCD.CxTFDate,TCD.CMPID,MGSL.GodownID,MGSL.StackNo,"
   tmp = tmp & " MGSL.LotNo,MGSL.CommodityID,MGSL.ExchangeTypeID,"
   tmp = tmp & " MGSL.ClientID,TCG.NoOfBags,TCG.DematWeight,MGSL.Flag,MGSL.Flag1,MGSL.GrossWt,"
   tmp = tmp & " MGSL.BagSize,MGSL.CDTFNo,MGSL.CDTFDate,MGSL.KRDate,MGSL.NContractID,'" & mFromDate & "',MGSL.G_UID,MGSL.CreatedBy,MGSL.CreatedDate,"
   tmp = tmp & " MGSL.UpdatedDate , MGSL.UpdatedBy,MGSL.KRNo,newid(),'0' 'GSLID',SpillageFlag,MadeupFLag"
   tmp = tmp & " From Txn_CxTF_GSL  TCG"
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix"
   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
   tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
   tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
   
   'Commented on 19 July 2012 after discussion with mayur sir due to leakage in stock data of NCDEX
   'tmp = tmp & " Where MGSL.SpillageFlag = 0 And (TCG.BillFlag is Null or TCG.BillFlag = '')"
   
   tmp = tmp & " Where MGSL.SpillageFlag = 0 "
   tmp = tmp & " And TCD.ExchangeTypeID=1 and TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' "
   
'Commented on 12 Nov 2010 Parag
'   If CmbPreFix.Text <> "" Then
'      tmp = tmp & " and TCG.SM_Prefix = '" & mSM_Prefix & "'"
'   End If
   
   tmp = tmp & " And TCG." & mGenInvoiceLocation
   
   Call TmpTable
      
   
   tmp = "Select * from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set BagSize = Round(Bagsize,0) "
      Call TmpTable
       
      ProgressBar1.Value = ProgressBar1.Value + 1
       
       ' Flag off all GSL which are used for selected month in transaction table so to avoid duplicate billing
       
'''       tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillFlag = 'I'"
'''       tmp = tmp & " Where  GSLDetailID in (Select TCG.GSLDetailID"
'''       tmp = tmp & " From Txn_CxTF_GSL  TCG"
'''       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix"
'''       tmp = tmp & " And TCG.TxnType = TCD.TxnType"
'''       tmp = tmp & " where  (TCG.BillFlag is Null or TCG.BillFlag = '')  And"
'''       tmp = tmp & " TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix) "
'''
'''       'tmp = tmp & " TCD.CxTFDate <= '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix)"
'''
'''    'Commented on 12 Nov 2010 Parag
'''    '   If CmbPreFix.Text <> "" Then
'''    '      tmp = tmp & " and Txn_CxTF_GSL.SM_Prefix = '" & mSM_Prefix & "'"
'''    '   End If
'''
'''       tmp = tmp & " And Txn_CxTF_GSL." & mGenInvoiceLocation
'''
'''       Call TmpTable
'''
'''       ProgressBar1.Value = ProgressBar1.Value + 1
    
    
'''        ' For Op and With
'''
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
'''        tmp = tmp & " (Select a.BillingCycleID"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'''        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'For Depo
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
'''        tmp = tmp & " (Select a.DepoBillingCycleID"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'''        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'''        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
    
    
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = "
'''        tmp = tmp & " (Select a.BillingUnit"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'''        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
        '==============================================
        
        
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Drno = "
'''        tmp = tmp & " (Select a.SPRCID"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'''        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
''        tmp = tmp & " (Select a.Amount"
''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
''        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
''        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''        If CmbPreFix.Text <> "" Then
''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
''        End If
''
''        Call TmpTable
''
''
''        ProgressBar1.Value = ProgressBar1.Value + 1
         
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = Null Where YearMonth = '" & mFromDate & "'"
         tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        
        Call TmpTable
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
            
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'S' Where DRNo is not null And YearMonth = '" & mFromDate & "' "
        tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        
        Call TmpTable
    
        
        ProgressBar1.Value = ProgressBar1.Value + 1
              
              
''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
''        tmp = tmp & " (Select  MSRCD.RatePerBag"
''        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
''        tmp = tmp & " Inner Join Mst_SpecialRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
''        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
''        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
''        tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
''        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
''
''        If CmbPreFix.Text <> "" Then
''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
''        End If
''
''
''        Call TmpTable
''
''
''        ProgressBar1.Value = ProgressBar1.Value + 1

       
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'''        tmp = tmp & " (Select  MSRCD.RatePerBag"
'''        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'''        tmp = tmp & " Inner Join Mst_SpecialRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
'''        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'''        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'''        tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
'''        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
       
       '--------------------------------------------------
       '| B. Apply  Base Rate to GSL Table.                |
       '--------------------------------------------------
       
       
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
'''        tmp = tmp & " (Select a.BillingCycleID"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
'''        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'''        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
'''        tmp = tmp & " (Select a.DepoBillingCycleID"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
'''        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'''        tmp = tmp & " a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
'''        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
    
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit ="
'''        tmp = tmp & " (Select a.Flag"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'''        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'''        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
    
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo ="
'''        tmp = tmp & " (Select a.SRCID"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'''        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'''        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
        
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
'''        tmp = tmp & " (Select a.RateOnQty"
'''        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'''        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'''        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'N' Where DRNo is not null And YearMonth = '" & mFromDate & "' And Flag1 is null And BillingCycleID is not Null And ExchangeTypeID = 1 "
        tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        
        Call TmpTable
       
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'''        tmp = tmp & " (Select MSRCD.RatePerBag"
'''        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'''        tmp = tmp & " Inner Join Mst_StorageRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
'''        tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
'''        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'''        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
'''        tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
'''        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'''        tmp = tmp & " (Select MSRCD.RatePerBag"
'''        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'''        tmp = tmp & " Inner Join Mst_StorageRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
'''        tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
'''        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'''        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
'''        tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'''        tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
'''        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
       '--------------------------------------------------
       '| C. Apply  No Of Cycles to GSL Table as per Billing Unit and Transaction Flag.                |
       '--------------------------------------------------
    
        'Op Bal and Monthly Cycle
''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''        tmp = tmp & " And BillingCycleID = 4  And (ContractID = 0 Or ContractID is Null)"
''
''        If CmbPreFix.Text <> "" Then
''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''        End If
''
''        Call TmpTable
''
''
''
''        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Op Bal and Forthnightly Cycle
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'''        tmp = tmp & " And BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        'Op Bal and Weekly Cycle
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'''        tmp = tmp & " And BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''        'Op Bal and Daily Cycle
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'''        tmp = tmp & " And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        'Depo and Withd and Monthly Cycle
'''
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag <> 'O'"
'''        tmp = tmp & " And BillingCycleID = 4 And (ContractID = 0 Or ContractID is Null)"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
         
        'Depo and Withd and Daily Cycle
                
'''        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  " 'And DateofWithdrawal is Null "
'''
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        'Withd and Daily Cycle
'''
'''        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''        ProgressBar1.Value = ProgressBar1.Value + 1

        'Deposit & Weekly
        
'''        mNoofdays = GetBillingCycleDays(2)
'''
'''        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) " 'And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DateOfWithdrawal is null "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
                
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        'Deposit & FortNightly
'''        mNoofdays = GetBillingCycleDays(3)
'''
'''        'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 2 "
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
'''        tmp = tmp & " And Day(CxTFDate) <= 15 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        'Case 1 Depo Date  > 15 Then No Of Cycle = 1
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
'''        tmp = tmp & " And Day(CxTFDate) > 15 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Withdrawal & weekly
'''        mNoofdays = GetBillingCycleDays(2)
'''
'''        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        'Withdrawal & Fortnightly
'''        mNoofdays = GetBillingCycleDays(3)
'''
'''
'''        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
'''        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        'Updating Invoice Amount And Invoice Date
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceBags * NoOfDays * Rate, InvoiceDate = '" & mToDate & "' "
'''        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)"
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '" & mToDate & "' "
'''        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'''        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
'''
'''    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month(CxTFDate) = Month(CDTFDate) "
'''
'''        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,CxTFDate) = 0 "
'''
'''        If CmbPreFix.Text <> "" Then
'''           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''        End If
'''
'''
'''        Call TmpTable
'''
'''        tmp = " Update Txn_InvMonGSLDetail Set OldBalanceWeight = BalanceWeight Where BillingUnit = 'Q' And BalanceWeight < 0 "
'''        Call TmpTable
'''
'''        tmp = " Update Txn_InvMonGSLDetail Set BalanceWeight = (BagSize * BalanceBags)/1000 Where BillingUnit = 'Q' And BalanceWeight < 0 "
'''        Call TmpTable
'''
'''        tmp = " Update Txn_InvMonGSLDetail Set InvoiceAmount = BalanceWeight*Rate*NoofDays Where BillingUnit = 'Q' And OldBalanceWeight < 0 "
'''        Call TmpTable
'''
'''
'''        ProgressBar1.Value = ProgressBar1.Value + 1
        
        
'''        tmp = " Insert into Txn_InvMonGSLDetail "
'''        tmp = tmp & " (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
'''        tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,Tflag,GrossWt,KRDate,"
'''        tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
'''        tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate,KRNo,ProcessedBy,MonG_UID)"
'''        tmp = tmp & " Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
'''        'tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,'O',GrossWt,KRDate,"
'''
'''        tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,TFlag,GrossWt,KRDate,"
'''        tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
'''        tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate,KRNo,'" & Gen_UserLoginID & "',MonG_UID "
'''        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
'''
'''        Call TmpTable
       
'''        tmp = " Select SM_Prefix from Mst_Location"
'''        tmp = tmp & " Where SM_Prefix not in "
'''        tmp = tmp & " (Select Distinct SM_Prefix From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ") "
'''        tmp = tmp & " And " & mGenInvoiceLocation & " "
'''
'''        Call TmpTable
'''
'''        For I = 1 To TmpRs.RecordCount
'''
'''            tmp = " Insert into Txn_InvMonGSLDetail "
'''            tmp = tmp & " (GSLID,SM_Prefix,CMPID,StackNo,LotNo, "
'''            tmp = tmp & " YearMonth,G_UID,CreatedBy,CreatedDate,ProcessedBy "
'''            tmp = tmp & " )"
'''            tmp = tmp & " Values (-1,'" & TmpRs!SM_Prefix & "',-1,'Dummy-P','Dummy-P','" & mFromDate & "',newid(),'" & Gen_UserLoginID & "','" & Now() & "','" & Gen_UserLoginID & "')"
'''
'''            Call Tmp1Table
'''            TmpRs.MoveNext
'''
'''        Next
    Else
'''
'''        tmp = "Select SM_prefix from Mst_Location Where " & mGenInvoiceLocation & " "
'''        Call TmpTable
'''
'''        For I = 1 To TmpRs.RecordCount
'''
'''            tmp = " Insert into Txn_InvMonGSLDetail "
'''            tmp = tmp & " (GSLID,SM_Prefix,CMPID,StackNo,LotNo, "
'''            tmp = tmp & " YearMonth,G_UID,CreatedBy,CreatedDate,ProcessedBy "
'''            tmp = tmp & " )"
'''            tmp = tmp & " Values (-1,'" & TmpRs!SM_Prefix & "',-1,'Dummy-P','Dummy-P','" & mFromDate & "',newid(),'" & Gen_UserLoginID & "','" & Now() & "','" & Gen_UserLoginID & "')"
'''
'''            Call Tmp1Table
'''
'''            TmpRs.MoveNext
'''        Next
    End If
    Call FillDaywiseBalance(I)
    
    tmp = "Drop Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
    Call TmpTable
    
'End If


ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done !!"
CmdProcess.Enabled = True
'CmdProcess.Enabled = True



Exit Sub

msgError:
    
    tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
    tmp = tmp & " drop table dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
    Call TmpTable
    
    MsgBox "Some problem occured during process!!!"
      

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

Call Grid_Head1
ProgressBar1.Value = 0

End Sub

Private Function ChkBOD() As Boolean
Dim Retval As Boolean
Retval = True

tmp = "Select ML.LocationName"
tmp = tmp & " From Mst_CMP MC Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
'tmp = tmp & " Where MC.CloseDate Is Null "

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where MC.CurrentDate <= #" & Format(mToDate, Gen_DatFormat) & "# "
Else
   tmp = tmp & " Where MC.CurrentDate <= '" & Format(mToDate, Gen_DatFormat) & "' "
End If

tmp = tmp & " And MC.Closedate is null "
tmp = tmp & " And ML." & mGenInvoiceLocation
tmp = tmp & " Order by ML.LocationName,MC.CMPName"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   Retval = False
End If

ChkBOD = Retval

End Function


Private Sub CmbMSC_GotFocus()
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

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")


If CDate(Format(mFromDate, "yyyymm")) > Format(Date, "yyyymm") Then
   MsgBox "Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

tmp1 = CmbMSC.Text

tmp = "select  ContractName  + '~' + convert(varchar(30),ContractID)  as ContractName,LocationID ,ContractID from Mst_StorageContract Where Flag='A'"

' commented by sagar as CMG wnat to check Balance of all active Reservation
''tmp = tmp & " And BookedFrom <= '" & mToDate & "' And  ExtendedPeriod >= '" & mFromDate & "'"
If CmbSLocation.Text <> "" Then
   tmp = tmp & " And LocationID=" & mSLocationID
End If
If CmbSCommodity.Text <> "" Then
   tmp = tmp & " And CommodityID=" & mCommodityID
End If

Call Tmp4Table

CmbMSC.Clear
If TmpRs4.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs4.RecordCount
    CmbMSC.AddItem TmpRs4!ContractName
    TmpRs4.MoveNext
Next
CmbMSC.Text = tmp1
End Sub
Private Sub CmbMSC_LostFocus()
If CmbMSC.Text = "" Then
   Exit Sub
End If
TmpRs4.MoveFirst
TmpRs4.Find "ContractName = '" & CmbMSC.Text & "'"
If TmpRs4.EOF Then
   MsgBox "Invalid selection "
   CmbMSC.SetFocus
   Exit Sub
End If
If mContractID <> TmpRs4!ContractID Then
   mContractID = TmpRs4!ContractID
   mLocationID = TmpRs4!LocationID
   mSM_Prefix = GetPreFixByLocation(TmpRs4!LocationID, "P")
   Call Grid_Head1
End If

End Sub

Private Sub FillDaywiseBalance(rowID_ As Variant)
mTableName = "TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""

tmp = "Select ML.LocationID, ML.LocationName, TINV.NContractID as ContractID,Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight',TINV.YearMonth"
tmp = tmp & " From " & mTableName & " TINV"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag in ('O','P') And TINV.NContractID>0 And TINV.MadeupFlag=0"
tmp = tmp & " Group by ML.LocationID, ML.LocationName, TINV.NContractID,TINV.YearMonth"

Call Tmp2Table

''tmp = "Select ML.LocationID, ML.LocationName, TINV.ContractID,Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight',TINV.YearMonth"
''tmp = tmp & " From " & mTableName & " TINV"
''tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
''tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
''tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
''tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
''tmp = tmp & " Where TINV.YearMonth = '" & mPrevMonthDate & "' And TINV.TFlag='O' And TINV.ContractID>0"
''tmp = tmp & " And GSLID in (select GSLID From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix=TINV.SM_Prefix)"
''tmp = tmp & " Group by ML.LocationID, ML.LocationName, TINV.ContractID,TINV.YearMonth"

tmp = "Select ML.LocationID, ML.LocationName, TINV.NContractID as ContractID,Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight',TINV.YearMonth,TINV.CxTFDate ,TINV.TFlag,TINV.CxTF_DetailsID"
tmp = tmp & " From " & mTableName & "  TINV"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag in ('D','W') And TINV.NContractID>0"
tmp = tmp & " Group by ML.LocationID, ML.LocationName, TINV.NContractID,TINV.CxTFDate ,TINV.TFlag,TINV.YearMonth,TINV.CxTF_DetailsID"

Call Tmp3Table

Call Grid_Head1
MSHFlexGrid2.Visible = True
Dim mDispDate As Variant
Dim mFWFrom, mFWTo, mSWFrom, mSWTo, mTWfrom, mTWTo, mFrWFrom, mFrWTo, mFFNEnd, mSFNStart As Variant

mFWFrom = Format(mFromDate, Gen_DatFormat)
mFWTo = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "07", Gen_DatFormat)
mSWFrom = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "08", Gen_DatFormat)
mSWTo = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "14", Gen_DatFormat)
mTWfrom = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "15", Gen_DatFormat)
mTWTo = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "21", Gen_DatFormat)
mFrWFrom = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "22", Gen_DatFormat)
mFrWTo = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & mTotalDays, Gen_DatFormat)

mFFNEnd = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "15", Gen_DatFormat)
mSFNStart = Format(Year(mFromDate) & "-" & Month(mFromDate) & "-" & "16", Gen_DatFormat)

mDispDate = mFromDate

MSHFlexGrid2.Rows = mTotalDays + 2

For I = 1 To mTotalDays

    MSHFlexGrid2.TextMatrix(I, 0) = mLocationID
    MSHFlexGrid2.TextMatrix(I, 1) = GetLocationName(Val(mLocationID)) 'MSHFlexGrid1.TextMatrix(rowID_, 3)
    MSHFlexGrid2.TextMatrix(I, 2) = mContractID 'MSHFlexGrid1.TextMatrix(rowID_, 1)
    MSHFlexGrid2.TextMatrix(I, 3) = Format(mDispDate, "dd-MMM-yyyy")
    
    MSHFlexGrid2.TextMatrix(I, 4) = 0
    MSHFlexGrid2.TextMatrix(I, 5) = 0
    MSHFlexGrid2.TextMatrix(I, 6) = 0
    MSHFlexGrid2.TextMatrix(I, 7) = 0
    MSHFlexGrid2.TextMatrix(I, 8) = 0
    MSHFlexGrid2.TextMatrix(I, 9) = 0
    MSHFlexGrid2.TextMatrix(I, 10) = 0
    MSHFlexGrid2.TextMatrix(I, 11) = 0
    
    mSpilBags = 0
    mSpilQty = 0
    
    If I = 1 Then
       If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
       TmpRs2.Find "ContractID=" & Val(MSHFlexGrid2.TextMatrix(I, 2))
       
       If TmpRs2.EOF = False Then
       
          mSpilBags = 0
          mSpilQty = 0
           
           
          MSHFlexGrid2.TextMatrix(I, 4) = Format(TmpRs2!BalanceBags + mSpilBags, "#####0")
          MSHFlexGrid2.TextMatrix(I, 5) = Format(TmpRs2!BalanceWeight + mSpilQty, "#####0.000")
           
          mSpilBags = 0
          mSpilQty = 0
           
          Call GetOpeningBalanceSpillagePreStock(Val(MSHFlexGrid2.TextMatrix(I, 2)), mSM_Prefix)
           
          MSHFlexGrid2.TextMatrix(I, 4) = Format(Val(MSHFlexGrid2.TextMatrix(I, 4)) - mSpilBags, "#####0")
          MSHFlexGrid2.TextMatrix(I, 5) = Format(Val(MSHFlexGrid2.TextMatrix(I, 5)) - mSpilQty, "#####0.000")
           
           
          '--For Made Up Bags
          mSpilBags = 0
          mSpilQty = 0
           
          mSpilBags = 0
          mSpilQty = 0
           
          Call GetMadeUpOpeningPreBalance(Val(MSHFlexGrid2.TextMatrix(I, 2)), mSM_Prefix)
           
          MSHFlexGrid2.TextMatrix(I, 4) = Format(Val(MSHFlexGrid2.TextMatrix(I, 4)) - mSpilBags, "#####0")
          MSHFlexGrid2.TextMatrix(I, 5) = Format(Val(MSHFlexGrid2.TextMatrix(I, 5)) - mSpilQty, "#####0.000")
          
          '-----------
          
       End If
       
    Else
       MSHFlexGrid2.TextMatrix(I, 4) = MSHFlexGrid2.TextMatrix(I - 1, 10)
       MSHFlexGrid2.TextMatrix(I, 5) = MSHFlexGrid2.TextMatrix(I - 1, 11)
    End If
    
    TmpRs3.Filter = "ContractID=" & Val(MSHFlexGrid2.TextMatrix(I, 2)) & " And  CxTFDate='" & Format(mDispDate, Gen_DatFormat) & "' And TFlag='D'"
    
    If TmpRs3.RecordCount > 0 Then
       For Ln = 1 To TmpRs3.RecordCount
           MSHFlexGrid2.TextMatrix(I, 6) = Val(MSHFlexGrid2.TextMatrix(I, 6)) + Format(TmpRs3!BalanceBags, "#####0")
           MSHFlexGrid2.TextMatrix(I, 7) = Val(MSHFlexGrid2.TextMatrix(I, 7)) + Format(TmpRs3!BalanceWeight, "#####0.000")
           If MSHFlexGrid2.TextMatrix(I, 12) <> "" Then
              MSHFlexGrid2.TextMatrix(I, 12) = MSHFlexGrid2.TextMatrix(I, 12) & "," & IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           Else
              MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           End If
           TmpRs3.MoveNext
       Next
    End If
    
    TmpRs3.Filter = ""
    TmpRs3.Filter = "ContractID=" & Val(MSHFlexGrid2.TextMatrix(I, 2)) & " And  CxTFDate='" & Format(mDispDate, Gen_DatFormat) & "' And TFlag='W'"
   
    If TmpRs3.RecordCount > 0 Then
       For Ln = 1 To TmpRs3.RecordCount
           mSpilBags = 0
           mSpilQty = 0
           Call GetSpillageWithdrawal(mSM_Prefix, TmpRs3!CxTF_DetailsID)
           
           MSHFlexGrid2.TextMatrix(I, 8) = Val(MSHFlexGrid2.TextMatrix(I, 8)) + Format(TmpRs3!BalanceBags + mSpilBags, "#####0")
           MSHFlexGrid2.TextMatrix(I, 9) = Val(MSHFlexGrid2.TextMatrix(I, 9)) + Format(TmpRs3!BalanceWeight + mSpilQty, "#####0.000")
           
           mSpilBags = 0
           mSpilQty = 0
           
           'Call GetMadeUpBagsWithdrawal(mSM_Prefix, TmpRs3!CxTFNo, Val(MSHFlexGrid2.TextMatrix(I, 2)))
           
           MSHFlexGrid2.TextMatrix(I, 8) = Val(MSHFlexGrid2.TextMatrix(I, 8)) + Format(mSpilBags, "#####0")
           MSHFlexGrid2.TextMatrix(I, 9) = Val(MSHFlexGrid2.TextMatrix(I, 9)) + Format(mSpilQty, "#####0.000")
                     
           If MSHFlexGrid2.TextMatrix(I, 12) <> "" Then
              MSHFlexGrid2.TextMatrix(I, 12) = MSHFlexGrid2.TextMatrix(I, 12) & "," & IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           Else
              MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           End If
           
           TmpRs3.MoveNext
       Next
    End If
    
    TmpRs3.Filter = ""
    
    MSHFlexGrid2.TextMatrix(I, 10) = Format(Val(MSHFlexGrid2.TextMatrix(I, 4)) + Val(MSHFlexGrid2.TextMatrix(I, 6)) - Val(MSHFlexGrid2.TextMatrix(I, 8)), "#####0")
    MSHFlexGrid2.TextMatrix(I, 11) = Format(Val(MSHFlexGrid2.TextMatrix(I, 5)) + Val(MSHFlexGrid2.TextMatrix(I, 7)) - Val(MSHFlexGrid2.TextMatrix(I, 9)), "#####0.000")
       
    mDispDate = CDate(mDispDate) + 1
    
Next


End Sub
Private Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 13
       
    .TextMatrix(0, 0) = "LocationID"
    .TextMatrix(0, 1) = "Location Name"
    .TextMatrix(0, 2) = "ContractID"
    .TextMatrix(0, 3) = "Date"
    .TextMatrix(0, 4) = "Opening Balance Bags"
    .TextMatrix(0, 5) = "Opening Balance Qty"
    .TextMatrix(0, 6) = "Deposit Bags"
    .TextMatrix(0, 7) = "Deposit Qty"
    .TextMatrix(0, 8) = "Withdrawal Bags"
    .TextMatrix(0, 9) = "Withdrawal Qty"
    .TextMatrix(0, 10) = "Balance Bags"
    .TextMatrix(0, 11) = "Balance Qty"
    .TextMatrix(0, 12) = "CxTF No"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 2400
    .ColWidth(2) = 1000
    .ColWidth(3) = 1400
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1000
    
    .RowHeight(0) = 700

End With
End Sub

Private Sub GetOpeningBalanceSpillageStock(mContractID_ As Variant, mPrefix_ As Variant)

tmp = "Select TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight'"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "'" ' And  TCD.CxTFDate <= '" & mToDate & "'"
tmp = tmp & " And TCG.SM_Prefix = '" & mPrefix_ & "'"
tmp = tmp & " And TCD.CxTF_DetailsID IN (Select Distinct TINV.CxTF_DetailsID From Txn_InvMonGSLDetail TINV Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag in ('W') And TINV.ContractID=" & mContractID_ & " And SM_Prefix='" & mPrefix_ & "')"
tmp = tmp & " And MGSL.SpillageFlag=1"
tmp = tmp & " Group By TCG.SM_Prefix"
tmp = tmp & " Order By TCG.SM_Prefix"

Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   mSpilBags = IIf(IsNull(TmpRs5!NoofBags), 0, TmpRs5!NoofBags)
   mSpilQty = IIf(IsNull(TmpRs5!DematWeight), 0, TmpRs5!DematWeight)
End If

End Sub

Private Sub GetSpillageWithdrawal(mPrefix_ As Variant, mCITFNo_ As Variant)

tmp = "Select TCG.CxTF_DetailsID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight'"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "'" ' And  TCD.CxTFDate <= '" & mToDate & "'"
tmp = tmp & " And TCG.SM_Prefix = '" & mPrefix_ & "'"
tmp = tmp & " And MGSL.SpillageFlag=1 And TCG.CxTF_DetailsID = " & mCITFNo_
tmp = tmp & " Group By TCG.SM_Prefix,TCG.CxTF_DetailsID"
tmp = tmp & " Order By TCG.SM_Prefix,TCG.CxTF_DetailsID"

Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   mSpilBags = IIf(IsNull(TmpRs5!NoofBags), 0, TmpRs5!NoofBags)
   mSpilQty = IIf(IsNull(TmpRs5!DematWeight), 0, TmpRs5!DematWeight)
End If

End Sub


Private Sub GetOpeningBalanceSpillagePreStock(mContractID_ As Variant, mPrefix_ As Variant)

''tmp = "Select TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight'"
''tmp = tmp & " From Txn_CxTF_GSL  TCG"
''tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
''tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
''tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate < '" & mFromDate & "'" ' And  TCD.CxTFDate <= '" & mToDate & "'"
''tmp = tmp & " And TCG.SM_Prefix = '" & mPrefix_ & "'"
''tmp = tmp & " And TCD.CxTFNo IN (Select Distinct TINV.CxTFNO From Txn_InvMonGSLDetail TINV Where TINV.YearMonth < '" & mFromDate & "' And TINV.TFlag in ('W') And TINV.ContractID=" & mContractID_ & " And SM_Prefix='" & mPrefix_ & "')"
''tmp = tmp & " And MGSL.SpillageFlag=1"
''tmp = tmp & " Group By TCG.SM_Prefix"
''tmp = tmp & " Order By TCG.SM_Prefix"

tmp = "Select TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight'"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate < '" & mFromDate & "'" ' And  TCD.CxTFDate <= '" & mToDate & "'"
tmp = tmp & " And TCG.SM_Prefix = '" & mPrefix_ & "'"
tmp = tmp & " And TCD.CxTF_DetailsID IN (Select Distinct TCG.CxTF_DetailsID From Txn_CxTF_GSL TCG  Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix Where TCG.TxnType = 'W' And MGSL.NContractID=" & mContractID_ & " And TCG.SM_Prefix='" & mPrefix_ & "')"
tmp = tmp & " And MGSL.SpillageFlag=1"
tmp = tmp & " Group By TCG.SM_Prefix"
tmp = tmp & " Order By TCG.SM_Prefix"


Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   mSpilBags = IIf(IsNull(TmpRs5!NoofBags), 0, TmpRs5!NoofBags)
   mSpilQty = IIf(IsNull(TmpRs5!DematWeight), 0, TmpRs5!DematWeight)
End If

End Sub
Private Sub GetMadeUpOpeningBalance(mContractID_ As Variant, mPrefix_ As Variant)

tmp = " Select TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight'"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL MGSL1 On MGSL1.Mst_GSL_ID = MGSL.Connected_Mst_GSL_ID And MGSL1.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "'"
tmp = tmp & " And TCG.SM_Prefix = '" & mPrefix_ & "'"
tmp = tmp & " And TCD.CxTF_DetailsID IN (Select Distinct CxTF_DetailsID From Txn_CxTF_GSL TCG  Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix Where TCG.TxnType = 'W' And MGSL.NContractID=" & mContractID_ & " And TCG.SM_Prefix='" & mPrefix_ & "')"
''(Select Distinct TINV.CxTFNO From Txn_InvMonGSLDetail TINV Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag in ('W') And TINV.ContractID=" & mContractID_ & " And SM_Prefix='" & mPrefix_ & "')"
tmp = tmp & " And MGSL.MadeupFlag=1 And MGSL1.NContractID=" & mContractID_ & ""
tmp = tmp & " Group By TCG.SM_Prefix"
tmp = tmp & " Order By TCG.SM_Prefix"

Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   mSpilBags = IIf(IsNull(TmpRs5!NoofBags), 0, TmpRs5!NoofBags)
   mSpilQty = IIf(IsNull(TmpRs5!DematWeight), 0, TmpRs5!DematWeight)
End If


End Sub

Private Sub GetMadeUpOpeningPreBalance(mContractID_ As Variant, mPrefix_ As Variant)


tmp = " Select TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight'"

tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL MGSL1 On MGSL1.Mst_GSL_ID = MGSL.Connected_Mst_GSL_ID And MGSL1.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate < '" & mFromDate & "'"
tmp = tmp & " And TCG.SM_Prefix = '" & mPrefix_ & "'"
'tmp = tmp & " And TCD.CxTFNo IN (Select Distinct TINV.CxTFNO From Txn_InvMonGSLDetail TINV Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag in ('W') And TINV.ContractID=" & mContractID_ & " And SM_Prefix='" & mPrefix_ & "')"
'tmp = tmp & " And TCD.CxTFNo IN (Select Distinct CxtfNo From Txn_CxTF_GSL TCG  Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix Where TCG.TxnType = 'W' And MGSL.NContractID=" & mContractID_ & " And TCG.SM_Prefix='" & mPrefix_ & "')"
tmp = tmp & " And MGSL.MadeupFlag=1 And MGSL1.NContractID=" & mContractID_ & ""
tmp = tmp & " Group By TCG.SM_Prefix"
tmp = tmp & " Order By TCG.SM_Prefix"

Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   mSpilBags = IIf(IsNull(TmpRs5!NoofBags), 0, TmpRs5!NoofBags)
   mSpilQty = IIf(IsNull(TmpRs5!DematWeight), 0, TmpRs5!DematWeight)
End If


End Sub


Private Sub GetMadeUpBagsWithdrawal(mPrefix_ As Variant, mCITFNo_ As Variant, mContractID_ As Variant)

tmp = " Select TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight'"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_GSL MGSL1 On MGSL1.Mst_GSL_ID = MGSL.Connected_Mst_GSL_ID And MGSL1.SM_Prefix = MGSL.SM_Prefix"
tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "'"
tmp = tmp & " And TCG.SM_Prefix = '" & mPrefix_ & "'"
tmp = tmp & " And TCD.CxTF_DetailsID = " & mCITFNo_
tmp = tmp & " And MGSL.MadeupFlag=1 And MGSL1.NContractID=" & mContractID_ & ""
tmp = tmp & " Group By TCG.SM_Prefix"
tmp = tmp & " Order By TCG.SM_Prefix"

Call Tmp5Table

If TmpRs5.RecordCount > 0 Then
   mSpilBags = IIf(IsNull(TmpRs5!NoofBags), 0, TmpRs5!NoofBags)
   mSpilQty = IIf(IsNull(TmpRs5!DematWeight), 0, TmpRs5!DematWeight)
End If

End Sub

Private Sub CmdStockExcel_Click()
If MSHFlexGrid2.Rows < 3 Then Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Stock Statement_" & CmbMonth.Text & "_" & CmbYear.Text & ".xls")
Call Xls_Detail_Rpt1
End Sub

Public Sub Xls_Detail_Rpt1()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Stock Statement_" & CmbMonth.Text & "_" & CmbYear.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        If C = 0 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid2.TextMatrix(I, C)
        Else
        
           If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
If mSLocationID <> RsMst_Location!LocationID Then
   mSLocationID = RsMst_Location!LocationID
   CmbMSC.Text = ""
   Call Grid_Head1
End If

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
If mCommodityID <> RsMst_Commodity!CommodityID Then
   mCommodityID = RsMst_Commodity!CommodityID
   CmbMSC.Text = ""
   Call Grid_Head1
End If
End Sub

