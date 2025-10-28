VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_ModifiedNNNRInvoice 
   Caption         =   "Modified NNNR Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   9000
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1440
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   15015
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
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   405
         Width           =   3165
      End
      Begin VB.ComboBox CmbNonNcdexClient 
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
         Left            =   9945
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   405
         Width           =   4980
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
         Left            =   13890
         TabIndex        =   8
         Top             =   1005
         Width           =   975
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
         Height          =   375
         Left            =   12915
         TabIndex        =   7
         Top             =   1005
         Width           =   975
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
         Left            =   6330
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   3540
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
         ItemData        =   "FrmRpt_ModifiedNNNRInvoice.frx":0000
         Left            =   1650
         List            =   "FrmRpt_ModifiedNNNRInvoice.frx":0002
         TabIndex        =   2
         Top             =   405
         Width           =   1455
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
         ItemData        =   "FrmRpt_ModifiedNNNRInvoice.frx":0004
         Left            =   120
         List            =   "FrmRpt_ModifiedNNNRInvoice.frx":0006
         TabIndex        =   1
         Top             =   405
         Width           =   1455
      End
      Begin VB.ComboBox CmbModified 
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
         ItemData        =   "FrmRpt_ModifiedNNNRInvoice.frx":0008
         Left            =   1080
         List            =   "FrmRpt_ModifiedNNNRInvoice.frx":0012
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   930
         Width           =   2220
      End
      Begin VB.Label LblStateID 
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
         Left            =   4170
         TabIndex        =   24
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Left            =   11430
         TabIndex        =   21
         Top             =   135
         Width           =   1170
      End
      Begin VB.Label LblLocationID 
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
         Height          =   240
         Left            =   7680
         TabIndex        =   20
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
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
         Left            =   2130
         TabIndex        =   19
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Month"
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
         Left            =   547
         TabIndex        =   18
         Top             =   165
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Status"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1020
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6645
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   11721
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label7 
         Caption         =   "Total Record Found :"
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
         Left            =   10800
         TabIndex        =   15
         Top             =   240
         Width           =   2235
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
         Left            =   13080
         TabIndex        =   14
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
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
         Left            =   360
         TabIndex        =   13
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Deposit and Withdrawal of lot done in same month so can't generate invoice for this withdrawal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   4605
      End
      Begin VB.Label Label8 
         Caption         =   "Storage Rate not availabe in Storage Rate Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6645
         TabIndex        =   11
         Top             =   270
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
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
         TabIndex        =   10
         Top             =   375
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmRpt_ModifiedNNNRInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mSM_Prefix As Variant
Dim mFromDate, mToDate, vMonth, mBillingMonth, mAmount, mSPRateFlag As Variant
Dim mMonth, mTotalDays, mBillingCycleID, mTxtTotDays, mFlag As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mTFlagOnUpdate, mTxnNo As Variant
Dim mStateID, mSMPrefixFromState As Variant
 
Private Sub CmbModified_LostFocus()
If CmbModified.Text = "" Then Exit Sub

If LCase(CmbModified.Text) <> "unchanged data" And LCase(CmbModified.Text) <> "changed data" Then
   MsgBox "Invalid Selection!!!"
   CmbModified.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I), I - 1
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
'Call ClearFrame
'Frame4.Enabled = False
Call Grid_Head
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
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
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
'Call ClearFrame
Call Grid_Head
'Frame4.Enabled = False
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
Private Sub CmdGo_Click()

'On Error GoTo mErrorMsg

Dim m As Double
Dim mR As Integer
ProgressBar1.Value = 0
'Call ClearFrame
Call Grid_Head
 
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
mSMPrefixFromState = ""
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "fag" And LCase(Gen_UserRole) <> "fagadmin" Then
   If CmbStateID.Text = "" Then
      MsgBox "Blank State and Location not allowed!!!"
      CmbStateID.SetFocus
      Exit Sub
   End If
   
   If CmbLocationID.Text = "" Then
      MsgBox "Blank Location not allowed!!!"
      CmbLocationID.SetFocus
      Exit Sub
   End If
Else
   If CmbStateID.Text <> "" And CmbLocationID.Text = "" Then
      tmp = "Select Distinct SM_Prefix from Mst_Location Where StateID = " & mStateID & ""
      Call TmpTable
      If TmpRs.RecordCount > 0 Then
         mSMPrefixFromState = " SM_Prefix in ('"
         For x = 1 To TmpRs.RecordCount
             If x = TmpRs.RecordCount Then
                mSMPrefixFromState = mSMPrefixFromState & TmpRs!SM_Prefix & "')"
             Else
                mSMPrefixFromState = mSMPrefixFromState & TmpRs!SM_Prefix & "','"
             End If
             TmpRs.MoveNext
         Next
      End If
   End If

End If





mMonth = CmbMonth.ListIndex + 1
vMonth = Month(Now)

'If mMonth > vMonth Then
'   MsgBox "Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If

'If CmbLocationID.Text = "" Then
'   MsgBox "Please Select Location"
'   CmbLocationID.SetFocus
'   Exit Sub
'End If


If mMonth < 10 Then
   mMonth = "0" & mMonth
End If

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

 


'tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.Txn_TempGenInvoice') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
'tmp = tmp & " drop table dbo.Txn_TempGenInvoice"
'Call TmpTable


Call TableTxn_CreateTempGenInvoice
Call TableTxn_TempGenInvoice


'If Gen_DBType = "MDB" Then
'   tmp = "Select * from Txn_InvMonGSLDetail Where SM_Prefix = '" & mSM_Prefix & "' And  YearMonth = #" & mFromDate & "# And TFLag  <> 'O'"
'Else
'   tmp = "Select * from Txn_InvMonGSLDetail Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & mFromDate & "' And TFLag  <> 'O'"
'End If
'
'Call TmpTable
'
'If TmpRs.RecordCount = 0 Then
'   MsgBox "Monthly transaction process for " & CmbLocationID.Text & " is not executed yet!!! "
'   Exit Sub
'End If
LblTotalRecord.Caption = ""
MsgBox "Wait till next message !!!"

tmp = " Alter table Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""
tmp = tmp & " add  OldRate Decimal(18,4),"
tmp = tmp & " OldBillingCycleID decimal(18,0),"
tmp = tmp & " OldBillingUnit Varchar(1),"
tmp = tmp & " UpdatedBy varchar(30),"
tmp = tmp & " UpdatedDate datetime"
Call Tmp1Table


tmp = " Insert Into Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & "(Mst_GSL_ID,SM_Prefix ,ExchangeTypeID ,CMPID ,ClientID,CommodityID,BalanceBags,BalanceWeight ,Rate,InvoiceAmount ,"
tmp = tmp & " TFlag,BillingCycleID,BillingUnit,NoOfDays,InvoiceDate,PINo,Bagsize,ContractID,CxTFDate,CDTFDate,YearMonth,CxTF_DetailsID,OldRate,OldBillingCycleID,OldBillingUnit,UpdatedBy,UpdatedDate)"
tmp = tmp & " Select Mst_GSL_ID,SM_Prefix ,ExchangeTypeID ,CMPID ,ClientID,CommodityID,BalanceBags,BalanceWeight ,Rate,InvoiceAmount ,"
tmp = tmp & " TFlag , BillingCycleID, BillingUnit, NoOfDays, InvoiceDate, PINo, Bagsize, ContractID, CxTFDate, CDTFDate, YearMonth,CxTF_DetailsID,OldRate,OldBillingCycleID,OldBillingUnit "
tmp = tmp & " ,UpdatedBy,UpdatedDate"
tmp = tmp & " From Txn_InvMonGslDetail"


If Gen_DBType = "MDB" Then
   tmp = tmp & " Where YearMonth = #" & mFromDate & "# And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
Else
   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
End If


If CmbLocationID.Text <> "" Then
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix & "'"
Else
   If mSMPrefixFromState <> "" Then
      tmp = tmp & " And " & mSMPrefixFromState & ""
   End If
End If

tmp = tmp & " union all "

tmp = tmp & " Select Mst_GSL_ID,SM_Prefix ,ExchangeTypeID ,CMPID ,ClientID,CommodityID,BalanceBags,BalanceWeight ,Rate,InvoiceAmount ,"
tmp = tmp & " TFlag , BillingCycleID, BillingUnit, NoOfDays, InvoiceDate, PINo, Bagsize, ContractID, CxTFDate, CDTFDate, YearMonth,CxTF_DetailsID,OldRate,OldBillingCycleID,OldBillingUnit "
tmp = tmp & " ,UpdatedBy,UpdatedDate"
tmp = tmp & " From Txn_InvMonGSLDetail2014to2015 "

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where YearMonth = #" & mFromDate & "# And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
Else
   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
End If

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix & "'"
Else
   If mSMPrefixFromState <> "" Then
      tmp = tmp & " And " & mSMPrefixFromState & ""
   End If
End If

tmp = tmp & " union all "

tmp = tmp & " Select Mst_GSL_ID,SM_Prefix ,ExchangeTypeID ,CMPID ,ClientID,CommodityID,BalanceBags,BalanceWeight ,Rate,InvoiceAmount ,"
tmp = tmp & " TFlag , BillingCycleID, BillingUnit, NoOfDays, InvoiceDate, PINo, Bagsize, ContractID, CxTFDate, CDTFDate, YearMonth,CxTF_DetailsID,OldRate,OldBillingCycleID,OldBillingUnit "
tmp = tmp & " ,UpdatedBy,UpdatedDate"
tmp = tmp & " From Txn_InvMonGSLDetailUpto2013 "

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where YearMonth = #" & mFromDate & "# And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
Else
   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
End If

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix & "'"
Else
   If mSMPrefixFromState <> "" Then
      tmp = tmp & " And " & mSMPrefixFromState & ""
   End If
End If

'tmp = tmp & " union all "
'
'tmp = tmp & " Select GSLID,SM_Prefix ,ExchangeTypeID ,CMPID ,ClientIDRow,CommodityID,BalanceBags,BalanceWeight ,Rate,InvoiceAmount ,"
'tmp = tmp & " TFlag , BillingCycleID, BillingUnit, NoOfDays, InvoiceDate, PINo, Bagsize, ContractID, CxTFDate, CDTFDate, YearMonth,CxTFNo,OldRate,OldBillingCycleID,OldBillingUnit "
'tmp = tmp & " ,UpdatedBy,UpdatedDate"
'tmp = tmp & " From Txn_InvMonGSLDetail2014to2015 "
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Where YearMonth = #" & mFromDate & "# And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
'Else
'   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
'End If
'
'If CmbLocationID.Text <> "" Then
'   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix & "'"
'Else
'   If mSMPrefixFromState <> "" Then
'      tmp = tmp & " And " & mSMPrefixFromState & ""
'   End If
'End If


Call TmpTable


tmp = " Select ML.LocationName ,MCMP.CMPName,MClient.ClientName,MCG.CommodityGroup, Me.ExchangeType ,"
tmp = tmp & " TINV.TFlag,TINV.YearMonth,TINV.CxTFDate ,TINV.CDTFDate,"
tmp = tmp & " Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight',"
tmp = tmp & " TINV.BillingCycleID,MBC1.BillingCycle 'OldBillingCycle',TINV.BillingUnit,TINV.Rate,TINV.OldBillingCycleID,TINV.OldBillingUnit,TINV.OldRate,"
tmp = tmp & " TInv.NoofDays , TInv.Bagsize, Sum(TInv.InvoiceAmount) 'InvoiceAmount',  MCMP.LocationID , TINV.CMPID, TINV.ClientID,"
tmp = tmp & " TINV.CommodityID, MC.CommodityGroupID, TINV.ExchangeTypeID,TINV.Mst_GSL_ID,MC.Commodity,MST.NAVCode,MBC.BillingCycle,"
tmp = tmp & " TInv.Rate , MST.DEscription,TInv.Bagsize, TInv.BillingCycleID, TInv.BillingUnit, TInv.CxTF_DetailsID, TInv.OldRate,TInv.OldBillingCycleID,TInv.OldBillingUnit,"
tmp = tmp & " TINV.UpdatedBy,TINV.UpdatedDate,TINV.PINo,TID.TINo,TID.TIDate,TID.PostInvoiceNo,TID.TIAMount,TID.TISTAmount "
tmp = tmp & " From Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & " TINV"
tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MClient On TINV.ClientID = MClient.ClientID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_BillingCycle MBC On TINV.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Left Join Mst_BillingCycle MBC1 On TINV.OldBillingCycleID = MBC1.BillingCycleID"
tmp = tmp & " Left Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"
tmp = tmp & " Left Join Txn_TIDetails TID On TID.PIFLag = 'NNNR' And TID.PIID = TINV.PINo "
tmp = tmp & " Where (TINV.ContractID = 0 Or TINV.ContractID Is Null)  " ' And TCG.BillFlag Is Null"
tmp = tmp & " and TINV.ExchangeTypeID = 1 And TINV.YearMonth='" & mFromDate & "'"

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And TINV.SM_Prefix = '" & mSM_Prefix & "'"
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
   If mSMPrefixFromState <> "" Then
      tmp = tmp & " And TINV." & mSMPrefixFromState & ""
   End If
End If

If CmbNonNcdexClient.Text <> "" Then
   tmp = tmp & " And MClient.ClientName = '" & CmbNonNcdexClient & "'"
End If

If LCase(CmbModified.Text) = "changed data" Then
   tmp = tmp & " and TINV.UpdatedBy is not null"
ElseIf LCase(CmbModified.Text) = "unchanged data" Then
   tmp = tmp & " and TINV.UpdatedBy is null"
End If
 
tmp = tmp & " Group by ML.LocationName , TInv.SM_PreFix, MCMP.CMPName, MClient.ClientName, MCG.CommodityGroup, Me.ExchangeType,"
tmp = tmp & " tINV.Bagsize,TINV.YearMonth,TINV.OldBillingCycleID,TINV.OldBillingUnit,TINV.OldRate, TINV.CxTFDate ,TINV.TFlag,"
tmp = tmp & " MCMP.LocationID ,MST.Description, TINV.CMPID, TINV.ClientID, TINV.CommodityID, MC.CommodityGroupID, TINV.ExchangeTypeID,TINV.Mst_GSL_ID,"
tmp = tmp & " TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.NoofDays,TINV.Rate, TInv.Bagsize ,TINV.BillingCycleID,"
tmp = tmp & " TINV.BillingUnit,TINV.CxTF_DetailsID,MBC1.BillingCycle, "
tmp = tmp & " TINV.UpdatedBy,TINV.UpdatedDate,TINV.PINo,TID.TINo,TID.TIDate,TID.PostInvoiceNo,TID.TIAmount,TID.TISTAmount"
tmp = tmp & " Order By ML.LocationName,MClient.ClientName,MCG.CommodityGroup "
 
Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   With MSHFlexGrid1
   For mR = 1 To TmpRs.RecordCount
       .TextMatrix(mR, 1) = TmpRs!Locationname
       .TextMatrix(mR, 2) = TmpRs!CMPName
       .TextMatrix(mR, 3) = TmpRs!ClientName
       .TextMatrix(mR, 4) = TmpRs!COmmodityGroup
       .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
       .TextMatrix(mR, 6) = TmpRs!tFlag
       .TextMatrix(mR, 7) = Format(IIf(IsNull(TmpRs!YearMonth), "", TmpRs!YearMonth), "dd-MMM-yyyy")
       .TextMatrix(mR, 8) = Format(IIf(IsNull(TmpRs!CxTFDate), "", TmpRs!CxTFDate), "dd-MMM-yyyy")
       .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!CDTFDate), "", TmpRs!CDTFDate)
       .TextMatrix(mR, 9) = Format(.TextMatrix(mR, 9), "dd-mmm-yyyy")
       
       .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
       .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
       ' Changes From below in structure
       .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle) '"Billing Cycle"
       .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!BillingUnit), "", TmpRs!BillingUnit) '"No of Cycles"
       .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!Rate), "", TmpRs!Rate) '"Rate"
       .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!OldBillingCycle), "", TmpRs!OldBillingCycle) '"Billing CycOldBillingCycleIDle"
       .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!OldBillingUnit), "", TmpRs!OldBillingUnit) '"No of Cycles"
       .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!OldRate), "", TmpRs!OldRate) '"Rate"
       'OldRate , OldBillingCycleID, OldBillingUnit
       .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!NoOfDays), "", TmpRs!NoOfDays)
       '.TextMatrix(mr, 19) = Format(IIf(IsNull(TmpRs!InvoiceAmount), 0, TmpRs!InvoiceAmount), "########0.0000") '"Total Amount"
       .TextMatrix(mR, 21) = TmpRs!Mst_GSL_ID
       .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!BagSize), "", TmpRs!BagSize)
       .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!InvoiceAmount), "", TmpRs!InvoiceAmount)
       '.TextMatrix(mr, 23) = IIf(IsNull(TmpRs!GSLID), "", TmpRs!GSLID)
       .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!Description), "", TmpRs!Description)
       '.TextMatrix(mr, 24) = IIf(IsNull(TmpRs!BillingUnit), "", TmpRs!BillingUnit)
       .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID)
       .TextMatrix(mR, 24) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy)
       .TextMatrix(mR, 25) = IIf(IsNull(TmpRs!UpdatedDate), "", TmpRs!UpdatedDate)
       .TextMatrix(mR, 25) = Format(.TextMatrix(mR, 25), "dd-mmm-yyyy HH:MM:SS")
       
       .TextMatrix(mR, 26) = IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)
       .TextMatrix(mR, 27) = IIf(IsNull(TmpRs!TINo), "", TmpRs!TINo)
       .TextMatrix(mR, 28) = IIf(IsNull(TmpRs!TIDate), "", TmpRs!TIDate)
       .TextMatrix(mR, 29) = Format(.TextMatrix(mR, 27), "dd-mmm-yyyy")
       .TextMatrix(mR, 29) = IIf(IsNull(TmpRs!PostInvoiceNo), "", TmpRs!PostInvoiceNo)
       .TextMatrix(mR, 30) = IIf(IsNull(TmpRs!TIAmount), "", TmpRs!TIAmount)
       .TextMatrix(mR, 31) = IIf(IsNull(TmpRs!TISTAmount), "", TmpRs!TISTAmount)
       
       If LCase(.TextMatrix(mR, 13)) = "b" Then
          .TextMatrix(mR, 13) = "Bags"
       ElseIf LCase(.TextMatrix(mR, 13)) = "q" Then
          .TextMatrix(mR, 13) = "Quantity"
       ElseIf LCase(.TextMatrix(mR, 13)) = "s" Then
          .TextMatrix(mR, 13) = "Sq Ft"
       End If
       
       If LCase(.TextMatrix(mR, 16)) = "b" Then
          .TextMatrix(mR, 16) = "Bags"
       ElseIf LCase(.TextMatrix(mR, 16)) = "q" Then
          .TextMatrix(mR, 16) = "Quantity"
       ElseIf LCase(.TextMatrix(mR, 16)) = "s" Then
          .TextMatrix(mR, 16) = "Sq Ft"
       End If
       
       'if Rate = 0 then mark as red
         If (Val(.TextMatrix(mR, 14)) = 0) Then
            For I = 0 To MSHFlexGrid1.Cols - 1
                .Col = I
                 .row = mR
                .CellBackColor = vbRed
            Next
         End If
         
       'if D and W in same month then mark as green
         If .TextMatrix(mR, 6) = "W" Then
            If Month(.TextMatrix(mR, 9)) & "~" & Year(.TextMatrix(mR, 9)) = Month(.TextMatrix(mR, 8)) & "~" & Year(.TextMatrix(mR, 8)) Then
               For I = 0 To MSHFlexGrid1.Cols - 1
                   .Col = I
                   .row = mR
                   .CellBackColor = vbGreen
               Next
            End If
         End If
        
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
End If
ProgressBar1.Value = ProgressBar1.Max

LblTotalRecord.Refresh
LblTotalRecord.Caption = TmpRs.RecordCount

tmp = "Drop Table Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""
Call TmpTable

MsgBox "Done !"
'CmdSave.Enabled = True

Exit Sub

mErrorMsg:

tmp = "Drop Table Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""
Call TmpTable

MsgBox "Some error occured while executing process!!!"


End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 32 '30 '26 '26 '27 '23 '25
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(0) = 800

     .TextMatrix(0, 0) = "Select"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "CMP Name"
     .TextMatrix(0, 3) = "Client Name"
     .TextMatrix(0, 4) = "Commodity"
     .TextMatrix(0, 5) = "Exchange Type"
     .TextMatrix(0, 6) = "Txn Type"
     .TextMatrix(0, 7) = "Billing Month" '"Year Month"
     .TextMatrix(0, 8) = "Txn Date" '"CxTF Date"
     .TextMatrix(0, 9) = "CDTF Date" '.TextMatrix(0, 19) = "CDTF Date"
     .TextMatrix(0, 10) = "No Of Bags" '"Balance Bags"
     .TextMatrix(0, 11) = "Weight in MT" '"Balance Weight"
     .TextMatrix(0, 12) = "Present Billing Cycle"
     .TextMatrix(0, 13) = "Present Billing Unit"
     .TextMatrix(0, 14) = "Present Billing Rate"
     .TextMatrix(0, 15) = "Old Billing Cycle"
     .TextMatrix(0, 16) = "Old Billing Unit"
     .TextMatrix(0, 17) = "Old Billing Rate"
     .TextMatrix(0, 18) = "No of Cycles"
     '.TextMatrix(0, 19) = "Rate"
     .TextMatrix(0, 19) = "Bagsize"
     .TextMatrix(0, 20) = "Total Amount"
     .TextMatrix(0, 21) = "GSLID"
     .TextMatrix(0, 22) = "ST NAV Code"
     '.TextMatrix(0, 24) = "Billing Unit"
     .TextMatrix(0, 23) = "CxTFNo"
     .TextMatrix(0, 24) = "Updated By"
     .TextMatrix(0, 25) = "Updated Date"
          
     .TextMatrix(0, 26) = "PI No"
     .TextMatrix(0, 27) = "TI No (Voucher No)"
     .TextMatrix(0, 28) = "TI Date"
     .TextMatrix(0, 29) = "PostInvoice No"
     .TextMatrix(0, 30) = "TI Amount"
     .TextMatrix(0, 31) = "TI ST AMount"
          
     .ColWidth(0) = 0
     .ColWidth(1) = 2100
     .ColWidth(2) = 2100
     .ColWidth(3) = 3500
     .ColWidth(4) = 1800
     .ColWidth(5) = 1600
     .ColWidth(6) = 800 '1200
     .ColWidth(7) = 1500
     .ColWidth(8) = 1500 '800
     .ColWidth(9) = 1300
     .ColWidth(10) = 1300
     .ColWidth(11) = 1200
     .ColWidth(12) = 800
     .ColWidth(13) = 800
     .ColWidth(14) = 1200
     .ColWidth(15) = 800 '1200
     .ColWidth(17) = 800
     .ColWidth(18) = 800
     .ColWidth(19) = 500
     .ColWidth(20) = 500
     .ColWidth(21) = 500
     .ColWidth(22) = 0
     .ColWidth(23) = 0
     .ColWidth(24) = 1000
     .ColWidth(25) = 1600
     
      
End With
End Sub

Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Select state!!!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID = " & mStateID & " ")
Else
   Call TableMst_Location("where StateID = " & mStateID & "  " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
    
   Call Grid_Head
'   Frame4.Enabled = False
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   Call Grid_Head
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub

 
 
 
Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
'Frame4.Enabled = False

Call Grid_Head

'For I = 1 To 12
'    CmbMonth.AddItem MonthName(I), I - 1
'Next
CmdGo.Enabled = True
CmdExcel.Enabled = True
'CmdSave.Enabled = False
ProgressBar1.Value = 0
End Sub
 
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.xls")
'Call Xls_Detail_Rpt


tmp = ""
ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If ans = vbNo Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.xls")
   Call Xls_Detail_Rpt
Else
   Path = App.Path
   strFilName = "\Excel\" & Gen_mTimeStamp & "Rpt_InvoiceDetails" & Gen_UserName & ".csv"
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        MSHFlexGrid1.row = I
        MSHFlexGrid1.Col = C
                
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbYellow
        End If
              
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
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
 
 
Private Sub CmbStateID_GotFocus()

tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where " & Gen_WcState)
Else
   Call TableMst_State '("Where RegionID = " & mRegionID & "")
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


