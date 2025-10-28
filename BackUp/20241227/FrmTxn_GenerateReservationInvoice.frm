VERSION 5.00
Begin VB.Form FrmTxn_GenerateReservationInvoice 
   Caption         =   "Generate Reservation Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ProgressBar1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   14970
      TabIndex        =   20
      Top             =   9480
      Width           =   15030
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   15015
      Begin VB.TextBox Text3 
         BackColor       =   &H00FF0000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   25
         Top             =   1230
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   495
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
         Left            =   7125
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   420
         Width           =   5745
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
         Left            =   3585
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   435
         Width           =   3495
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
         Left            =   60
         TabIndex        =   12
         Top             =   435
         Width           =   1935
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
         Left            =   2070
         TabIndex        =   11
         Top             =   435
         Width           =   1455
      End
      Begin VB.CommandButton CmdShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   12945
         TabIndex        =   10
         Top             =   300
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7095
         TabIndex        =   9
         Top             =   885
         Width           =   495
      End
      Begin VB.CommandButton CmdShowAll 
         Caption         =   "Show All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   13920
         TabIndex        =   8
         Top             =   300
         Width           =   975
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
         Height          =   450
         Left            =   12945
         TabIndex        =   7
         Top             =   765
         Width           =   975
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
         Height          =   450
         Left            =   13935
         TabIndex        =   6
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Billing Details not updated in Storage Contract master"
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
         Left            =   7740
         TabIndex        =   26
         Top             =   1245
         Width           =   5055
      End
      Begin VB.Label Label6 
         Caption         =   "BOD not updated for any one of CMP in selected Location"
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
         Left            =   750
         TabIndex        =   22
         Top             =   975
         Width           =   6015
      End
      Begin VB.Label Label5 
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
         Left            =   9412
         TabIndex        =   19
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   4612
         TabIndex        =   17
         Top             =   180
         Width           =   1440
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
         Left            =   667
         TabIndex        =   16
         Top             =   180
         Width           =   720
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
         Left            =   2550
         TabIndex        =   15
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Reservation Billing for previous month not generated"
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
         Left            =   7725
         TabIndex        =   14
         Top             =   900
         Width           =   5055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   1605
      Width           =   15015
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
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
         TabIndex        =   1
         Top             =   195
         Width           =   1440
      End
      Begin VB.PictureBox MSHFlexGrid1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4035
         ScaleWidth      =   14715
         TabIndex        =   2
         Top             =   480
         Width           =   14775
      End
      Begin VB.PictureBox MSHFlexGrid2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2715
         ScaleWidth      =   14715
         TabIndex        =   23
         Top             =   4680
         Visible         =   0   'False
         Width           =   14775
      End
      Begin VB.PictureBox MSHFlexGrid3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   105
         ScaleHeight     =   2955
         ScaleWidth      =   14715
         TabIndex        =   24
         Top             =   4650
         Width           =   14775
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
         Left            =   9120
         TabIndex        =   4
         Top             =   120
         Width           =   2295
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
         Left            =   11520
         TabIndex        =   3
         Top             =   120
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmTxn_GenerateReservationInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mEmployeeID, mEmployeeName As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mUserTypeFlag As Variant
Dim mSelRow As Variant
Dim mSelCol As Variant
Dim mMonth, mFromDate, mToDate, mPreviousMonth, mTotalDays, mRow, mMonthListindex As Variant
Dim mMsg, mGridMsg As String
Dim oSMTP As SMTPSession
Dim mYear As Variant
Dim mClientIDRow, LocationID, mSM_Prefix As Variant
Dim mChooseLocation  As String
Dim mSpilBags, mSpilQty As Variant
Dim mMidFromDate As Variant
Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Refresh
    .Rows = 2
    .Cols = 19
    .WordWrap = True
    
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "Contract ID"
    .TextMatrix(0, 2) = "Contract Name"
    .TextMatrix(0, 3) = "Location Name"
    .TextMatrix(0, 4) = "Client Name"
    .TextMatrix(0, 5) = "Commodity"
    .TextMatrix(0, 6) = "Contract Date"
    .TextMatrix(0, 7) = "Booked From"
    .TextMatrix(0, 8) = "Booked To"
    .TextMatrix(0, 9) = "Extended Upto"
    .TextMatrix(0, 10) = "Booked As Per"
    .TextMatrix(0, 11) = "SM_Prefix" '"Booked Rate"
    .TextMatrix(0, 12) = "Booked Quantity"
    .TextMatrix(0, 13) = "Peak From"
    .TextMatrix(0, 14) = "Remarks"
    .TextMatrix(0, 15) = "Location ID"
    .TextMatrix(0, 16) = "Service Tax NAV Code"
    .TextMatrix(0, 17) = "Adhoc Invoice Amount (Only for information) "
    .TextMatrix(0, 18) = "Master ContractID"
    
    .ColWidth(0) = 700
    .ColWidth(1) = 1000
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1600
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 0
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 1000
    .ColWidth(14) = 2000
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 2000
    .RowHeight(0) = 800
    
End With
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()

CmdSave.Enabled = False
ProgressBar1.Value = 0

If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No data to Save!!!"
   Exit Sub
End If

If ChkSelection = False Then
   MsgBox "No data to selected for Saving!!!"
   Exit Sub
End If

Call TableTxn_ReservationInvoice(" Where YearMonth='" & mFromDate & "'")
Call TableTxn_ReservationInvoiceDetail

''tmp = "Select ML.LocationName,TInv.ContractID, TInv.SM_PreFix,Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight'"
''tmp = tmp & " From Txn_InvMonGSLDetail TINV"
''tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
''tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
''tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
''tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
''tmp = tmp & " Inner join Mst_StorageContract MSC on MSC.ContractID=TINV.ContractID"
''tmp = tmp & " Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag='O'"
''tmp = tmp & " Group by ML.LocationName, TInv.ContractID, TInv.SM_PreFix"

tmp = "Select ML.LocationID, ML.LocationName, TINV.ContractID,Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight',TINV.YearMonth"
tmp = tmp & " From Txn_InvMonGSLDetail TINV"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag in ('O','P') And TINV.ContractID>0 And TINV.MadeupFlag=0"
tmp = tmp & " Group by ML.LocationID, ML.LocationName, TINV.ContractID,TINV.YearMonth"

Call Tmp2Table

''tmp = "Select ML.LocationID, ML.LocationName, TINV.ContractID,Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight',TINV.YearMonth"
''tmp = tmp & " From Txn_InvMonGSLDetail TINV"
''tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
''tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
''tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
''tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
''tmp = tmp & " Where TINV.YearMonth = '" & mPrevMonthDate & "' And TINV.TFlag='O' And TINV.ContractID>0"
''tmp = tmp & " And GSLID in (select GSLID From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.SM_Prefix=TINV.SM_Prefix)"
''tmp = tmp & " Group by ML.LocationID, ML.LocationName, TINV.ContractID,TINV.YearMonth"


tmp = "Select ML.LocationID, ML.LocationName, TINV.ContractID,Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight',TINV.YearMonth,TINV.CxTFDate ,TINV.TFlag,TINV.CxTF_DetailsID"
tmp = tmp & " From Txn_InvMonGSLDetail TINV"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " left Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Where TINV.YearMonth = '" & mFromDate & "' And TINV.TFlag in ('D','W') And TINV.ContractID>0"
tmp = tmp & " Group by ML.LocationID, ML.LocationName, TINV.ContractID,TINV.CxTFDate ,TINV.TFlag,TINV.YearMonth,TINV.CxTF_DetailsID"

Call Tmp3Table

ProgressBar1.Max = MSHFlexGrid1.Rows + 5


For I = 1 To MSHFlexGrid1.Rows
    If MSHFlexGrid1.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       Call FillDaywiseBalance(I)
'       RsTxn_ReservationInvoice.AddNew
'       RsTxn_ReservationInvoice!InvoiceNo = GetMaxInvoiceNo
'       RsTxn_ReservationInvoice!ContractID = MSHFlexGrid1.TextMatrix(I, 1)
'       RsTxn_ReservationInvoice!LocationID = MSHFlexGrid1.TextMatrix(I, 15)
'       RsTxn_ReservationInvoice!ServiceTaxNAVCode = MSHFlexGrid1.TextMatrix(I, 16)
'       RsTxn_ReservationInvoice!YearMonth = CDate(mFromDate)
'       RsTxn_ReservationInvoice!BillingDate = CDate(mToDate)
'       RsTxn_ReservationInvoice!BillingAmount = Val(MSHFlexGrid1.TextMatrix(I, 13))
''       RsTxn_ReservationInvoice!ForwardToID = GetEmployeeIDFromEmailid("sandeep.m@ncml.com")
'       RsTxn_ReservationInvoice!G_UID = GetGUID
'       RsTxn_ReservationInvoice!Flag = "P"
'       RsTxn_ReservationInvoice!CreatedBy = Gen_UserLoginID
'       RsTxn_ReservationInvoice!CreatedDate = Now
'       RsTxn_ReservationInvoice.Update
'       RsTxn_ReservationInvoice.Requery
    
    End If
    
    
    '5   GST with Tax    GSTS18
    '6   GST without Tax GSTS00
    
    
    tmp = "Select a.InvoiceDate,a.LocationID,a.ClientID,a.ContractID,st.NAVCode,Sum(a.Amount) 'Amount'"
    tmp = tmp & " from Txn_TempInvoice a"
    tmp = tmp & " Inner Join Mst_ServiceTax st on a.ServiceTaxID = st.ServiceTaxID"
    tmp = tmp & " Where a.ContractID = " & Val(MSHFlexGrid1.TextMatrix(I, 1)) & " and a.InvoiceDate = '" & CDate(mFromDate) & "' And Isnull(a.ReverseFlag,'N') = 'N' And a.Flag = 'A' And Isnull(a.BatchID,0) <> 0 And a.IETypeID = 4"
    tmp = tmp & " Group  By a.InvoiceDate,a.LOcationID,a.ClientID,a.ContractID,st.NAVCode"
    
    Call TmpTable
    
    For xy = 1 To TmpRs.RecordCount
        mAdhocAmt = TmpRs!Amount
        tmp = "Select * from Txn_ReservationInvoice Where YearMonth = '" & CDate(mFromDate) & "' And ContractID = " & TmpRs!ContractID & " And LocationID = " & TmpRs!LocationID & " And ServiceTaxNAVCode = '" & TmpRs!NavCode & "'"
        Call Tmp1Table
         
        For pqr = 1 To TmpRs1.RecordCount
            If TmpRs1!BillingAmount >= Val(mAdhocAmt) Then
               TmpRs1!Discount = Val(TmpRs1!Discount) + Val(mAdhocAmt)
               TmpRs1!ActualAmount = Val(TmpRs1!BillingAmount) - Val(TmpRs1!Discount)
               mAdhocAmt = 0
            Else
               TmpRs1!Discount = Val(TmpRs1!Discount) + Val(TmpRs1!BillingAmount)
               TmpRs1!ActualAmount = Val(TmpRs1!BillingAmount) - Val(TmpRs1!Discount)
               mAdhocAmt = Val(mAdhocAmt) - Val(TmpRs1!BillingAmount)
            End If
            TmpRs1.Update
            If mAdhocAmt = 0 Then Exit For
            TmpRs1.MoveNext
        Next pqr
        TmpRs.MoveNext
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!"
LblTotalRecord.Caption = ""

Call Grid_Head

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

Private Sub CmdShow_Click()

Call Grid_Head
Call Grid_Head2
LblTotalRecord.Caption = ""
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

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If mMonthListindex = 0 Then
   mPreviousMonth = Format(mYear - 1 & "-" & 12 & "-" & "01", "yyyy-mm-dd")
Else
   mPreviousMonth = Format(CmbYear.Text & "-" & mMonthListindex & "-" & "01", "yyyy-mm-dd")
End If

'mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & "30", "yyyy-mm-dd")

If CDate(Format(mFromDate, "yyyymm")) < "201107" Then
   MsgBox "Invoicing for months previous to july not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

'If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
'   MsgBox "Present and Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If

wc = ""
'mGenContractID = ""

tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
'Commented on 06 June 2014
'tmp = tmp & " Where TIG.ExchangeTypeID=1  And TIG.YearMonth = '" & mFromDate & "'"
tmp = tmp & " Where (TIG.ExchangeTypeID=1 Or TIG.STackNo = 'Dummy-P')  And TIG.YearMonth = '" & mFromDate & "'"

Call TmpTable
mChooseLocation = ""
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       If mChooseLocation = "" Then
          mChooseLocation = "(" & "'" & TmpRs!SM_Prefix & "'"
       Else
          mChooseLocation = mChooseLocation & "," & "'" & TmpRs!SM_Prefix & "'"
       End If
       
       If I = TmpRs.RecordCount Then
          mChooseLocation = mChooseLocation & ")"
       End If
       TmpRs.MoveNext
   Next
Else
   MsgBox " Opening balance not processed for selected month!!"
   Exit Sub
End If

If Gen_DBType = "MDB" Then
   wc = " Where MSC.BookedFrom <= #" & mToDate & "# And  MSC.ExtendedPeriod >= #" & mFromDate & "#"
   wc = wc & " And ContractID not in (Select ContractID from Txn_ReservationInvoice Where YearMonth = #" & mFromDate & "#) "
Else
   wc = " Where MSC.BookedFrom <= '" & mToDate & "' And  MSC.ExtendedPeriod >= '" & mFromDate & "'"
   wc = wc & " And ContractID not in (Select ContractID from Txn_ReservationInvoice Where YearMonth = '" & mFromDate & "') "
End If
 
'wc = wc & " And ContractID not in (Select ContractID from Txn_ReservationInvoice Where YearMonth = '" & mFromDate & "') "
 
If CmbLocationID.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " And ML." & Gen_WcLocation
   End If
End If

If CmbNonNcdexClient.Text <> "" Then
   wc = wc & " And MC.ClientName = '" & CmbNonNcdexClient.Text & "'"
End If

wc = wc & " And MSC.Flag='A' And ML.SM_Prefix in " & mChooseLocation

tmp = "Select MSC.ContractID,MSC.ContractName,MC.ClientName,ML.LocationName,MComm.Commodity, "
tmp = tmp & " MSC.ContractDate,MSC.BookedFrom,MSC.BookedTo,MSC.ExtendedPeriod,MSC.BookedAsPer, "
tmp = tmp & " ML.SM_Prefix,MSC.PeakFrom,'' BookedQty,MSC.Remarks,MSC.LocationID,MST.NAVCode,MC.ClientID,MSC.MasterStorageContractID,MSC.SCMFlag,Isnull(MSC.StockManagmentContract,0) 'StockManagmentContract' "
tmp = tmp & " From Mst_StorageContract MSC"
tmp = tmp & " Inner Join Mst_Client MC on MSC.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_Location ML On MSC.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_Commodity MComm On MSC.CommodityID = MComm.CommodityID"
tmp = tmp & " Left Join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "

tmp1 = " Order By ML.LocationName,MSC.MasterStorageContractID,MSC.ContractID"

tmp = tmp & wc & tmp1


Call TmpTable

tmp = "Select Distinct MC.ClientID "
tmp = tmp & " From Mst_StorageContract MSC"
tmp = tmp & " Inner Join Mst_Client MC on MSC.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_Location ML On MSC.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_Commodity MComm On MSC.CommodityID = MComm.CommodityID"
tmp = tmp & " Left Join Mst_ServiceTax MST On MComm.ServiceTaxID = MST.ServiceTaxID "
tmp = tmp & wc

Wc1 = tmp

tmp = "Select FromDate,ToDate,BillingCycleID,BillingQty,BillingRate,ContractID from Mst_SContractBillingDetail"
tmp = tmp & " Where FromDate<='" & mToDate & "' And ToDate>='" & mFromDate & "'"

Call Tmp5Table

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 1

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!ContractID
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!ContractName
       MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!Locationname
       MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!ClientName
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(I, 6) = Format(TmpRs!ContractDate, "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 7) = Format(TmpRs!BookedFrom, "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 8) = Format(TmpRs!BookedTo, "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 9) = Format(TmpRs!ExtendedPeriod, "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 10) = TmpRs!BookedAsPer
       MSHFlexGrid1.TextMatrix(I, 11) = TmpRs!SM_Prefix
       'MSHFlexGrid1.TextMatrix(I, 12) = TmpRs!BookedQty
       If TmpRs!PeakFrom = "C" Then
          MSHFlexGrid1.TextMatrix(I, 13) = "Closing Balance"
       ElseIf TmpRs!PeakFrom = "OC" Then
          MSHFlexGrid1.TextMatrix(I, 13) = "Opening & Closing Bal"
       End If
       
       MSHFlexGrid1.TextMatrix(I, 14) = TmpRs!Remarks
       MSHFlexGrid1.TextMatrix(I, 15) = TmpRs!LocationID
       If TmpRs!StockManagmentContract = False Then
          MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
       Else
          MSHFlexGrid1.TextMatrix(I, 16) = GetServiceTaxNAVCode("5")
       End If
       MSHFlexGrid1.TextMatrix(I, 17) = GetADhocInvoiceAmount(TmpRs!LocationID, TmpRs!ClientID)
       MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs!MasterStorageContractID), "", TmpRs!MasterStorageContractID)
       
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.Text = strUnChecked
       If TmpRs!SCMFlag = False Then
           If mFromDate > "2011-07-01" And Format(TmpRs!BookedFrom, "yyyymm") < Format(mFromDate, "yyyymm") Then
              tmp = "Select * from Txn_ReservationInvoice Where ContractID = " & TmpRs!ContractID & " And YearMonth = '" & mPreviousMonth & "' "
              Call Tmp1Table
        
              If TmpRs1.RecordCount = 0 Then
                 MSHFlexGrid1.row = I
                 For C = 0 To MSHFlexGrid1.Cols - 1
                     MSHFlexGrid1.Col = C
                     MSHFlexGrid1.CellBackColor = vbRed
                 Next
              End If
           End If
       Else
           If mFromDate > "2017-10-01" And Format(TmpRs!BookedFrom, "yyyymm") < Format(mFromDate, "yyyymm") Then
              tmp = "Select * from Txn_ReservationInvoice Where ContractID = " & TmpRs!ContractID & " And YearMonth = '" & mPreviousMonth & "' "
              Call Tmp1Table
        
              If TmpRs1.RecordCount = 0 Then
                 MSHFlexGrid1.row = I
                 For C = 0 To MSHFlexGrid1.Cols - 1
                     MSHFlexGrid1.Col = C
                     MSHFlexGrid1.CellBackColor = vbRed
                 Next
              End If
           End If
       End If
       
       If ChkBOD(TmpRs!LocationID) = False Then
          MSHFlexGrid1.row = I
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.Col = C
              MSHFlexGrid1.CellBackColor = vbYellow
          Next
       End If
       If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
       TmpRs5.Find "ContractID=" & TmpRs!ContractID
       If TmpRs5.EOF = True Then
          MSHFlexGrid1.row = I
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.Col = C
              If MSHFlexGrid1.CellBackColor <> vbRed And MSHFlexGrid1.CellBackColor <> vbYellow Then
                 MSHFlexGrid1.CellBackColor = vbBlue
              End If
          Next
       End If
       Call AddUnmappedGSL(TmpRs!ClientID, TmpRs!LocationID, TmpRs!ContractID)
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
Else
   MsgBox "No record found!!!"
End If

ProgressBar1.Value = ProgressBar1.Max

LblTotalRecord.Caption = TmpRs.RecordCount

CmdSave.Enabled = True

End Sub
Private Function GetADhocInvoiceAmount(mLocID_ As Variant, mClientID_ As Variant)
Dim Retval As Variant
Retval = 0

tmp = "Select LocationID,ClientID,Sum(Amount) As 'Amount' From Txn_TempInvoice "
tmp = tmp & " Where IETypeID = 4 And LocationID = " & mLocID_ & " And ClientID = " & mClientID_ & " And Month(InvoiceDate) = " & mMonth & "  And Year(InvoiceDate) = " & mYear & ""
tmp = tmp & " Group By LocationID,ClientID "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1!Amount
End If

GetADhocInvoiceAmount = Retval

End Function


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
CmdSave.Enabled = False

Call Grid_Head
Call ClearData

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
For I = 2011 To Year(Now)
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

Private Sub CmbLocationID_GotFocus()
tmp1 = CmbLocationID.Text

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

tmp = " Select Distinct TIG.SM_Prefix From Txn_InvMonGSLDetail TIG"
'Commented on 06 Feb 2014
'tmp = tmp & " Where TIG.ExchangeTypeID=1 And TIG.YearMonth = '" & mFromate & "'"
tmp = tmp & " Where (TIG.ExchangeTypeID=1 Or TIG.StackNo = 'Dummy-P') And TIG.YearMonth = '" & mFromDate & "'"

Call TmpTable
mChooseLocation = ""
For I = 1 To TmpRs.RecordCount
    If mChooseLocation = "" Then
       mChooseLocation = "(" & "'" & TmpRs!SM_Prefix & "'"
    Else
       mChooseLocation = mChooseLocation & "," & "'" & TmpRs!SM_Prefix & "'"
    End If
    
    If I = TmpRs.RecordCount Then
       mChooseLocation = mChooseLocation & ")"
    End If
    TmpRs.MoveNext
Next

If mChooseLocation <> "" Then
   If Gen_WcLocation = "" Then
      Call TableMst_Location(" Where SM_Prefix in " & mChooseLocation)
   Else
      Call TableMst_Location("where " & Gen_WcLocation & " And SM_Prefix in " & mChooseLocation)
   End If
Else
   MsgBox "No record found!!! "
   Exit Sub
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
CmbLocationID.Text = tmp1
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
   Call Grid_Head
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

Private Sub ClearData()

CmbMonth.Text = ""
CmbYear.Text = ""
CmbLocationID.Text = ""
CmbNonNcdexClient.Text = ""
CmdSave.Enabled = False
ChkCheckAll.Value = 0


End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_ReservationInvoiceDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_ReservationInvoiceDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 5

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

Private Sub ChkCheckAll_Click()

If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed And MSHFlexGrid1.CellBackColor <> vbYellow And MSHFlexGrid1.CellBackColor <> vbBlue Then
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 0) = strChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If

End Sub

Private Sub MSHFlexGrid1_Click()
Dim ans As Variant
With MSHFlexGrid1
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
           .Col = 0
        If (.CellBackColor) <> vbRed And (.CellBackColor) <> vbYellow And (.CellBackColor) <> vbBlue Then
           If .TextMatrix(.row, .Col) = strUnChecked Then
              .TextMatrix(.row, .Col) = strChecked
           Else
              .TextMatrix(.row, .Col) = strUnChecked
           End If
        Else
           If .CellBackColor = vbRed Then
              If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Then
                 ans = MsgBox("Reservation Billing for previous month not generated. Still you want to process!!!", vbYesNo)
                 If ans = vbYes Then
                    'MsgBox "Reservation Billing for previous month not generated !!!"
                    If .TextMatrix(.row, .Col) = strUnChecked Then
                       .TextMatrix(.row, .Col) = strChecked
                    Else
                       .TextMatrix(.row, .Col) = strUnChecked
                    End If
                 End If
              Else
                 MsgBox "Reservation Billing for previous month not generated !!!"
              End If
           End If
           If .CellBackColor = vbYellow Then
              MsgBox "BOD not updated for any one of CMP in " & .TextMatrix(.row, 3) & "  !!!"
           End If
           If .CellBackColor = vbBlue Then
              MsgBox "Billing details not found in Storage Contract Master!!!"
           End If
        End If
     End If
End With
End Sub

Private Function GetMaxInvoiceNo() As Double
Dim Retval As Double
tmp = "Select max(InvoiceNo) from Txn_ReservationInvoice"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxInvoiceNo = Retval
End Function

Private Function ChkBOD(mLocationID_ As Double) As Boolean
Dim Retval As Boolean
Retval = True

tmp = "Select ML.LocationName"
tmp = tmp & " From Mst_CMP MC Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
'tmp = tmp & " Where MC.CloseDate Is Null "

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where MC.CurrentDate < #" & Format(mToDate, Gen_DatFormat) & "# "
Else
   tmp = tmp & " Where MC.CurrentDate < '" & Format(mToDate, Gen_DatFormat) & "' "
End If

tmp = tmp & " And MC.Closedate is null "

tmp = tmp & " And ML.LocationID = " & mLocationID_

tmp = tmp & " Order by ML.LocationName,MC.CMPName"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = False
End If

ChkBOD = Retval

End Function

Private Sub FillDaywiseBalance(rowID_ As Variant)

Call Grid_Head1

Dim mDispDate, mBillingAmt, mTotQty As Variant
Dim mFWFrom, mFWTo, mSWFrom, mSWTo, mTWfrom, mTWTo, mFrWFrom, mFrWTo, mFFNEnd, mSFNStart, mPerStartDate, mPerEndDate As Variant

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

For x = 1 To mTotalDays

    MSHFlexGrid2.TextMatrix(x, 0) = MSHFlexGrid1.TextMatrix(rowID_, 15)
    MSHFlexGrid2.TextMatrix(x, 1) = MSHFlexGrid1.TextMatrix(rowID_, 3)
    MSHFlexGrid2.TextMatrix(x, 2) = MSHFlexGrid1.TextMatrix(rowID_, 1)
    MSHFlexGrid2.TextMatrix(x, 3) = Format(mDispDate, "dd-MMM-yyyy")
    MSHFlexGrid2.TextMatrix(x, 4) = 0
    MSHFlexGrid2.TextMatrix(x, 5) = 0
    MSHFlexGrid2.TextMatrix(x, 6) = 0
    MSHFlexGrid2.TextMatrix(x, 7) = 0
    MSHFlexGrid2.TextMatrix(x, 8) = 0
    MSHFlexGrid2.TextMatrix(x, 9) = 0
    MSHFlexGrid2.TextMatrix(x, 10) = 0
    MSHFlexGrid2.TextMatrix(x, 11) = 0
    
    mSpilBags = 0
    mSpilQty = 0
    
    If x = 1 Then
       If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
       TmpRs2.Find "ContractID=" & Val(MSHFlexGrid2.TextMatrix(x, 2))
       
       If TmpRs2.EOF = False Then
       
          mSpilBags = 0
          mSpilQty = 0
           
        '  Call GetOpeningBalanceSpillageStock(Val(MSHFlexGrid2.TextMatrix(I, 2)), MSHFlexGrid1.TextMatrix(rowID_, 11))
           
          MSHFlexGrid2.TextMatrix(x, 4) = Format(TmpRs2!BalanceBags + mSpilBags, "#####0")
          MSHFlexGrid2.TextMatrix(x, 5) = Format(TmpRs2!BalanceWeight + mSpilQty, "#####0.000")
           
          mSpilBags = 0
          mSpilQty = 0
           
          Call GetOpeningBalanceSpillagePreStock(Val(MSHFlexGrid2.TextMatrix(x, 2)), MSHFlexGrid1.TextMatrix(rowID_, 11))
           
          MSHFlexGrid2.TextMatrix(x, 4) = Format(Val(MSHFlexGrid2.TextMatrix(x, 4)) - mSpilBags, "#####0")
          MSHFlexGrid2.TextMatrix(x, 5) = Format(Val(MSHFlexGrid2.TextMatrix(x, 5)) - mSpilQty, "#####0.000")
           
           
          '--For Made Up Bags
          mSpilBags = 0
          mSpilQty = 0
          
         ' Call GetMadeUpOpeningBalance(Val(MSHFlexGrid2.TextMatrix(I, 2)), MSHFlexGrid1.TextMatrix(rowID_, 11))
           
         ' MSHFlexGrid2.TextMatrix(I, 4) = Val(MSHFlexGrid2.TextMatrix(I, 4)) + Format(mSpilBags, "#####0")
         ' MSHFlexGrid2.TextMatrix(I, 5) = Val(MSHFlexGrid2.TextMatrix(I, 5)) + Format(mSpilQty, "#####0.000")
           
          mSpilBags = 0
          mSpilQty = 0
           
          Call GetMadeUpOpeningPreBalance(Val(MSHFlexGrid2.TextMatrix(x, 2)), MSHFlexGrid1.TextMatrix(rowID_, 11))
           
          MSHFlexGrid2.TextMatrix(x, 4) = Format(Val(MSHFlexGrid2.TextMatrix(x, 4)) - mSpilBags, "#####0")
          MSHFlexGrid2.TextMatrix(x, 5) = Format(Val(MSHFlexGrid2.TextMatrix(x, 5)) - mSpilQty, "#####0.000")
          
          '-----------
          
       End If
       
    Else
       MSHFlexGrid2.TextMatrix(x, 4) = MSHFlexGrid2.TextMatrix(x - 1, 10)
       MSHFlexGrid2.TextMatrix(x, 5) = MSHFlexGrid2.TextMatrix(x - 1, 11)
    End If
    
    If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
    TmpRs3.Filter = "ContractID=" & Val(MSHFlexGrid2.TextMatrix(x, 2)) & " And  CxTFDate='" & Format(mDispDate, Gen_DatFormat) & "' And TFlag='D'"
    
    If TmpRs3.RecordCount > 0 Then
       For Ln = 1 To TmpRs3.RecordCount
           MSHFlexGrid2.TextMatrix(x, 6) = Val(MSHFlexGrid2.TextMatrix(x, 6)) + Format(TmpRs3!BalanceBags, "#####0")
           MSHFlexGrid2.TextMatrix(x, 7) = Val(MSHFlexGrid2.TextMatrix(x, 7)) + Format(TmpRs3!BalanceWeight, "#####0.000")
           If MSHFlexGrid2.TextMatrix(x, 12) <> "" Then
              MSHFlexGrid2.TextMatrix(x, 12) = MSHFlexGrid2.TextMatrix(x, 12) & "," & IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           Else
              MSHFlexGrid2.TextMatrix(x, 12) = IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           End If
           TmpRs3.MoveNext
       Next
    End If
    
    TmpRs3.Filter = ""
    If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
    TmpRs3.Filter = "ContractID=" & Val(MSHFlexGrid2.TextMatrix(x, 2)) & " And  CxTFDate='" & Format(mDispDate, Gen_DatFormat) & "' And TFlag='W'"
    
    If TmpRs3.RecordCount > 0 Then
       For Ln = 1 To TmpRs3.RecordCount
           mSpilBags = 0
           mSpilQty = 0
           Call GetSpillageWithdrawal(MSHFlexGrid1.TextMatrix(rowID_, 11), TmpRs3!CxTF_DetailsID)
           
           MSHFlexGrid2.TextMatrix(x, 8) = Val(MSHFlexGrid2.TextMatrix(x, 8)) + Format(TmpRs3!BalanceBags + mSpilBags, "#####0")
           MSHFlexGrid2.TextMatrix(x, 9) = Val(MSHFlexGrid2.TextMatrix(x, 9)) + Format(TmpRs3!BalanceWeight + mSpilQty, "#####0.000")
           
           mSpilBags = 0
           mSpilQty = 0
           
''           Call GetMadeUpBagsWithdrawal(MSHFlexGrid1.TextMatrix(rowID_, 11), TmpRs3!CxTFNo, Val(MSHFlexGrid2.TextMatrix(I, 2)))
           
           MSHFlexGrid2.TextMatrix(x, 8) = Val(MSHFlexGrid2.TextMatrix(x, 8)) + Format(mSpilBags, "#####0")
           MSHFlexGrid2.TextMatrix(x, 9) = Val(MSHFlexGrid2.TextMatrix(x, 9)) + Format(mSpilQty, "#####0.000")
                     
           If MSHFlexGrid2.TextMatrix(x, 12) <> "" Then
              MSHFlexGrid2.TextMatrix(x, 12) = MSHFlexGrid2.TextMatrix(x, 12) & "," & IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           Else
              MSHFlexGrid2.TextMatrix(x, 12) = IIf(IsNull(TmpRs3!CxTF_DetailsID), "", TmpRs3!CxTF_DetailsID)
           End If
           
           TmpRs3.MoveNext
       Next
    End If
    
    TmpRs3.Filter = ""
    If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
    MSHFlexGrid2.TextMatrix(x, 10) = Format(Val(MSHFlexGrid2.TextMatrix(x, 4)) + Val(MSHFlexGrid2.TextMatrix(x, 6)) - Val(MSHFlexGrid2.TextMatrix(x, 8)), "#####0")
    MSHFlexGrid2.TextMatrix(x, 11) = Format(Val(MSHFlexGrid2.TextMatrix(x, 5)) + Val(MSHFlexGrid2.TextMatrix(x, 7)) - Val(MSHFlexGrid2.TextMatrix(x, 9)), "#####0.000")
       
    mDispDate = CDate(mDispDate) + 1
    
Next

Call TableTxn_ReservationDailyStock(" Where MonthYear='" & Format(mFromDate, Gen_DatFormat) & "' And ContractID=" & Val(MSHFlexGrid1.TextMatrix(rowID_, 1)))

If RsTxn_ReservationDailyStock.RecordCount = 0 Then
   For x = 1 To mTotalDays
       RsTxn_ReservationDailyStock.AddNew
'       RsTxn_ReservationDailyStock!RDSID = GetMaxRDSID
       RsTxn_ReservationDailyStock!MonthYear = Format(mFromDate, Gen_DatFormat)
       RsTxn_ReservationDailyStock!TxnDate = Format(MSHFlexGrid2.TextMatrix(x, 3), Gen_DatFormat)
       RsTxn_ReservationDailyStock!LocationID = Val(MSHFlexGrid2.TextMatrix(x, 0))
       RsTxn_ReservationDailyStock!ContractID = Val(MSHFlexGrid2.TextMatrix(x, 2))
       RsTxn_ReservationDailyStock!BalanceBags = Val(MSHFlexGrid2.TextMatrix(x, 10))
       RsTxn_ReservationDailyStock!BalanceWeight = Format(MSHFlexGrid2.TextMatrix(x, 11), "#####0.000")
       
       RsTxn_ReservationDailyStock!OpnBalBags = Val(MSHFlexGrid2.TextMatrix(x, 4))
       RsTxn_ReservationDailyStock!OpnBalQty = Format(MSHFlexGrid2.TextMatrix(x, 5), "#####0.000")
       RsTxn_ReservationDailyStock!DepoBags = Val(MSHFlexGrid2.TextMatrix(x, 6))
       RsTxn_ReservationDailyStock!DepoQty = Format(MSHFlexGrid2.TextMatrix(x, 7), "#####0.000")
       RsTxn_ReservationDailyStock!WidraBags = Val(MSHFlexGrid2.TextMatrix(x, 8))
       RsTxn_ReservationDailyStock!WidraQty = Format(MSHFlexGrid2.TextMatrix(x, 9), "#####0.000")
       
       RsTxn_ReservationDailyStock!CxTF_DetailsID = MSHFlexGrid2.TextMatrix(x, 12)
       
       RsTxn_ReservationDailyStock.Update
   Next
   
End If

If Val(MSHFlexGrid1.TextMatrix(rowID_, 18)) <= 0 Then
   RsTxn_ReservationInvoice.AddNew
'   RsTxn_ReservationInvoice!InvoiceNo = GetMaxInvoiceNo
   RsTxn_ReservationInvoice!ContractID = MSHFlexGrid1.TextMatrix(rowID_, 1)
   RsTxn_ReservationInvoice!LocationID = MSHFlexGrid1.TextMatrix(rowID_, 15)
   RsTxn_ReservationInvoice!ServiceTaxNAVCode = MSHFlexGrid1.TextMatrix(rowID_, 16)
   RsTxn_ReservationInvoice!YearMonth = CDate(mFromDate)
   RsTxn_ReservationInvoice!BillingDate = CDate(mToDate)
   RsTxn_ReservationInvoice!BillingAmount = 0 'Val(MSHFlexGrid1.TextMatrix(rowID_, 13))
   RsTxn_ReservationInvoice!G_UID = GetGUID
   mG_UIDInv = RsTxn_ReservationInvoice!G_UID
   RsTxn_ReservationInvoice!Flag = "P"
   RsTxn_ReservationInvoice!CreatedBy = Gen_UserLoginID
   RsTxn_ReservationInvoice!CreatedDate = Now
   RsTxn_ReservationInvoice.Update

    tmp = "Select InvoiceNo From Txn_ReservationInvoice Where G_UID='" & mG_UIDInv & "'"
    Call Tmp7Table
    
    If TmpRs7.RecordCount > 0 Then
    '   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
       mInvNo = TmpRs7!InvoiceNo
    End If

Else
   If RsTxn_ReservationInvoice.RecordCount > 0 Then RsTxn_ReservationInvoice.MoveFirst
   RsTxn_ReservationInvoice.Find "ContractID=" & Val(MSHFlexGrid1.TextMatrix(rowID_, 18))
   If RsTxn_ReservationInvoice.EOF = True Then
      RsTxn_ReservationInvoice.AddNew
'      RsTxn_ReservationInvoice!InvoiceNo = GetMaxInvoiceNo
      RsTxn_ReservationInvoice!ContractID = MSHFlexGrid1.TextMatrix(rowID_, 18)
      RsTxn_ReservationInvoice!LocationID = MSHFlexGrid1.TextMatrix(rowID_, 15)
      RsTxn_ReservationInvoice!ServiceTaxNAVCode = MSHFlexGrid1.TextMatrix(rowID_, 16)
      RsTxn_ReservationInvoice!YearMonth = CDate(mFromDate)
      RsTxn_ReservationInvoice!BillingDate = CDate(mToDate)
      RsTxn_ReservationInvoice!BillingAmount = 0 'Val(MSHFlexGrid1.TextMatrix(rowID_, 13))
      RsTxn_ReservationInvoice!G_UID = GetGUID
      RsTxn_ReservationInvoice!Flag = "P"
      RsTxn_ReservationInvoice!CreatedBy = Gen_UserLoginID
      RsTxn_ReservationInvoice!CreatedDate = Now
      RsTxn_ReservationInvoice.Update
   
        tmp = "Select InvoiceNo From Txn_ReservationInvoice Where G_UID='" & mG_UIDInv & "'"
        Call Tmp7Table
        
        If TmpRs7.RecordCount > 0 Then
        '   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
           mInvNo = TmpRs7!InvoiceNo
        End If
   
   End If
End If

tmp = "Select FromDate,ToDate,BillingCycleID,BillingQty,BillingRate from Mst_SContractBillingDetail"
tmp = tmp & " Where FromDate<='" & mToDate & "' And ToDate>='" & mFromDate & "' And ContractID=" & Val(MSHFlexGrid1.TextMatrix(rowID_, 1))

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   For x = 1 To TmpRs4.RecordCount
       tmp = "Select TxnDate,BalanceBags,BalanceWeight,OpnBalBags,OpnBalQty from Txn_ReservationDailyStock Where MonthYear='" & Format(mFromDate, Gen_DatFormat) & "' And ContractID =" & MSHFlexGrid1.TextMatrix(rowID_, 1) & " And TxnDate>='" & Format(TmpRs4!FromDate, Gen_DatFormat) & "' And TxnDate<='" & Format(TmpRs4!ToDate, Gen_DatFormat) & "'"
       Call Tmp6Table
'Daily Billing Cycle
       If TmpRs4!BillingCycleID = 1 Then
          If TmpRs6.RecordCount > 0 Then
             For Ln = 1 To TmpRs6.RecordCount
                 RsTxn_ReservationInvoiceDetail.AddNew
'                 RsTxn_ReservationInvoiceDetail!RIDID = GetMaxRIDID
                 RsTxn_ReservationInvoiceDetail!InvoiceNo = mInvNo 'RsTxn_ReservationInvoice!InvoiceNo
                 RsTxn_ReservationInvoiceDetail!BillingCycleID = TmpRs4!BillingCycleID
                 RsTxn_ReservationInvoiceDetail!FromDate = TmpRs6!TxnDate
                 RsTxn_ReservationInvoiceDetail!ToDate = TmpRs6!TxnDate
                 If LCase(MSHFlexGrid1.TextMatrix(rowID_, 10)) = "weight in mt" Then
                    If LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "closing balance" Then
                        If TmpRs6!BalanceWeight > TmpRs4!BillingQty Then
                           RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs6!BalanceWeight
                        Else
                           RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
                        End If
                    ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "opening & closing bal" Then
                        If TmpRs6!BalanceWeight > TmpRs6!OpnBalQty Then
                           If TmpRs6!BalanceWeight > TmpRs4!BillingQty Then
                              RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs6!BalanceWeight
                           Else
                              RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
                           End If
                        Else
                           If TmpRs6!OpnBalQty > TmpRs4!BillingQty Then
                              RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs6!OpnBalQty
                           Else
                              RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
                           End If
                        End If
                    End If
                    
                    RsTxn_ReservationInvoiceDetail!Rate = TmpRs4!BillingRate
                    RsTxn_ReservationInvoiceDetail!Amount = TmpRs4!BillingRate * RsTxn_ReservationInvoiceDetail!PeakValue
                    RsTxn_ReservationInvoiceDetail!Flag = "Q"
                 ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 10)) = "no of bags" Then
                    If LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "closing balance" Then
                        If TmpRs6!BalanceBags > Val(TmpRs4!BillingQty) Then 'Format(TmpRs4!BillingQty, "#####") Then
                           RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs6!BalanceBags
                        Else
                           RsTxn_ReservationInvoiceDetail!PeakValue = Format(TmpRs4!BillingQty, "#####.00")
                        End If
                    ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "opening & closing bal" Then
                        If TmpRs6!BalanceBags > TmpRs6!OpnBalBags Then
                           If TmpRs6!BalanceBags > TmpRs4!BillingQty Then 'Format(TmpRs4!BillingQty, "#####") Then
                              RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs6!BalanceBags
                           Else
                              RsTxn_ReservationInvoiceDetail!PeakValue = Format(TmpRs4!BillingQty, "#####.00")
                           End If
                        Else
                           If TmpRs6!OpnBalBags > TmpRs4!BillingQty Then 'Format(TmpRs4!BillingQty, "#####") Then
                              RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs6!OpnBalBags
                           Else
                              RsTxn_ReservationInvoiceDetail!PeakValue = Format(TmpRs4!BillingQty, "#####.00")
                           End If
                        End If
                    End If
                    RsTxn_ReservationInvoiceDetail!Rate = TmpRs4!BillingRate
                    RsTxn_ReservationInvoiceDetail!Amount = TmpRs4!BillingRate * RsTxn_ReservationInvoiceDetail!PeakValue
                    RsTxn_ReservationInvoiceDetail!Flag = "B"
                 ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 10)) = "fixed" Then
                    RsTxn_ReservationInvoiceDetail!PeakValue = Null
           
                    RsTxn_ReservationInvoiceDetail!Rate = TmpRs4!BillingRate
                    RsTxn_ReservationInvoiceDetail!Amount = TmpRs4!BillingRate
                    RsTxn_ReservationInvoiceDetail!Flag = "F"
                 End If
                 RsTxn_ReservationInvoiceDetail!ContractID = Val(MSHFlexGrid1.TextMatrix(rowID_, 1))
                 RsTxn_ReservationInvoiceDetail.Update
                 TmpRs6.MoveNext
             Next
          End If
'Weekly Billing Cycle
       ElseIf TmpRs4!BillingCycleID = 2 Then
          If TmpRs6.RecordCount > 0 Then
             If TmpRs4!FromDate <= CDate(mFWTo) And TmpRs4!ToDate >= CDate(mFWFrom) Then
                Call AddBillingCycleWisePeak(mFWFrom, mFWTo, rowID_, TmpRs4!BillingRate)
             End If
             If TmpRs4!FromDate <= CDate(mSWTo) And TmpRs4!ToDate >= CDate(mSWFrom) Then
                Call AddBillingCycleWisePeak(mSWFrom, mSWTo, rowID_, TmpRs4!BillingRate)
             End If
             If TmpRs4!FromDate <= CDate(mTWTo) And TmpRs4!ToDate >= CDate(mTWfrom) Then
                Call AddBillingCycleWisePeak(mTWfrom, mTWTo, rowID_, TmpRs4!BillingRate)
             End If
             If TmpRs4!FromDate <= CDate(mFrWTo) And TmpRs4!ToDate >= CDate(mFrWFrom) Then
                Call AddBillingCycleWisePeak(mFrWFrom, mFrWTo, rowID_, TmpRs4!BillingRate)
             End If
          End If
'Fortnightly Billing Cycle
       ElseIf TmpRs4!BillingCycleID = 3 Then
          If TmpRs6.RecordCount > 0 Then
             If TmpRs4!FromDate <= CDate(mFFNEnd) And TmpRs4!ToDate >= CDate(mFWFrom) Then
                Call AddBillingCycleWisePeak(mFWFrom, mFFNEnd, rowID_, TmpRs4!BillingRate)
             End If
             If TmpRs4!FromDate <= CDate(mFrWTo) And TmpRs4!ToDate >= CDate(mSFNStart) Then
                Call AddBillingCycleWisePeak(mSFNStart, mFrWTo, rowID_, TmpRs4!BillingRate)
             End If
          End If
'Monthly Billing Cycle
       ElseIf TmpRs4!BillingCycleID = 4 Then
          If TmpRs6.RecordCount > 0 Then
             Call AddBillingCycleWisePeak(mFWFrom, mFrWTo, rowID_, TmpRs4!BillingRate)
          End If
'Period Billing Cycle
       ElseIf TmpRs4!BillingCycleID = 5 Then
          If TmpRs6.RecordCount > 0 Then
             mPerStartDate = CDate(mFWFrom)
             mPerEndDate = CDate(mFrWTo)
             If TmpRs4!FromDate > CDate(mFWFrom) Then mPerStartDate = TmpRs4!FromDate
             If TmpRs4!ToDate < CDate(mFrWTo) Then mPerEndDate = TmpRs4!ToDate
             mPerTotDays = DateDiff("d", mPerStartDate, mPerEndDate) + 1
             mPerRate = mPerTotDays * (TmpRs4!BillingRate / mTotalDays)
             Call AddBillingCycleWisePeak(mPerStartDate, mPerEndDate, rowID_, mPerRate)
          End If
       End If
       TmpRs4.MoveNext
   Next
End If

mRemarks = GenerateRemarks(mInvNo)

tmp = "Update Txn_ReservationInvoice Set BillingAmount= (select Sum(isnull(Amount,0)) from Txn_ReservationInvoiceDetail Where InvoiceNo=" & mInvNo & ")"
tmp = tmp & " Where InvoiceNo=" & RsTxn_ReservationInvoice!InvoiceNo & ""
Call TmpTable

tmp = "Update Txn_ReservationInvoice Set Remarks = '" & mRemarks & "' Where InvoiceNo=" & RsTxn_ReservationInvoice!InvoiceNo & ""
Call TmpTable

tmp = "Update Txn_InvMonGSLDetail Set PINO=" & mInvNo & ", InvoiceDate='" & Format(mToDate, Gen_DatFormat) & "' Where ContractID=" & MSHFlexGrid1.TextMatrix(rowID_, 1) & " And YearMonth='" & Format(mFromDate, Gen_DatFormat) & "'"
Call TmpTable

tmp = "Select  BillingAmount from Txn_ReservationInvoice Where InvoiceNo=" & mInvNo & ""
Call TmpTable
mBillingAmt = 0
If TmpRs.RecordCount > 0 Then
   mBillingAmt = TmpRs!BillingAmount
End If
tmp = "Select Mst_GSL_ID,PINO,InvoiceDate,ContractID,isnull(BalanceWeight,0)BalanceWeight,InvoiceAmount,MonG_UID from Txn_InvMonGSLDetail  Where PINO=" & mInvNo & " And ContractID=" & MSHFlexGrid1.TextMatrix(rowID_, 1) & " And YearMonth='" & Format(mFromDate, Gen_DatFormat) & "'"   '' 'Where YearMonth ='" & mFromDate & "' and isnull(ContractID,0)>0 And PINO is not null"
Call Tmp5Table

mTotQty = 0

If TmpRs5.RecordCount > 0 Then
   For Ln = 1 To TmpRs5.RecordCount
       If TmpRs5!BalanceWeight > 0 Then
          mTotQty = mTotQty + TmpRs5!BalanceWeight
       End If
       TmpRs5.MoveNext
   Next Ln
   If mTotQty > 0 Then
      tmp = " Update Txn_InvMonGSLDetail Set InvoiceAmount=" & mBillingAmt & " * BalanceWeight/" & mTotQty & "  Where PINO=" & mInvNo & " And ContractID=" & Val(MSHFlexGrid1.TextMatrix(rowID_, 1)) & " And YearMonth ='" & Format(mFromDate, Gen_DatFormat) & "' And BalanceWeight>0"
      Call TmpTable
   End If
End If


End Sub
Private Function GenerateRemarks(mInv_)


Dim Retval As Variant

mCommodityName = ""
tmp = "Select com.Commodity from Txn_ReservationInvoice a Inner Join Mst_StorageContract c on a.ContractID = c.ContractID Left Join Mst_Commodity com on c.CommodityID = com.CommodityID"
tmp = tmp & " Where c.CommodityID is not null and a.InvoiceNo = " & mInv_ & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mCommodityName = TmpRs!Commodity
End If

mBookedAsPer = ""
mContractRemarks = ""
tmp = "Select c.BookedAsPer,c.Remarks from Txn_ReservationInvoice a Inner Join Mst_StorageContract c on a.ContractID = c.ContractID "
tmp = tmp & " Where a.InvoiceNo = " & mInv_ & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mBookedAsPer = TmpRs!BookedAsPer
   mContractRemarks = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
End If

mLocationName = ""

tmp = "Select c.LocationName from Txn_ReservationInvoice a Inner Join Mst_Location c on a.LocationID = c.LocationID "
tmp = tmp & " Where a.InvoiceNo = " & mInv_ & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mLocationName = TmpRs!Locationname
End If

If LCase(mBookedAsPer) <> "fixed" Then
   mNarration = "Being Storage Reservation charge for " & mCommodityName & ""
Else
   mNarration = mContractRemarks
End If


tmp = "Select a.FromDate,a.ToDate,a.PeakValue,a.Rate,c.BillingCycle from Txn_ReservationInvoiceDetail  a"
tmp = tmp & " Inner Join Mst_BillingCycle c on a.BillingCycleID = c.BillingCycleID"
tmp = tmp & " Where a.InvoiceNO = " & mInv_ & ""

Call TmpTable
abc = 0

For x = 1 To TmpRs.RecordCount
    If LCase(mBookedAsPer) = "fixed" Then
        'mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Unit/Month "
    ElseIf LCase(mBookedAsPer) = "no Of bags" And mInvCommodityID = 43 Then
        If x = TmpRs.RecordCount Then
           If CDate(TmpRs!ToDate) < "" & CDate(mToDate) & "" Then
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bales/" & TmpRs!BillingCycle & " till " & Format(TmpRs!ToDate, "DD-MMM-YYYY") & ""
              abc = 1
           Else
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bales/" & TmpRs!BillingCycle & " "
           End If
        ElseIf x = 1 Then
           If CDate(TmpRs!FromDate) > "" & CDate(mFromDate) & "" Then
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bales/" & TmpRs!BillingCycle & " "
              abc = 2
              mMidFromDate = TmpRs!FromDate
           Else
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bales/" & TmpRs!BillingCycle & " "
           End If
        Else
           mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bales/" & TmpRs!BillingCycle & " "
        End If
    ElseIf LCase(mBookedAsPer) = "no of bags" And mInvCommodityID <> 43 Then
        If x = TmpRs.RecordCount Then
           If CDate(TmpRs!ToDate) < "" & CDate(mToDate) & "" Then
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bags/" & TmpRs!BillingCycle & " till " & Format(TmpRs!ToDate, "DD-MMM-YYYY") & ""
              abc = 1
           Else
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bags/" & TmpRs!BillingCycle & " "
           End If
        ElseIf x = 1 Then
           If CDate(TmpRs!FromDate) > "" & CDate(mFromDate) & "" Then
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bags/" & TmpRs!BillingCycle & " "
              abc = 2
              mMidFromDate = TmpRs!FromDate
           Else
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bags/" & TmpRs!BillingCycle & " "
           End If
        Else
           mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per Bags/" & TmpRs!BillingCycle & " "
        End If
    Else
        If x = TmpRs.RecordCount Then
           If CDate(TmpRs!ToDate) < "" & CDate(mToDate) & "" Then
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per MT/" & TmpRs!BillingCycle & " till " & Format(TmpRs!ToDate, "DD-MMM-YYYY") & ""
              abc = 1
           Else
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per MT/" & TmpRs!BillingCycle & " "
           End If
        ElseIf x = 1 Then
           If CDate(TmpRs!FromDate) > "" & CDate(mFromDate) & "" Then
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per MT/" & TmpRs!BillingCycle & " "
              mMidFromDate = TmpRs!FromDate
              abc = 2
           Else
              mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per MT/" & TmpRs!BillingCycle & " "
           End If
        Else
           mNarration = mNarration & " " & TmpRs!PeakValue & "@" & TmpRs!Rate & "/Per MT/" & TmpRs!BillingCycle & " "
        End If
    End If
    TmpRs.MoveNext
Next

If abc = 0 Then
   mNarration = mNarration & " For " & Format(CDate(mFromDate), "MMM-YYYY") & " at " & mLocationName & " Location. "
ElseIf abc = 2 Then
   mNarration = mNarration & " at " & mLocationName & " Location. (Start from " & Format(CDate(mMidFromDate), "DD-MMM-YYYY") & ")"
Else
   mNarration = mNarration & " at " & mLocationName & " Location. "
End If

GenerateRemarks = mNarration

End Function
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

Private Function GetMaxRDSID() As Double
Dim Retval As Double
tmp = "Select max(RDSID) from Txn_ReservationDailyStock"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxRDSID = Retval
End Function

Private Function GetMaxRIDID() As Double
Dim Retval As Double
tmp = "Select max(RIDID) from Txn_ReservationInvoiceDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxRIDID = Retval
End Function

Private Sub AddBillingCycleWisePeak(mFrom_ As Variant, mTo_ As Variant, rowID_ As Variant, mRate_ As Variant)
tmp = "Select Max(BalanceWeight)BalanceWeight,Max(BalanceBags)BalanceBags,Max(OpnBalQty)OpnBalQty,Max(OpnBalBags)OpnBalBags from Txn_ReservationDailyStock Where ContractID =" & MSHFlexGrid1.TextMatrix(rowID_, 1) & " And (TxnDate>='" & mFrom_ & "' And TxnDate<='" & mTo_ & "') " 'And TxnDate>='" & TmpRs4!FromDate & "' And TxnDate<='" & TmpRs4!ToDate & "'"
Call Tmp7Table
If TmpRs7.RecordCount > 0 Then
    RsTxn_ReservationInvoiceDetail.AddNew
'    RsTxn_ReservationInvoiceDetail!RIDID = GetMaxRIDID
    RsTxn_ReservationInvoiceDetail!InvoiceNo = RsTxn_ReservationInvoice!InvoiceNo
    RsTxn_ReservationInvoiceDetail!BillingCycleID = TmpRs4!BillingCycleID
    RsTxn_ReservationInvoiceDetail!FromDate = mFrom_
    RsTxn_ReservationInvoiceDetail!ToDate = mTo_
    If LCase(MSHFlexGrid1.TextMatrix(rowID_, 10)) = "weight in mt" Then
       If LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "closing balance" Then
          If TmpRs7!BalanceWeight > TmpRs4!BillingQty Then
             RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!BalanceWeight
          Else
             RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
          End If
       ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "opening & closing bal" Then
          If TmpRs7!BalanceWeight > TmpRs7!OpnBalQty Then
             If TmpRs7!BalanceWeight > TmpRs4!BillingQty Then
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!BalanceWeight
             Else
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
             End If
          Else
             If TmpRs7!OpnBalQty > TmpRs4!BillingQty Then
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!OpnBalQty
             Else
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
             End If
          End If
       End If
                    
                    
''       If TmpRs7!BalanceWeight > TmpRs4!BillingQty Then
''          RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!BalanceWeight
''       Else
''          RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
''       End If
       RsTxn_ReservationInvoiceDetail!Rate = TmpRs4!BillingRate
       RsTxn_ReservationInvoiceDetail!Amount = mRate_ * RsTxn_ReservationInvoiceDetail!PeakValue 'TmpRs4!BillingRate
       RsTxn_ReservationInvoiceDetail!Flag = "Q"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 10)) = "no of bags" Then
       If LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "closing balance" Then
          If TmpRs7!BalanceBags > Val(TmpRs4!BillingQty) Then
             RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!BalanceBags
          Else
             RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
          End If
       ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 13)) = "opening & closing bal" Then
          If TmpRs7!BalanceBags > TmpRs7!OpnBalBags Then
             If TmpRs7!BalanceBags > Val(TmpRs4!BillingQty) Then
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!BalanceBags
             Else
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
             End If
          Else
             If TmpRs7!OpnBalBags > Val(TmpRs4!BillingQty) Then
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!OpnBalBags
             Else
                RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
             End If
          End If
       End If

''       If TmpRs7!BalanceBags > TmpRs4!BillingQty Then
''          RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs7!BalanceBags
''       Else
''          RsTxn_ReservationInvoiceDetail!PeakValue = TmpRs4!BillingQty
''       End If
       RsTxn_ReservationInvoiceDetail!Rate = TmpRs4!BillingRate
       RsTxn_ReservationInvoiceDetail!Amount = mRate_ * RsTxn_ReservationInvoiceDetail!PeakValue
       RsTxn_ReservationInvoiceDetail!Flag = "B"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(rowID_, 10)) = "fixed" Then
       RsTxn_ReservationInvoiceDetail!PeakValue = Null
       RsTxn_ReservationInvoiceDetail!Rate = TmpRs4!BillingRate
       RsTxn_ReservationInvoiceDetail!Amount = mRate_ 'TmpRs4!BillingRate
       RsTxn_ReservationInvoiceDetail!Flag = "F"
    End If
    RsTxn_ReservationInvoiceDetail!ContractID = Val(MSHFlexGrid1.TextMatrix(rowID_, 1))
    RsTxn_ReservationInvoiceDetail.Update
End If

End Sub

Private Sub Grid_Head2()

With MSHFlexGrid3
    .Clear
    .Rows = 3
    .WordWrap = True
    .Cols = 20
     .FixedRows = 2
     .TextMatrix(1, 0) = "Select"
     .TextMatrix(1, 1) = "Location Name"
     .TextMatrix(1, 2) = "CMP Name"
     .TextMatrix(1, 3) = "Client Name"
     .TextMatrix(1, 4) = "Commodity"
     .TextMatrix(1, 5) = "Exchange Type"
     .TextMatrix(1, 6) = "Txn Type"
     .TextMatrix(1, 7) = "Billing Month" '"Year Month"
     .TextMatrix(1, 8) = "Txn Date" '"CxTF Date"
     .TextMatrix(1, 9) = "CDTF Date" '.TextMatrix(0, 19) = "CDTF Date"
     .TextMatrix(1, 10) = "No Of Bags" '"Balance Bags"
     .TextMatrix(1, 11) = "Weight in MT" '"Balance Weight"
     .TextMatrix(1, 12) = "Bagsize"
     .TextMatrix(1, 13) = "GSLID"
     .TextMatrix(1, 14) = "ClientIDRow"
     .TextMatrix(1, 15) = "CommodityID"
     .TextMatrix(1, 16) = "ExchangeTypeID"
     .TextMatrix(1, 17) = "LocationID"
     .TextMatrix(1, 18) = "SM_Prefix"
     .TextMatrix(1, 19) = "CxTFNo"
     
     .ColWidth(0) = 800
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
     .ColWidth(19) = 0
     
End With

End Sub
Private Sub AddUnmappedGSL(ClientID_ As Variant, LocationID_ As Variant, ContractID_ As Variant)

tmp = "Select ML.LocationName,TINV.SM_PreFix,MCMP.CMPName,MClient.ClientName,MCG.CommodityGroup, Me.ExchangeType ,"
tmp = tmp & " Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight', TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, TINV.ClientID, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TINV.ExchangeTypeID,TINV.Mst_GSL_ID,TINV.CDTFDate,MC.Commodity,"
tmp = tmp & " TInv.Bagsize , TInv.CxTF_DetailsID"
tmp = tmp & " From Txn_InvMonGSLDetail  TINV"
tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MClient On TINV.ClientID = MClient.ClientID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Where TINV.ExchangeTypeID = 1 and YearMonth='" & mFromDate & "' And isnull(ContractID,0)=0 And isnull(PINo,0)=0"
tmp = tmp & " And TINV.ClientID=" & ClientID_ & " And ML.LocationID=" & LocationID_

tmp = tmp & " Group By ML.LocationName,TINV.SM_PreFix,MCMP.CMPName,MClient.ClientName,MCG.CommodityGroup, Me.ExchangeType ,"
tmp = tmp & " TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, TINV.ClientID, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TInv.ExchangeTypeID , TInv.Mst_GSL_ID, TInv.CDTFDate, MC.Commodity, TInv.Bagsize, TInv.CxTF_DetailsID"

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   With MSHFlexGrid3
   For I = 1 To TmpRs4.RecordCount
        
        mR = MSHFlexGrid3.Rows - 1
       .TextMatrix(mR, 1) = TmpRs4!Locationname
       .TextMatrix(mR, 2) = TmpRs4!CMPName
       .TextMatrix(mR, 3) = TmpRs4!ClientName
       .TextMatrix(mR, 4) = TmpRs4!COmmodityGroup
       .TextMatrix(mR, 5) = IIf(IsNull(TmpRs4!ExchangeType), "", TmpRs4!ExchangeType)
       .TextMatrix(mR, 6) = TmpRs4!tFlag
       .TextMatrix(mR, 7) = Format(IIf(IsNull(TmpRs4!YearMonth), "", TmpRs4!YearMonth), "dd-MMM-yyyy")
       .TextMatrix(mR, 8) = Format(IIf(IsNull(TmpRs4!CxTFDate), "", TmpRs4!CxTFDate), "dd-MMM-yyyy")
       .TextMatrix(mR, 9) = IIf(IsNull(TmpRs4!CDTFDate), "", TmpRs4!CDTFDate)
       .TextMatrix(mR, 9) = Format(.TextMatrix(mR, 9), "dd-mmm-yyyy")
       
       .TextMatrix(mR, 10) = IIf(IsNull(TmpRs4!BalanceBags), 0, TmpRs4!BalanceBags)
       .TextMatrix(mR, 11) = IIf(IsNull(TmpRs4!BalanceWeight), "", TmpRs4!BalanceWeight)
       .TextMatrix(mR, 12) = ContractID_ 'Format(IIf(IsNull(TmpRs4!BagSize), 0, TmpRs4!BagSize), "########0.00")
       .TextMatrix(mR, 13) = TmpRs4!Mst_GSL_ID
       .TextMatrix(mR, 14) = IIf(IsNull(TmpRs4!ClientID), "", TmpRs4!ClientID)
       .TextMatrix(mR, 15) = IIf(IsNull(TmpRs4!CommodityID), "", TmpRs4!CommodityID)
       .TextMatrix(mR, 16) = IIf(IsNull(TmpRs4!ExchangeTypeID), "", TmpRs4!ExchangeTypeID)
       .TextMatrix(mR, 17) = IIf(IsNull(TmpRs4!LocationID), "", TmpRs4!LocationID)
       .TextMatrix(mR, 18) = IIf(IsNull(TmpRs4!SM_Prefix), "", TmpRs4!SM_Prefix)
       .TextMatrix(mR, 19) = IIf(IsNull(TmpRs4!CxTF_DetailsID), "", TmpRs4!CxTF_DetailsID) '"BillingUnit"
       TmpRs4.MoveNext
      MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
    Next
    End With
End If
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
tmp = tmp & " And TCD.CxTF_DetailsID IN (Select Distinct TCG.CxTF_DetailsID From Txn_CxTF_GSL TCG  Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix Where TCG.TxnType = 'W' And MGSL.NContractID=" & mContractID_ & " And TCG.SM_Prefix='" & mPrefix_ & "')"
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
