VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmTxn_UpdateGodownBalanceStock 
   Caption         =   "Godwon Balance Stock Update"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14955
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         ItemData        =   "FrmTxn_UpdateGodownBalanceStock.frx":0000
         Left            =   8880
         List            =   "FrmTxn_UpdateGodownBalanceStock.frx":0002
         TabIndex        =   7
         Top             =   2220
         Width           =   1500
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
         ItemData        =   "FrmTxn_UpdateGodownBalanceStock.frx":0004
         Left            =   6720
         List            =   "FrmTxn_UpdateGodownBalanceStock.frx":0006
         TabIndex        =   6
         Top             =   2220
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         Height          =   400
         Left            =   10620
         TabIndex        =   8
         Top             =   2205
         Width           =   1260
      End
      Begin VB.CheckBox Chk_Transaction 
         Caption         =   "No of Transaction, Bags and Quantity"
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
         TabIndex        =   5
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CheckBox Chk_Balance 
         Caption         =   "Balance Bags And Quantity"
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
         TabIndex        =   4
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3255
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
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   4320
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
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   705
         Width           =   4320
      End
      Begin VB.ComboBox CmbGodownNo 
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
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1155
         Width           =   4320
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
         Left            =   8160
         TabIndex        =   14
         Top             =   2280
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
         Left            =   6000
         TabIndex        =   13
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   765
      End
      Begin VB.Label LblCMPName 
         AutoSize        =   -1  'True
         Caption         =   "CMP"
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
         Left            =   120
         TabIndex        =   10
         Top             =   765
         Width           =   435
      End
      Begin VB.Label lblGodownNo 
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
         Left            =   120
         TabIndex        =   9
         Top             =   1215
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmTxn_UpdateGodownBalanceStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mGodownID As Variant

Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   If RsMst_Location.RecordCount > 0 Then
      CmbLocationID.SetFocus
   End If
   Exit Sub
End If
tmp = CmbCMPName.Text
Call TableMst_CMP(" Where CloseDate is null And LocationID = " & mLocationID)
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbGodownNo.Text = ""
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodownNo.Text = ""
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodownNo_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownNo.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
CmbGodownNo.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownNo.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownNo.Text = tmp
End Sub

Private Sub CmbGodownNo_LostFocus()
If CmbGodownNo.Text = "" Then Exit Sub

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownNo.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   If CmbGodownNo.Enabled = True Then
      CmbGodownNo.SetFocus
   End If
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationID <> 0 ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   CmbCMPName.Text = ""
   CmbGodownNo.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

If mLocationID <> RsMst_Location!LocationID Then
   CmbCMPName.Text = ""
   CmbGodownNo.Text = ""
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdUpdate_Click()

MsgBox "Wait till next message!!!"
ProgressBar1.Value = 0

If Chk_Balance.Value = vbUnchecked And Chk_Transaction.Value = vbUnchecked Then
    MsgBox "Select any one CheckBox to Update Data!!!"
    Exit Sub
End If

If Chk_Transaction.Value = vbChecked Then
    If CmbMonth.Text = "" Then
        MsgBox "Blank Month not Allowed!!!"
        CmbMonth.SetFocus
        Exit Sub
    End If
    If CmbYear.Text = "" Then
        MsgBox "Blank Year not Allowed!!!"
        CmbYear.SetFocus
        Exit Sub
    End If
    'mMonthListindex month in digit
    mStartDate = Format(CDate("1-" & CmbMonth.Text & "-" & CmbYear.Text), "yyyy-mm-dd")
    mTotalDays = Day(DateSerial(Year(CDate(mStartDate)), Month(CDate(mStartDate)) + 1, 0))
    mEndDate = Format(CDate(mTotalDays & "-" & CmbMonth.Text & "-" & CmbYear.Text), "yyyy-mm-dd")
    If mStartDate > Format(Date, "yyyy-mm-dd") Then
        MsgBox "Future month not allowed!!!"
        Exit Sub
    End If
End If
Dim wc As Variant
wc = ""

If CmbGodownNo.Text <> "" Then
    wc = wc & " And MGSL.GodownID = " & mGodownID
Else
    If CmbCMPName.Text <> "" Then
        wc = wc & " And MGSL.CMPID = " & mCMPID
    Else
        If CmbLocationID.Text <> "" Then
            wc = wc & " And ML.LocationID = " & mLocationID
        End If
    End If
End If
    

tmp = "Update Mst_Godown set BalanceBags = 0, BalanceQty = 0,NoofBags = 0, NoOfQty = 0 , NoOfTransactions = 0,YearMonth = '" & mStartDate & "' " 'Where GodownID = " & TmpRs!GodownID & " "
Call Tmp1Table
    
    
    
If Chk_Balance.Value = vbChecked Then
    tmp = " Select MGSL.GodownID, Sum (MGSL.BalanceBags) 'BalanceBags', sum(MGSL.BalanceWeight) 'BalanceWeight'"
    tmp = tmp & " from Mst_GSL MGSL"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
    tmp = tmp & " Inner Join Mst_Cmp MCMP on MCMP.CMPID = MGSL.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
    tmp = tmp & " where MGSL.flag in ('O','J','R') And MG.CloseDate is null"
    tmp = tmp & wc & " group by MGSL.GodownID"
    Call TmpTable
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = TmpRs.RecordCount + 1

    If TmpRs.RecordCount > 0 Then
        For I = 1 To TmpRs.RecordCount
            tmp = "Update Mst_Godown set BalanceBags = " & TmpRs!BalanceBags & ", BalanceQty = " & TmpRs!BalanceWeight & " Where GodownID = " & TmpRs!GodownID & " "
            Call Tmp1Table
            TmpRs.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
    ProgressBar1.Value = ProgressBar1.Max
    End If
End If

If Chk_Transaction.Value = vbChecked Then
    If CmbGodownNo.Text <> "" Then
        tmp = "Select YearMonth from Mst_Godown Where GodownID = " & mGodownID
        Call TmpTable
        
        If IsNull(TmpRs!YearMonth) = False Then
            If Format(TmpRs!YearMonth, "yyyy-mm-dd") >= mStartDate Then
                ans = MsgBox("Data already Update For month of " & Format(TmpRs!YearMonth, "MMM-yyyy") & ", Do you want to update it???", vbYesNo)
                If ans = vbNo Then
                    MsgBox "Done!!!"
                    Exit Sub
                End If
            End If
        End If
    End If
    
    tmp = " Select  MGSL.GodownID, sum(TCGSL.NoOfBags) NoOfBags, sum(TCGSL.GrossWeight) GrossWeight"
    tmp = tmp & " From Txn_CxTF_Details  TCD"
    tmp = tmp & " Inner Join Txn_CxTF_GSL TCGSL on TCGSL.CxTFNo = TCD.CxTFNo And TCGSL.TxnType = TCD.TxnType And TCGSL.SM_Prefix = TCD.SM_Prefix"
    tmp = tmp & " Inner Join Mst_GSL MGSL on  MGSL.GSLID = TCGSL.GSLID And MGSL.SM_Prefix = TCGSL.SM_Prefix"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
    tmp = tmp & " Inner Join Mst_Cmp MCMP on MCMP.CMPID = MGSL.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
    If Gen_DBType <> "MDB" Then
        tmp = tmp & " Where TCD.CxTFDate between '" & mStartDate & "' And '" & mEndDate & "' And MG.CloseDate is null "
    Else
        tmp = tmp & " Where TCD.CxTFDate between #" & mStartDate & "# And #" & mEndDate & "# And MG.CloseDate is null "
    End If
    tmp = tmp & wc & " group by MGSL.GodownID"
    Call TmpTable
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = TmpRs.RecordCount + 1
    

    If TmpRs.RecordCount > 0 Then
        For I = 1 To TmpRs.RecordCount
            tmp = "Update Mst_Godown set NoofBags = " & TmpRs!NoofBags & ", NoOfQty = " & TmpRs!GrossWeight & " Where GodownID = " & TmpRs!GodownID & " "
            Call Tmp1Table
            TmpRs.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
    ProgressBar1.Value = ProgressBar1.Max
    End If
    
    tmp = " Select"
    tmp = tmp & " MGSL.GodownID, count(distinct convert(varchar, TCD.CxTFNo)+''+TCD.TxnType) NoofTransaction "
    tmp = tmp & " From Txn_CxTF_Details TCD"
    tmp = tmp & " Inner Join Txn_CxTF_GSL TCGSL on TCGSL.CxTFNo = TCD.CxTFNo And TCGSL.TxnType = TCD.TxnType And TCGSL.SM_Prefix = TCD.SM_Prefix"
    tmp = tmp & " Inner Join Mst_GSL MGSL on  MGSL.GSLID = TCGSL.GSLID And MGSL.SM_Prefix = TCGSL.SM_Prefix"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
    tmp = tmp & " Inner Join Mst_Cmp MCMP on MCMP.CMPID = MGSL.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
    If Gen_DBType <> "MDB" Then
        tmp = tmp & " Where TCD.CxTFDate between '" & mStartDate & "' And '" & mEndDate & "' And MG.CloseDate is null"
    Else
        tmp = tmp & " Where TCD.CxTFDate between #" & mStartDate & "# And #" & mEndDate & "# And MG.CloseDate is null"
    End If
    tmp = tmp & wc & " group by MGSL.GodownID"
    Call TmpTable
    
    ProgressBar1.Value = 0
    ProgressBar1.Max = TmpRs.RecordCount + 1

    If TmpRs.RecordCount > 0 Then
        For I = 1 To TmpRs.RecordCount
            tmp = "Update Mst_Godown set YearMonth = '" & mStartDate & "', NoOfTransactions = " & TmpRs!NoOfTransaction & " Where GodownID = " & TmpRs!GodownID & " "
            Call Tmp1Table
            TmpRs.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
    ProgressBar1.Value = ProgressBar1.Max
    End If
End If

MsgBox "Done!!!"

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
For I = 2010 To Year(Now)
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

