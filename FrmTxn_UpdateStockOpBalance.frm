VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_UpdateStockOpBalance 
   Caption         =   "Update Stock Opening Balance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8820
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   15090
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6960
         TabIndex        =   25
         Top             =   120
         Width           =   8055
         Begin VB.CommandButton CmdGodownUltilizationProcess 
            Caption         =   "Godown Ultilization Update"
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
            Left            =   5160
            TabIndex        =   30
            Top             =   360
            Width           =   2760
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   360
            Left            =   720
            TabIndex        =   28
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
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
            Format          =   107675649
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   360
            Left            =   2775
            TabIndex        =   29
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
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
            Format          =   107675649
            CurrentDate     =   39884
         End
         Begin VB.Label Label10 
            Caption         =   "To"
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
            Left            =   2280
            TabIndex        =   27
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label9 
            Caption         =   "From"
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
            TabIndex        =   26
            Top             =   360
            Width           =   600
         End
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
         Left            =   7575
         TabIndex        =   24
         Top             =   8340
         Visible         =   0   'False
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
         Left            =   4695
         TabIndex        =   12
         Top             =   8340
         Width           =   1440
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
         Left            =   6135
         TabIndex        =   11
         Top             =   8340
         Width           =   1440
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
         ItemData        =   "FrmTxn_UpdateStockOpBalance.frx":0000
         Left            =   2595
         List            =   "FrmTxn_UpdateStockOpBalance.frx":0002
         TabIndex        =   10
         Top             =   375
         Width           =   1500
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
         TabIndex        =   9
         Top             =   8760
         Visible         =   0   'False
         Width           =   795
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
         Left            =   4320
         TabIndex        =   8
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
         ItemData        =   "FrmTxn_UpdateStockOpBalance.frx":0004
         Left            =   120
         List            =   "FrmTxn_UpdateStockOpBalance.frx":0006
         TabIndex        =   7
         Top             =   375
         Width           =   2415
      End
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
         TabIndex        =   6
         Top             =   795
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7065
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   12462
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame3 
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
         Left            =   120
         TabIndex        =   3
         Top             =   6645
         Visible         =   0   'False
         Width           =   5895
         Begin VB.TextBox TxtConfirmation 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   120
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   765
            Width           =   7725
         End
         Begin VB.Label Label8 
            Caption         =   "Enter the words 'All Checkpoint items are completed' in the box below and then press the 'process' key to proceed further"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   7800
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   7065
         Left            =   6960
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   12462
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         TabIndex        =   22
         Top             =   120
         Width           =   495
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
         TabIndex        =   21
         Top             =   8400
         Visible         =   0   'False
         Width           =   885
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
         Left            =   1027
         TabIndex        =   20
         Top             =   135
         Width           =   600
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
         Left            =   4335
         TabIndex        =   19
         Top             =   780
         Width           =   2355
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
         Left            =   6735
         TabIndex        =   18
         Top             =   780
         Width           =   90
      End
      Begin VB.Label Label1 
         Caption         =   "Check Points :"
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
         Left            =   150
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   $"FrmTxn_UpdateStockOpBalance.frx":0008
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   150
         TabIndex        =   16
         Top             =   1665
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Label Label5 
         Caption         =   $"FrmTxn_UpdateStockOpBalance.frx":00A3
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   150
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Label Label6 
         Caption         =   $"FrmTxn_UpdateStockOpBalance.frx":0191
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   150
         TabIndex        =   14
         Top             =   5280
         Visible         =   0   'False
         Width           =   6705
      End
   End
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   45
      TabIndex        =   0
      Top             =   8850
      Width           =   15090
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   200
         Left            =   75
         TabIndex        =   1
         Top             =   210
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmTxn_UpdateStockOpBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mToDate, mNoofDays As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mGenInvoiceLocation As Variant
Dim mCountLocation As Variant

Dim mTotalDays, mTotalMonths As Variant
Dim mWideStartDate, mWideEndDate As Variant
Dim mStartDate, mEndDate As Variant

Private Sub ChkCheckAll_Click()
If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
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

Private Sub CmdExit_Click()
Unload Me
End Sub

'Commented on 19 July 2013 for Insertion of PWR lots having deposit flag in Monthwise GSL Table from April to June 2013
'Private Sub CmdProcess_Click()
'
''On Error GoTo msgError
'
'ProgressBar1.Value = 0
'Call Grid_HeadDetails
'
'If CmbMonth.Text = "" Then
'   MsgBox " Blank Month not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox " Blank Year not allowed  !!!"
'   CmbYear.SetFocus
'   Exit Sub
'End If
'
'If MSHFlexGrid1.Rows < 3 Then
'   MsgBox " No Location populated in Grid  !!!"
'   Exit Sub
'End If
'
'If ChkSelection = False Then
'   MsgBox " No Location selected from Grid  !!!"
'   Exit Sub
'End If
'
'If LCase(Gen_UserRole) <> "power" Then
''   If LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
''      If mCountLocation = 5 Then
''         MsgBox "You can't select more than 5 location. Only 'Power' user can select all location!!!"
''         Exit Sub
''      End If
''   Else
''      If mCountLocation = 1 Then
''         MsgBox "You can't select more than 1 location. Only 'Power' user can select all location!!!"
''         Exit Sub
''      End If
''   End If
'End If
'
'mMonth = mMonthListindex + 1
'
'If mMonth < 10 Then
'   mMonth = "0" & mMonth
'End If
'
'
'mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
'
'mYear = Val(CmbYear.Text)
'mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
'
'mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")
'
'mNoofdays = ""
'
'If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
'   MsgBox "Present and Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'CmdProcess.Enabled = False
'
'mGenInvoiceLocation = ""
'
'If MSHFlexGrid1.Rows > 2 Then
'      mGenInvoiceLocation = "SM_Prefix in ("
'      For i = 1 To MSHFlexGrid1.Rows - 1
'          If MSHFlexGrid1.TextMatrix(i, 1) = "" Then Exit For
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
'
'
'
''Commented temprorly for Test server
'
''If ChkBOD = False Then
''   MsgBox " BOD pending for Location selected from Grid  !!!"
''   CmdProcess.Enabled = True
''   Exit Sub
''End If
'
'
'
''If CmbPreFix.Text <> "" Then
''   tmp = "Select Count(*) as 'TotalRows' from Txn_InvMonGSLDetail Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & mFromDate & "' And TFlag = 'O' "
'''Else
'''   tmp = "Select Count(*) as 'TotalRows' from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And TFlag = 'O' "
''
''  Call TmpTable
''
''  If TmpRs!TotalRows > 0 Then
''     If CmbPreFix.Text <> "" Then
''        MsgBox "Opening balance process for " & CmbPreFix.Text & " " & CmbMonth.Text & "-" & CmbYear.Text & "  month is already executed!!!"
'''     Else
'''        MsgBox "Opening balance process for this month is already executed!!!"
''     End If
''     Exit Sub
''  End If
''
''
''End If
'
'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
'   If TxtConfirmation.Text = "" Then
'      MsgBox "Blank Confirmation message not allowed!!!"
'      TxtConfirmation.SetFocus
'      CmdProcess.Enabled = True
'      Exit Sub
'   End If
'   If LCase(Trim(TxtConfirmation.Text)) <> "all checkpoint items are completed" Then
'      MsgBox "Confirmation message not correct!!!"
'      TxtConfirmation.SetFocus
'      CmdProcess.Enabled = True
'      Exit Sub
'   End If
'End If
'
'Gen_mTimeStamp = GetTimeStamp
'
''-- Create Temp Table with record of GSL
'MsgBox "Wait till next message!!!"
'
''Commented on 02 July 2011
'
''tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSL') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
''tmp = tmp & " drop table dbo.TempGSL"
''Call TmpTable
'
''Commented On 12 Nov 2010
''If CmbPreFix.Text <> "" Then
''Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)
''Else
''   Call TableTxn_CreateTempGSL(mFromDate, "", mGenInvoiceLocation)
''End If
'
'
'Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)
'
''-- Find Distinct location for those location which are having Withdral txn.
'
'tmp = " Select Distinct TCG.SM_Prefix"
'tmp = tmp & " From Txn_CxTF_GSL  TCG"
'tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
''tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'
''Commented on 12 Nov 2010 Parag
''If CmbPreFix.Text <> "" Then
''   tmp = tmp & " Where TCG.SM_Prefix = '" & mSM_Prefix & "' And TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' "
''Else
''   tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And  '" & Format(Date, Gen_DatFormat) & "'"
''End If
'
'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And " & mGenInvoiceLocation
'tmp = tmp & " And TCG." & mGenInvoiceLocation
'tmp = tmp & " Order By TCG.SM_Prefix"
'
'Call TmpTable
'
''If TmpRs.RecordCount > 0 Then
'   ProgressBar1.Value = 0
'   ProgressBar1.Max = TmpRs.RecordCount + 1
'   For i = 1 To TmpRs.RecordCount
'
'       '---fetch Withdrawal record for adding back to opening stock. which is > [To Date]
''
''       tmp = "Select TCG.GSLID,TCG.SM_Prefix,TCG.CxTFNo,TCD.CxTFDate,TCG.NoOfBags,TCG.DematWeight,TCG.TxnType"
''       tmp = tmp & " From Txn_CxTF_GSL  TCG"
''       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
''       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
''       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
''       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '" & mFromDate & "' And  '" & Format(Date, Gen_DatFormat) & "' "
''       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
''       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID,TCD.CxTFDate"
''
'
'       tmp = "Select TCG.GSLID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
'       tmp = tmp & " From Txn_CxTF_GSL  TCG"
'       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
'       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
'       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate > '" & mToDate & "' "
'       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'" 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
'       tmp = tmp & " Group By TCG.SM_Prefix,TCG.GSLID "
'       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID "
'
'
'       Call Tmp1Table
'
'       '--Update temp table with Withdrawal transaction record to get opening stock
'
'
'       Call TableTxn_TempGSL(" Where SM_Prefix = '" & TmpRs!SM_Prefix & "'")
'
'
'       If RsTxn_TempGSL.RecordCount > 0 Then
'          For j = 1 To TmpRs1.RecordCount
'              If RsTxn_TempGSL.RecordCount > 0 Then
'                 RsTxn_TempGSL.MoveFirst
'              End If
'
''              If TmpRs1!GSLID = 37 Or TmpRs1!GSLID = 663 Or TmpRs1!GSLID = 664 Or TmpRs1!GSLID = 665 Or TmpRs1!GSLID = 666 Or TmpRs1!GSLID = 674 Or TmpRs1!GSLID = 675 Or TmpRs1!GSLID = 676 Or TmpRs1!GSLID = 677 Or TmpRs1!GSLID = 678 Or TmpRs1!GSLID = 681 Or TmpRs1!GSLID = 682 Or TmpRs1!GSLID = 683 Or TmpRs1!GSLID = 684 Or TmpRs1!GSLID = 686 Or TmpRs1!GSLID = 687 Or TmpRs1!GSLID = 688 Or TmpRs1!GSLID = 730 Then
''                 a = a
''              End If
'
'              RsTxn_TempGSL.Find "GSLID=" & TmpRs1!GSLID
'              If Not RsTxn_TempGSL.EOF Then
'                 RsTxn_TempGSL!BalanceBags = RsTxn_TempGSL!BalanceBags + TmpRs1!NoofBags
'                 RsTxn_TempGSL!BalanceWeight = RsTxn_TempGSL!BalanceWeight + TmpRs1!DematWeight
'                 RsTxn_TempGSL!CDTFBags = RsTxn_TempGSL!CDTFBags - TmpRs1!NoofBags
'                 RsTxn_TempGSL!CDTFWeight = RsTxn_TempGSL!CDTFWeight - TmpRs1!DematWeight
'                 RsTxn_TempGSL!Flag = "O"
'                 RsTxn_TempGSL.Update
'              End If
'           TmpRs1.MoveNext
'          Next
'       End If
'       TmpRs.MoveNext
'       ProgressBar1.Value = ProgressBar1.Value + 1
'   Next
'
'
'   ProgressBar1.Value = 0
'
'   If TmpRs.RecordCount < 50 Then
'      ProgressBar1.Max = 1000
'   End If
'   '-- Delete All data from temp table having Flag='E'
'
'   tmp = "Delete from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where Flag = 'E' "
'   Call TmpTable
'
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
'   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Add YearMonth DateTime,BillingCycleID decimal(18,0),BillingUnit Varchar(1),Drno decimal(18,0),Rate decimal(18,4), "
'   tmp = tmp & " NoofDays decimal(18,0) ,InvoiceAmount decimal(18,2),TFLag Varchar(1),CxTFNo Decimal(18,0),CxTFDate DateTime,InvoiceDate DateTime,DateOfWithdrawal Datetime  "
'
'   Call TmpTable
'
'
'
'   '-- Insert all Temp data into GSLInvoice table for billing
'
''   tmp = " Insert into Txn_InvMonGSLDetail "
''   tmp = tmp & " (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
''   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,Tflag,GrossWt,KRDate,"
''   tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
''   tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo)"
''   tmp = tmp & " Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
''   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,'O',GrossWt,KRDate,"
''   tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
''   tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo"
''   tmp = tmp & " From TempGSL"
''
''   Call TmpTable
'
'   If Gen_DBType = "MDB" Then
'      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = #" & mFromDate & "#"
'   Else
'      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = '" & mFromDate & "'"
'   End If
'   Call TmpTable
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
''   tmp = "Drop Table TempGSL"
''   Call TmpTable
'
'
'   'YearMonth,BillingCycleID,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount
'
'   'ProgressBar1.Value = ProgressBar1.Value + 1
'   ' Copy GSL from transaction table for selected month for generation of billing
'
''   tmp = " Insert into Txn_InvMonGSLDetail (GSLID,SM_Prefix,TFlag,CxTFNo,CxTFDate,CMPID,"
'
'' Commented on 27 July 2011 for Updating Op Balance
'
''   tmp = " Insert into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " (GSLID,SM_Prefix,TFlag,CxTFNo,CxTFDate,CMPID,"
''   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
''   tmp = tmp & " ClientIDRow,BalanceBags,BalanceWeight,Flag,Flag1,GrossWt,"
''   tmp = tmp & " BagSize,CDTFNo,CDTFDate,KRDate,NContractID,YearMonth,G_UID,CreatedBy,CreatedDate,"
''   tmp = tmp & " UpdatedDate , UpdatedBy)"
''   tmp = tmp & " Select TCG.GSLID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTFNo,TCD.CxTFDate,TCD.CMPID,MGSL.GodownID,MGSL.StackNo,"
''   tmp = tmp & " MGSL.LotNo,MGSL.CommodityID,MGSL.ExchangeTypeID,"
''   tmp = tmp & " MGSL.ClientIDRow,TCG.NoOfBags,TCG.DematWeight,MGSL.Flag,MGSL.Flag1,MGSL.GrossWt,"
''   tmp = tmp & " MGSL.BagSize,MGSL.CDTFNo,MGSL.CDTFDate,MGSL.KRDate,MGSL.NContractID,'" & mFromDate & "',MGSL.G_UID,MGSL.CreatedBy,MGSL.CreatedDate,"
''   tmp = tmp & " MGSL.UpdatedDate , MGSL.UpdatedBy"
''   tmp = tmp & " From Txn_CxTF_GSL  TCG"
''   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix"
''   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
''   tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
''   tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
''   tmp = tmp & " Where MGSL.SpillageFlag = 0 And (TCG.BillFlag is Null or TCG.BillFlag = '')"
''   tmp = tmp & " and TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' "
''
'''Commented on 12 Nov 2010 Parag
'''   If CmbPreFix.Text <> "" Then
'''      tmp = tmp & " and TCG.SM_Prefix = '" & mSM_Prefix & "'"
'''   End If
''
''   tmp = tmp & " And TCG." & mGenInvoiceLocation
''
''   Call TmpTable
'
''   Commented on 22 March 2011 order given by Mayur Sir
''   tmp = "Select SM_Prefix,CxTFNo,TFlag,CDTFNo,CDTFDate,CxTFDate From TempGSL "
''   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And ExchangeTypeID = 1 "
''   tmp = tmp & " And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,CxTFDate) = 0 And "
''   tmp = tmp & " (SpillageFlag = 0 Or SpillageFlag is Null) And " & mGenInvoiceLocation & " "
''   tmp = tmp & " Group by SM_Prefix,CxTFNo,TFlag,CDTFNo,CDTFDate,CxTFDate"
''
''   Call TmpTable
''
''
''   For I = 1 To TmpRs.RecordCount
''
''       tmp = " Update TempGSL Set DateofWithdrawal = '" & TmpRs!CxtfDate & "'"
''       tmp = tmp & " Where CxTFNo = " & TmpRs!CdtfNo & "  And SM_Prefix = '" & TmpRs!SM_Prefix & "' And "
''       tmp = tmp & " YearMonth = '" & mFromDate & "' And TFlag = 'D' And ExchangeTypeID = 1 "
''       tmp = tmp & " And (ContractID = 0 Or ContractID is Null) And (SpillageFlag = 0 Or SpillageFlag is Null) "
''
''       Call Tmp1Table
''
''       TmpRs.MoveNext
''   Next
'
'   'Implemented on 03 Nov 2010
'   'Due to decimal place defination difference in GSL Master table and  GSL Invoice Detail Table
'
'   tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set BagSize = Round(Bagsize,0) "
'   Call TmpTable
'
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
'   ' Flag off all GSL which are used for selected month in transaction table so to avoid duplicate billing
'
'   tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillFlag = 'I'"
'   tmp = tmp & " Where  GSLDetailID in (Select TCG.GSLDetailID"
'   tmp = tmp & " From Txn_CxTF_GSL  TCG"
'   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix"
'   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
'   tmp = tmp & " where  (TCG.BillFlag is Null or TCG.BillFlag = '')  And"
'   tmp = tmp & " TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix) "
'
'   'tmp = tmp & " TCD.CxTFDate <= '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix)"
'
''Commented on 12 Nov 2010 Parag
''   If CmbPreFix.Text <> "" Then
''      tmp = tmp & " and Txn_CxTF_GSL.SM_Prefix = '" & mSM_Prefix & "'"
''   End If
'
'   tmp = tmp & " And Txn_CxTF_GSL." & mGenInvoiceLocation
'
'   Call TmpTable
'
'   ProgressBar1.Value = ProgressBar1.Value + 1
'
'   '--------------------------------------------------
'   '| A. Apply Special Rate to GSL Table.                |
'   '--------------------------------------------------
'
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingCycleID = "
''    tmp = tmp & " (Select a.BillingCycleID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
'
'    ' For Op and With
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
'    tmp = tmp & " (Select a.BillingCycleID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'For Depo
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
''    tmp = tmp & " (Select a.DepoBillingCycleID"
''    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
''    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingUnit = "
''    tmp = tmp & " (Select a.BillingUnit"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = "
'    tmp = tmp & " (Select a.BillingUnit"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Drno = "
''    tmp = tmp & " (Select a.SPRCID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate = "
''    tmp = tmp & " (Select a.Amount"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
''    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
''    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
''    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Flag1 = Null Where YearMonth = '" & mFromDate & "'"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Flag1 = 'S' Where DRNo is not null And YearMonth = '" & mFromDate & "' "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
''    tmp = tmp & " (Select  MSRCD.RatePerBag"
''    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD"
''    tmp = tmp & " Inner Join Mst_SpecialRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
''    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
''    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = Txn_InvMonGslDetail.GSLID"
''    tmp = tmp & " And TIMGD.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix And TIMGD.BillingUnit = 'B'"
''    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.BillingUnit = 'B' And Txn_InvMonGslDetail.Flag1 = 'S' "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''   '--------------------------------------------------
''   '| B. Apply  Base Rate to GSL Table.                |
''   '--------------------------------------------------
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingCycleID ="
''    tmp = tmp & " (Select a.BillingCycleID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null)  "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingUnit ="
''    tmp = tmp & " (Select a.Flag"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.DRNo ="
''    tmp = tmp & " (Select a.SRCID"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
''    tmp = tmp & " (Select a.RateOnQty"
''    tmp = tmp & " from  Txn_InvMonGslDetail b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
''    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Flag1 = 'N' Where DRNo is not null And YearMonth = '" & mFromDate & "' And Flag1 is null And BillingCycleID is not Null And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
''    tmp = tmp & " (Select MSRCD.RatePerBag"
''    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD"
''    tmp = tmp & " Inner Join Mst_StorageRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
''    tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
''    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
''    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
''    tmp = tmp & " And TIMGD.GSLID = Txn_InvMonGslDetail.GSLID And TIMGD.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
''    tmp = tmp & " And TIMGD.DRNo = Txn_InvMonGslDetail.DRNo"
''    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
''    tmp = tmp & " Where Txn_InvMonGslDetail.BillingUnit = 'B' And Txn_InvMonGslDetail.Flag1 = 'N'"
''    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''   '--------------------------------------------------
''   '| C. Apply  No Of Cycles to GSL Table as per Billing Unit and Transaction Flag.                |
''   '--------------------------------------------------
''
''    'Op Bal and Monthly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 4  And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Op Bal and Forthnightly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Op Bal and Weekly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''    'Op Bal and Daily Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
''    tmp = tmp & " And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Depo and Withd and Monthly Cycle
''    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag <> 'O'"
''    tmp = tmp & " And BillingCycleID = 4 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Depo and Withd and Daily Cycle
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Deposit & Weekly
''    mNoofdays = GetBillingCycleDays(2)
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Floor((4 - Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)"
'''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
'''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
'''
'''    Call TmpTable
''
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = 1 Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID in (2) And"
'''    tmp = tmp & " Floor ((4 - (Convert(Numeric, Day(CxTFDate)) / 7) + 1)) < 1"
'''
'''    Call TmpTable
''
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Deposit & FortNightly
''    mNoofdays = GetBillingCycleDays(3)
''
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Floor(((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1))"
'''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
'''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'''
'''    Call TmpTable
''
''
'''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1)) End "
'''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
'''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''
''    tmp = " Update Txn_InvMonGslDetail Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''    'Withdrawal & weekly
''    mNoofdays = GetBillingCycleDays(2)
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Withdrawal & Fortnightly
''    mNoofdays = GetBillingCycleDays(3)
''
''
''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Updating Invoice Amount And Invoice Date
''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceBags * NoOfDays * Rate, InvoiceDate = '" & mToDate & "' "
''    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
''    tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '" & mToDate & "' "
''    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
''    tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
'''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month(CxTFDate) = Month(CDTFDate) "
''
''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And DateDiff(MOnth,CDTFDate,CxTFDate) = 0 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
'
'
'    '==============================================
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Drno = "
'    tmp = tmp & " (Select a.SPRCID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
'    tmp = tmp & " (Select a.Amount"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
'    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
'    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = Null Where YearMonth = '" & mFromDate & "'"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'S' Where DRNo is not null And YearMonth = '" & mFromDate & "' "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'    tmp = tmp & " (Select  MSRCD.RatePerBag"
'    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'    tmp = tmp & " Inner Join Mst_SpecialRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
'    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
'    tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
'    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
''    tmp = tmp & " (Select  MSRCD.RatePerBag"
''    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
''    tmp = tmp & " Inner Join Mst_SpecialRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
''    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
''    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
''    tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
''    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
'   '--------------------------------------------------
'   '| B. Apply  Base Rate to GSL Table.                |
'   '--------------------------------------------------
'
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
'    tmp = tmp & " (Select a.BillingCycleID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
''    tmp = tmp & " (Select a.DepoBillingCycleID"
''    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
''    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
''    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
''    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
''    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
''    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
''    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
''    tmp = tmp & " a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
''    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit ="
'    tmp = tmp & " (Select a.Flag"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo ="
'    tmp = tmp & " (Select a.SRCID"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
'    tmp = tmp & " (Select a.RateOnQty"
'    tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
'    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
'    tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
'    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
'    tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'N' Where DRNo is not null And YearMonth = '" & mFromDate & "' And Flag1 is null And BillingCycleID is not Null And ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
'    tmp = tmp & " (Select MSRCD.RatePerBag"
'    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
'    tmp = tmp & " Inner Join Mst_StorageRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
'    tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
'    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
'    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
'    tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
'    tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
'    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
'    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
''    tmp = tmp & " (Select MSRCD.RatePerBag"
''    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
''    tmp = tmp & " Inner Join Mst_StorageRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
''    tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
''    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
''    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
''    tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
''    tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
''    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
''
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
'
'   '--------------------------------------------------
'   '| C. Apply  No Of Cycles to GSL Table as per Billing Unit and Transaction Flag.                |
'   '--------------------------------------------------
'
'    'Op Bal and Monthly Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 4  And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Op Bal and Forthnightly Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Op Bal and Weekly Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'    'Op Bal and Daily Cycle
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
'    tmp = tmp & " And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Depo and Withd and Monthly Cycle
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag <> 'O'"
''    tmp = tmp & " And BillingCycleID = 4 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Depo and Withd and Daily Cycle
''
''
'''    Commented on 16 March 2011 Parag
''
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null) "
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''   Call TmpTable
''
''
''
''''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  And DateDiff(Month,CxTFDate,DateofWithdrawal) = 0 "
'''
'''    'tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Day(DateOfWithdrawal)-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  And DateDiff(Month,CxTFDate,DateofWithdrawal) = 0 "
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''    Call TmpTable
''
''    'tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  " 'And DateofWithdrawal is Null "
''
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  " 'And DateofWithdrawal is Null "
''
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Withd and Daily Cycle
''
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Deposit & Weekly
''    mNoofdays = GetBillingCycleDays(2)
''
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Floor((4 - Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)"
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
'''
'''    Call TmpTable
''
''
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID in (2) And"
'''    tmp = tmp & " Floor ((4 - (Convert(Numeric, Day(CxTFDate)) / 7) + 1)) < 1"
'''
'''    Call TmpTable
''
''
''' Commented on 16th March 2011 722 PM Parag
''
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''
'''    Call TmpTable
''
''
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) " 'And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DateOfWithdrawal is null "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''
''
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor(((Case CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/" & mNoofdays & ",1)))) End) - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor(((Case CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/7,1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/7,1)))) End) - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0"
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Deposit & FortNightly
''    mNoofdays = GetBillingCycleDays(3)
''
''
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Floor(((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1))"
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
'''
'''    Call TmpTable
''
''
'''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1)) End "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''
''' Commented on 16th March 2011 722 PM Parag
''
'''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) "
''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
''
''
'''    Call TmpTable
''
''
''    'Commented on 21 March 2011
'''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor(((Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End) - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor(((Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End) - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''
'''    Call TmpTable
''
''
'''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
'''    tmp = tmp & " And Day(DateOfWithdrawal) <= 15 And Day(CxTFDate) <= 15 "
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''    Call TmpTable
'''
'''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 2 "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
'''    tmp = tmp & " And Day(DateOfWithdrawal) >15 And Day(CxTFDate) <= 15"
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''    Call TmpTable
'''
'''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
'''    tmp = tmp & " And Day(DateOfWithdrawal) <= 31 And Day(CxTFDate) > 15 "
'''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
'''
'''    Call TmpTable
''
''
'''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
'''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
'''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateofWithdrawal is null "
''
'''    If CmbPreFix.Text <> "" Then
'''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'''    End If
''
''
''    'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 2 "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
''    tmp = tmp & " And Day(CxTFDate) <= 15 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Case 1 Depo Date  > 15 Then No Of Cycle = 1
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
''    tmp = tmp & " And Day(CxTFDate) > 15 "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Withdrawal & weekly
''    mNoofdays = GetBillingCycleDays(2)
''
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) "
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''
''    Call TmpTable
''
''    ProgressBar1.Value = ProgressBar1.Value + 1
''
''    'Withdrawal & Fortnightly
''    mNoofdays = GetBillingCycleDays(3)
''
''
''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
''    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
''    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
''
''    If CmbPreFix.Text <> "" Then
''       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
''    End If
''
''    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    'Updating Invoice Amount And Invoice Date
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceBags * NoOfDays * Rate, InvoiceDate = '" & mToDate & "' "
'    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)"
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '" & mToDate & "' "
'    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
'    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
''    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month(CxTFDate) = Month(CDTFDate) "
'
'    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,CxTFDate) = 0 "
'
'    If CmbPreFix.Text <> "" Then
'       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
'    End If
'
'
'    Call TmpTable
'
'
'    ProgressBar1.Value = ProgressBar1.Value + 1
'
'
''    tmp = " Insert into Txn_InvMonGSLDetail "
''    tmp = tmp & " (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
''    tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,Tflag,GrossWt,KRDate,"
''    tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
''    tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate)"
''    tmp = tmp & " Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
''    'tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,'O',GrossWt,KRDate,"
''
''    tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,TFlag,GrossWt,KRDate,"
''    tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
''    tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate "
''    tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
''
''    Call TmpTable
'
''***************************************************************************************************************
''Commented on 03 Sept 2011 by Parag for Displaying in Grid and for updating using for loop multiple location
''***************************************************************************************************************
''''''''    tmp = "Select * from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And " & mGenInvoiceLocation & " And TFlag = 'O'"
''''''''    Call TmpTable
''''''''
''''''''
''''''''    tmp = "Select * from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where YearMonth = '" & mFromDate & "'  And " & mGenInvoiceLocation & " "
''''''''
''''''''    Call Tmp1Table
''''''''
''''''''    mCount = 0
''''''''    For I = 1 To TmpRs1.RecordCount
''''''''        mBGSLID = TmpRs1!BGSLID
''''''''a:        If IsNull(mBGSLID) = False Then
''''''''             If TmpRs.RecordCount > 0 Then
''''''''                TmpRs.MoveFirst
''''''''             End If
''''''''             TmpRs.Find "GSLID = " & mBGSLID 'TmpRs1!BGSLID
''''''''             If TmpRs.EOF = False Then
''''''''                GoTo l
''''''''             Else
''''''''                tmp = " Select BGSLID From Mst_GSL Where " & mGenInvoiceLocation & " And GSLID = " & mBGSLID & ""
''''''''                Call Tmp3Table
''''''''                mBGSLID = TmpRs3!BGSLID
''''''''                GoTo a
''''''''             End If
''''''''          End If
''''''''        If TmpRs.RecordCount > 0 Then
''''''''           TmpRs.MoveFirst
''''''''        End If
''''''''        TmpRs.Find "GSLID = " & TmpRs1!GSLID
''''''''
''''''''        If TmpRs.EOF = True Then
''''''''           mCount = mCount + 1
''''''''           TmpRs.AddNew
''''''''           TmpRs!GSLID = TmpRs1!GSLID
''''''''           TmpRs!SM_Prefix = TmpRs1!SM_Prefix
''''''''           TmpRs!CMPID = TmpRs1!CMPID
''''''''           TmpRs!CommodityID = TmpRs1!CommodityID
''''''''           TmpRs!GodownID = TmpRs1!GodownID
''''''''           TmpRs!StackNo = TmpRs1!StackNo
''''''''           TmpRs!LotNo = TmpRs1!LotNo
''''''''           TmpRs!BalanceBags = TmpRs1!BalanceBags
''''''''           TmpRs!BalanceWeight = TmpRs1!BalanceWeight
''''''''           TmpRs!ExchangeTypeID = TmpRs1!ExchangeTypeID
''''''''           TmpRs!ClientIDRow = TmpRs1!ClientIDRow
''''''''           TmpRs!Flag = TmpRs1!Flag
''''''''           TmpRs!Flag1 = TmpRs1!Flag1
''''''''           TmpRs!tFlag = TmpRs1!tFlag
''''''''           TmpRs!GrossWt = TmpRs1!GrossWt
''''''''           TmpRs!KRDate = TmpRs1!KRDate
''''''''           TmpRs!ContractID = TmpRs1!ContractID
''''''''           TmpRs!BagSize = TmpRs1!BagSize
''''''''           TmpRs!YearMonth = TmpRs1!YearMonth
''''''''           TmpRs!G_UID = TmpRs1!G_UID
''''''''           TmpRs!CreatedBy = TmpRs1!CreatedBy
''''''''           TmpRs!CreatedDate = TmpRs1!CreatedDate
''''''''           TmpRs!UpdatedDate = TmpRs1!UpdatedDate
''''''''           TmpRs!UpdatedBy = TmpRs1!UpdatedBy
''''''''           TmpRs!CDTFDate = TmpRs1!CDTFDate
''''''''           TmpRs!CdtfNo = TmpRs1!CdtfNo
''''''''           TmpRs!BillingCycleID = TmpRs1!BillingCycleID
''''''''           TmpRs!BillingUnit = TmpRs1!BillingUnit
''''''''           TmpRs!DRNo = TmpRs1!DRNo
''''''''           TmpRs!Rate = TmpRs1!Rate
''''''''           TmpRs!NoofDays = TmpRs1!NoofDays
''''''''           TmpRs!InvoiceAmount = TmpRs1!InvoiceAmount
''''''''           TmpRs!CxTFNo = TmpRs1!CxTFNo
''''''''           TmpRs!CxtfDate = TmpRs1!CxtfDate
''''''''           TmpRs!InvoiceDate = TmpRs1!InvoiceDate
''''''''           TmpRs.Update
''''''''        Else
''''''''           If TmpRs!BalanceBags <> TmpRs1!BalanceBags And TmpRs!BalanceWeight <> TmpRs1!BalanceWeight And TmpRs!InvoiceAmount <> TmpRs1!InvoiceAmount Then
''''''''              TmpRs!BalanceBags = TmpRs1!BalanceBags
''''''''              TmpRs!BalanceWeight = TmpRs1!BalanceWeight
'''''''''              TmpRs!BillingCycleID = TmpRs1!BillingCycleID
'''''''''              TmpRs!BillingUnit = TmpRs1!BillingUnit
'''''''''              TmpRs!DRNo = TmpRs1!DRNo
'''''''''              TmpRs!Rate = TmpRs1!Rate
'''''''''              TmpRs!NoofDays = TmpRs1!NoofDays
''''''''
'''''''''              If LCase(TmpRs1!BillingUnit) = "b" Then
'''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceBags * TmpRs1!Rate * TmpRs1!NoofDays
'''''''''              Else
'''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceWeight * TmpRs1!Rate * TmpRs1!NoofDays
'''''''''              End If
''''''''
''''''''              If LCase(TmpRs1!BillingUnit) = "b" Then
''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceBags * TmpRs!Rate * TmpRs!NoofDays
''''''''              Else
''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceWeight * TmpRs!Rate * TmpRs!NoofDays
''''''''              End If
''''''''
''''''''              TmpRs.Update
''''''''           ElseIf TmpRs!InvoiceAmount <> TmpRs1!InvoiceAmount Then
'''''''''              TmpRs!BillingCycleID = TmpRs1!BillingCycleID
'''''''''              TmpRs!BillingUnit = TmpRs1!BillingUnit
'''''''''              TmpRs!DRNo = TmpRs1!DRNo
'''''''''              TmpRs!Rate = TmpRs1!Rate
'''''''''              TmpRs!NoofDays = TmpRs1!NoofDays
''''''''
'''''''''              If LCase(TmpRs1!BillingUnit) = "b" Then
'''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceBags * TmpRs1!Rate * TmpRs1!NoofDays
'''''''''              Else
'''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceWeight * TmpRs1!Rate * TmpRs1!NoofDays
'''''''''              End If
''''''''
''''''''              If LCase(TmpRs1!BillingUnit) = "b" Then
''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceBags * TmpRs!Rate * TmpRs!NoofDays
''''''''              Else
''''''''                 TmpRs!InvoiceAmount = TmpRs1!BalanceWeight * TmpRs!Rate * TmpRs!NoofDays
''''''''              End If
''''''''
''''''''
''''''''              TmpRs.Update
''''''''           End If
''''''''        End If
''''''''        TmpRs1.MoveNext
''''''''l:    Next
''''''''    tmp = "Drop Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
''''''''    Call TmpTable
''''''''
'''''''''End If
''***************************************************************************************************************
''***************************************************************************************************************
'
'Mcount = 0
'
'tmp = "Select Distinct SM_Prefix from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where ExchangeTypeID = 1 And YearMonth = '" & mFromDate & "' Order By SM_Prefix " ' And " & mGenInvoiceLocation & " "
'Call Tmp4Table
'
'For mrec = 1 To TmpRs4.RecordCount
'
'    tmp = "Select * from Txn_InvMonGSLDetail Where ExchangeTypeID = 1 And YearMonth = '" & mFromDate & "' And TFlag = 'O' And SM_Prefix = '" & TmpRs4!SM_Prefix & "'" '" & mGenInvoiceLocation & " And TFlag = 'O'"
'    Call TmpTable
'
'
'    tmp = "Select * from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where ExchangeTypeID = 1 And YearMonth = '" & mFromDate & "'  And SM_Prefix = '" & TmpRs4!SM_Prefix & "'" 'And " & mGenInvoiceLocation & " "
'    Call Tmp1Table
'
''    mCount = 0
'    For i = 1 To TmpRs1.RecordCount
'        mBGSLID = TmpRs1!BGSLID
'a:        If IsNull(mBGSLID) = False Then
'             If TmpRs.RecordCount > 0 Then
'                TmpRs.MoveFirst
'             End If
'             TmpRs.Find "GSLID = " & mBGSLID 'TmpRs1!BGSLID
'             If TmpRs.EOF = False Then
'                GoTo L
'             Else
'                tmp = " Select BGSLID From Mst_GSL Where SM_Prefix = '" & TmpRs4!SM_Prefix & "' And GSLID = " & mBGSLID & "" '" & mGenInvoiceLocation & " And GSLID = " & mBGSLID & ""
'                Call Tmp3Table
'                mBGSLID = TmpRs3!BGSLID
'                GoTo a
'             End If
'          End If
'        If TmpRs.RecordCount > 0 Then
'           TmpRs.MoveFirst
'        End If
'        TmpRs.Find "GSLID = " & TmpRs1!GSLID
'
'        If TmpRs.EOF = True Then
'           Mcount = Mcount + 1
'           MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
'
'           TmpRs.AddNew
'           TmpRs!GSLID = TmpRs1!GSLID
'           TmpRs!SM_Prefix = TmpRs1!SM_Prefix
'           TmpRs!CMPID = TmpRs1!CMPID
'           TmpRs!CommodityID = TmpRs1!CommodityID
'           TmpRs!GodownID = TmpRs1!GodownID
'           TmpRs!StackNo = TmpRs1!StackNo
'           TmpRs!LotNo = TmpRs1!LotNo
'           TmpRs!BalanceBags = TmpRs1!BalanceBags
'           TmpRs!BalanceWeight = TmpRs1!BalanceWeight
'           TmpRs!ExchangeTypeID = TmpRs1!ExchangeTypeID
'           TmpRs!ClientIDRow = TmpRs1!ClientIDRow
'           TmpRs!Flag = TmpRs1!Flag
'           TmpRs!Flag1 = TmpRs1!Flag1
'           TmpRs!tFlag = TmpRs1!tFlag
'           TmpRs!GrossWt = TmpRs1!GrossWt
'           TmpRs!KRDate = TmpRs1!KRDate
'           TmpRs!ContractID = TmpRs1!ContractID
'           TmpRs!BagSize = TmpRs1!BagSize
'           TmpRs!yearmonth = TmpRs1!yearmonth
'           TmpRs!G_UID = TmpRs1!G_UID
'           TmpRs!CreatedBy = TmpRs1!CreatedBy
'           TmpRs!CreatedDate = TmpRs1!CreatedDate
'           TmpRs!UpdatedDate = TmpRs1!UpdatedDate
'           TmpRs!UpdatedBy = TmpRs1!UpdatedBy
'           TmpRs!CDTFDate = TmpRs1!CDTFDate
'           TmpRs!CdtfNo = TmpRs1!CdtfNo
'           TmpRs!BillingCycleID = TmpRs1!BillingCycleID
'           TmpRs!BillingUnit = TmpRs1!BillingUnit
'           TmpRs!DRNo = TmpRs1!DRNo
'           TmpRs!Rate = TmpRs1!Rate
'           TmpRs!NoofDays = TmpRs1!NoofDays
'           TmpRs!InvoiceAMount = TmpRs1!InvoiceAMount
'           TmpRs!CxTFNo = TmpRs1!CxTFNo
'           TmpRs!CxTFDate = TmpRs1!CxTFDate
'           TmpRs!InvoiceDate = TmpRs1!InvoiceDate
'           TmpRs.Update
'
''           MSHFlexGrid2.TextMatrix(Mcount, 0) = GetLocationFromCMP(TmpRs1!CMPID)
''           MSHFlexGrid2.TextMatrix(Mcount, 1) = TmpRs1!GSLID
''           MSHFlexGrid2.TextMatrix(Mcount, 2) = TmpRs1!SM_Prefix
''           MSHFlexGrid2.TextMatrix(Mcount, 3) = GetCMPName(TmpRs1!CMPID)
''           MSHFlexGrid2.TextMatrix(Mcount, 4) = GetCommodityName(TmpRs1!CommodityID)
''           MSHFlexGrid2.TextMatrix(Mcount, 5) = GetGodownName(TmpRs1!GodownID, "N")
''           MSHFlexGrid2.TextMatrix(Mcount, 6) = TmpRs1!StackNo
''           MSHFlexGrid2.TextMatrix(Mcount, 7) = TmpRs1!LotNo
''           MSHFlexGrid2.TextMatrix(Mcount, 8) = TmpRs1!BalanceBags
''           MSHFlexGrid2.TextMatrix(Mcount, 9) = TmpRs1!BalanceWeight
''           MSHFlexGrid2.TextMatrix(Mcount, 10) = "Non NCDEX" ' TmpRs1!ExchangeTypeID
''           MSHFlexGrid2.TextMatrix(Mcount, 11) = GetClientName(TmpRs1!ClientIDRow)
''           MSHFlexGrid2.TextMatrix(Mcount, 12) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
''           MSHFlexGrid2.TextMatrix(Mcount, 13) = IIf(IsNull(TmpRs1!Flag1), "", TmpRs1!Flag1)
''           MSHFlexGrid2.TextMatrix(Mcount, 14) = TmpRs1!tFlag
''           MSHFlexGrid2.TextMatrix(Mcount, 15) = IIf(IsNull(TmpRs1!ContractID), "", TmpRs1!ContractID)
''           MSHFlexGrid2.TextMatrix(Mcount, 16) = IIf(IsNull(TmpRs1!BagSize), "", TmpRs1!BagSize)
''           MSHFlexGrid2.TextMatrix(Mcount, 17) = TmpRs1!YearMonth
''           MSHFlexGrid2.TextMatrix(Mcount, 18) = IIf(IsNull(TmpRs1!CDTFDate), "", TmpRs1!CDTFDate)
''           MSHFlexGrid2.TextMatrix(Mcount, 19) = IIf(IsNull(TmpRs1!CdtfNo), "", TmpRs1!CdtfNo)
''           If IsNull(TmpRs1!BillingCycleID) = False Then
''              MSHFlexGrid2.TextMatrix(Mcount, 20) = IIf(IsNull(TmpRs1!BillingCycleID), "", GetBillingCycle(TmpRs1!BillingCycleID, "N"))
''           Else
''              MSHFlexGrid2.TextMatrix(Mcount, 20) = ""
''           End If
''           MSHFlexGrid2.TextMatrix(Mcount, 21) = IIf(IsNull(TmpRs1!BillingUnit), "", TmpRs1!BillingUnit)
''           MSHFlexGrid2.TextMatrix(Mcount, 22) = IIf(IsNull(TmpRs1!DRNo), "", TmpRs1!DRNo)
''           MSHFlexGrid2.TextMatrix(Mcount, 23) = IIf(IsNull(TmpRs1!Rate), "", TmpRs1!Rate)
''           MSHFlexGrid2.TextMatrix(Mcount, 24) = IIf(IsNull(TmpRs1!NoofDays), "", TmpRs1!NoofDays)
''           MSHFlexGrid2.TextMatrix(Mcount, 25) = IIf(IsNull(TmpRs1!InvoiceAmount), "", TmpRs1!InvoiceAmount)
''           MSHFlexGrid2.TextMatrix(Mcount, 26) = IIf(IsNull(TmpRs1!CxTFNo), "", TmpRs1!CxTFNo)
''           MSHFlexGrid2.TextMatrix(Mcount, 27) = IIf(IsNull(TmpRs1!CxtfDate), "", TmpRs1!CxtfDate)
''           MSHFlexGrid2.TextMatrix(Mcount, 28) = IIf(IsNull(TmpRs1!InvoiceDate), "", TmpRs1!InvoiceDate)
'        Else
'           If TmpRs!BalanceBags <> TmpRs1!BalanceBags And TmpRs!BalanceWeight <> TmpRs1!BalanceWeight And TmpRs!InvoiceAMount <> TmpRs1!InvoiceAMount Then
'              TmpRs!BalanceBags = TmpRs1!BalanceBags
'              TmpRs!BalanceWeight = TmpRs1!BalanceWeight
'              If LCase(TmpRs1!BillingUnit) = "b" Then
'                 TmpRs!InvoiceAMount = TmpRs1!BalanceBags * TmpRs!Rate * TmpRs!NoofDays
'              Else
'                 TmpRs!InvoiceAMount = TmpRs1!BalanceWeight * TmpRs!Rate * TmpRs!NoofDays
'              End If
'
'              TmpRs.Update
'           ElseIf TmpRs!InvoiceAMount <> TmpRs1!InvoiceAMount Then
'              If LCase(TmpRs1!BillingUnit) = "b" Then
'                 TmpRs!InvoiceAMount = TmpRs1!BalanceBags * TmpRs!Rate * TmpRs!NoofDays
'              Else
'                 TmpRs!InvoiceAMount = TmpRs1!BalanceWeight * TmpRs!Rate * TmpRs!NoofDays
'              End If
'              TmpRs.Update
'           End If
''         If TmpRs!BalanceBags <> TmpRs1!BalanceBags And TmpRs!BalanceWeight <> TmpRs1!BalanceWeight And TmpRs!InvoiceAmount <> TmpRs1!InvoiceAmount Then
''            MSHFlexGrid2.TextMatrix(Mcount, 0) = GetLocationFromCMP(TmpRs1!CMPID)
''            MSHFlexGrid2.TextMatrix(Mcount, 1) = TmpRs1!GSLID
''            MSHFlexGrid2.TextMatrix(Mcount, 2) = TmpRs1!SM_Prefix
''            MSHFlexGrid2.TextMatrix(Mcount, 3) = GetCMPName(TmpRs1!CMPID)
''            MSHFlexGrid2.TextMatrix(Mcount, 4) = GetCommodityName(TmpRs1!CommodityID)
''            MSHFlexGrid2.TextMatrix(Mcount, 5) = GetGodownName(TmpRs1!GodownID, "N")
''            MSHFlexGrid2.TextMatrix(Mcount, 6) = TmpRs1!StackNo
''            MSHFlexGrid2.TextMatrix(Mcount, 7) = TmpRs1!LotNo
''            MSHFlexGrid2.TextMatrix(Mcount, 8) = TmpRs1!BalanceBags
''            MSHFlexGrid2.TextMatrix(Mcount, 9) = TmpRs1!BalanceWeight
''            MSHFlexGrid2.TextMatrix(Mcount, 10) = "Non NCDEX" ' TmpRs1!ExchangeTypeID
''            MSHFlexGrid2.TextMatrix(Mcount, 11) = GetClientName(TmpRs1!ClientIDRow)
''            MSHFlexGrid2.TextMatrix(Mcount, 12) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
''            MSHFlexGrid2.TextMatrix(Mcount, 13) = IIf(IsNull(TmpRs1!Flag1), "", TmpRs1!Flag1)
''            MSHFlexGrid2.TextMatrix(Mcount, 14) = TmpRs1!tFlag
''            MSHFlexGrid2.TextMatrix(Mcount, 15) = IIf(IsNull(TmpRs1!ContractID), "", TmpRs1!ContractID)
''            MSHFlexGrid2.TextMatrix(Mcount, 16) = IIf(IsNull(TmpRs1!BagSize), "", TmpRs1!BagSize)
''            MSHFlexGrid2.TextMatrix(Mcount, 17) = TmpRs1!YearMonth
''            MSHFlexGrid2.TextMatrix(Mcount, 18) = IIf(IsNull(TmpRs1!CDTFDate), "", TmpRs1!CDTFDate)
''            MSHFlexGrid2.TextMatrix(Mcount, 19) = IIf(IsNull(TmpRs1!CdtfNo), "", TmpRs1!CdtfNo)
''            MSHFlexGrid2.TextMatrix(Mcount, 20) = IIf(IsNull(TmpRs1!BillingCycleID), "", GetBillingCycle(TmpRs1!BillingCycleID, "N"))
''            MSHFlexGrid2.TextMatrix(Mcount, 21) = IIf(IsNull(TmpRs1!BillingUnit), "", TmpRs1!BillingUnit)
''            MSHFlexGrid2.TextMatrix(Mcount, 22) = IIf(IsNull(TmpRs1!DRNo), "", TmpRs1!DRNo)
''            MSHFlexGrid2.TextMatrix(Mcount, 23) = IIf(IsNull(TmpRs1!Rate), "", TmpRs1!Rate)
''            MSHFlexGrid2.TextMatrix(Mcount, 24) = IIf(IsNull(TmpRs1!NoofDays), "", TmpRs1!NoofDays)
''            If LCase(TmpRs1!BillingUnit) = "b" Then
''               MSHFlexGrid2.TextMatrix(Mcount, 25) = TmpRs1!BalanceBags * TmpRs!Rate * TmpRs!NoofDays
''            Else
''               MSHFlexGrid2.TextMatrix(Mcount, 25) = TmpRs1!BalanceWeight * TmpRs!Rate * TmpRs!NoofDays
''            End If
'''            MSHFlexGrid2.TextMatrix(Mcount, 25) = IIf(IsNull(TmpRs1!InvoiceAmount), "", TmpRs1!InvoiceAmount)
''            MSHFlexGrid2.TextMatrix(Mcount, 26) = IIf(IsNull(TmpRs1!CxTFNo), "", TmpRs1!CxTFNo)
''            MSHFlexGrid2.TextMatrix(Mcount, 27) = IIf(IsNull(TmpRs1!CxtfDate), "", TmpRs1!CxtfDate)
''            MSHFlexGrid2.TextMatrix(Mcount, 28) = IIf(IsNull(TmpRs1!InvoiceDate), "", TmpRs1!InvoiceDate)
''
''            MSHFlexGrid2.Row = Mcount
''            For c = 0 To MSHFlexGrid2.Cols - 1
''                MSHFlexGrid2.Col = c
''                MSHFlexGrid2.CellBackColor = vbRed
''            Next
''         End If
'        End If
'       TmpRs1.MoveNext
'L:    Next
'    TmpRs4.MoveNext
'Next
'
'tmp = "Drop Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
'Call TmpTable
'
'ProgressBar1.Value = ProgressBar1.Max
'
'ChkCheckAll.Value = 0
''Call Grid_Head
'MsgBox "Done !!"
'
'CmdProcess.Enabled = True
'TxtConfirmation = ""
'
'Exit Sub
'
'msgError:
'
'    tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
'    tmp = tmp & " drop table dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
'    Call TmpTable
'
'    MsgBox "Some problem occured during process!!!"
'    TxtConfirmation = ""
'
'End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_StkOpBalanceDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_StkOpBalanceDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid2.Rows

For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid2.TextMatrix(I, C)
        End If
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = C
                
        If MSHFlexGrid2.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbRed
        ElseIf MSHFlexGrid2.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid2.CellBackColor = vbYellow Then
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

Private Sub CmdGodownUltilizationProcess_Click()


If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date Should Be Less Than To Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

ProgressBar1.Value = 0
ProgressBar1.Max = 100000

mWideStartDate = Format(TxtFromDate, "YYYY-MM-") & "01"
mWideEndDate = Format(TxtToDate, "YYYY-MM-") & "01"
mTotalMonths = DateDiff("M", mWideStartDate, mWideEndDate) + 1

mStartDate = mWideStartDate
mEndDate = mWideEndDate

For I = 0 To mTotalMonths
    
    mTotalDays = Day(DateSerial(Year(mStartDate), Month(mStartDate) + 1, 0))
    ''--1
    tmp = "Select Distinct GodownID from Txn_InvMonGSLDetail Where YEarMonth = '" & mStartDate & "' And GodownID is not null "
    Call TmpTable
    
    If TmpRs.RecordCount = 0 Then GoTo l
    
    mGodownIDs = ""
    For x = 1 To TmpRs.RecordCount
        If mGodownIDs = "" Then
           mGodownIDs = "(" & TmpRs!GodownID
        Else
           mGodownIDs = mGodownIDs & "," & TmpRs!GodownID
        End If
        TmpRs.MoveNext
    Next
    
    If mGodownIDs <> "" Then
       mGodownIDs = mGodownIDs & ")"
    End If
    
    
    
    tmp = " Select TIMGD.GodownID,(" & mTotalDays & " - Day(MG.StartDate)) + 1 'CapacityNoOfDay'"
    tmp = tmp & " From Txn_InvMonGSLDetail  TIMGD "
    tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID "
    tmp = tmp & " Where TIMGD.YearMonth = '" & mStartDate & "' And DateDiff(Month,MG.StartDate,TIMGD.YearMonth) = 0 "
    tmp = tmp & " And DateDiff(Month,MG.StartDate,isnull(MG.CloseDate,'2040-01-01')) > 0 And TIMGD.CapacityNoOfDay is null "
    tmp = tmp & " And (MG.CloseDate >= '" & mStartDate & "' Or MG.CloseDate is null) And TIMGD.GodownID in " & mGodownIDs & " "
    tmp = tmp & " Group By TIMGD.GodownID,MG.StartDate "
    
    Call TmpTable
    
    For x = 1 To TmpRs.RecordCount
    
        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
        tmp = tmp & " Where Txn_InvMonGSLDetail.GodownID = " & TmpRs!GodownID & " "
        tmp = tmp & " And Txn_InvMonGSLDetail.YearMonth = '" & mStartDate & "' And Txn_InvMonGSLDetail.CapacityNoOfDay is null"
        
        Call Tmp1Table
        
        TmpRs.MoveNext
    Next
    'tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.CapacityNoOfDay = 30 Where Txn_InvMonGSLDetail.GodownID = 3424 "
    'tmp = tmp & " And Txn_InvMonGSLDetail.YearMonth = '" & mStartDate & "' And Txn_InvMonGSLDetail.CapacityNoOfDay is null"
    'Call TmpTable
    
    ''--2
    tmp = " Select TIMGD.GodownID,(Day(MG.CloseDate)) 'CapacityNoOfDay' "
    tmp = tmp & " From Txn_InvMonGSLDetail  TIMGD "
    tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID "
    tmp = tmp & " Where TIMGD.YearMonth = '" & mStartDate & "' And DateDiff(Month,TIMGD.YearMonth,MG.CloseDate) = 0 "
    tmp = tmp & " And DateDiff(Month,MG.StartDate,isnull(MG.CloseDate,'2040-01-01')) > 0 "
    tmp = tmp & " And TIMGD.CapacityNoOfDay is null And (MG.CloseDate >= '" & mStartDate & "' Or MG.CloseDate is null) And TIMGD.GodownID in " & mGodownIDs & " "
    tmp = tmp & " Group By TIMGD.GodownID,MG.CloseDate"
    
    Call TmpTable
    
    
    For x = 1 To TmpRs.RecordCount
    
        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
        tmp = tmp & " Where Txn_InvMonGSLDetail.GodownID = " & TmpRs!GodownID & " "
        tmp = tmp & " And Txn_InvMonGSLDetail.YearMonth = '" & mStartDate & "' And Txn_InvMonGSLDetail.CapacityNoOfDay is null"
        
        Call Tmp1Table
        
        TmpRs.MoveNext
    Next
    
    ''--3
    tmp = " Select TIMGD.GodownID, " & mTotalDays & " 'CapacityNoOfDay'"
    tmp = tmp & " From Txn_InvMonGSLDetail  TIMGD"
    tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
    tmp = tmp & " Where TIMGD.YearMonth = '" & mStartDate & "' And DateDiff(Month,MG.StartDate,TIMGD.YearMonth) > 0"
    tmp = tmp & " And DateDiff(Month,MG.StartDate,isnull(MG.CloseDate,'2040-01-01')) > 0"
    tmp = tmp & " And TIMGD.CapacityNoOfDay is null And (MG.CloseDate >= '" & mStartDate & "' Or MG.CloseDate is null) And TIMGD.GodownID in " & mGodownIDs & ""
    tmp = tmp & " Group By TIMGD.GodownID,MG.CloseDate"
    Call TmpTable
    
    
    For x = 1 To TmpRs.RecordCount
    
        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
        tmp = tmp & " Where Txn_InvMonGSLDetail.GodownID = " & TmpRs!GodownID & " "
        tmp = tmp & " And Txn_InvMonGSLDetail.YearMonth = '" & mStartDate & "' And Txn_InvMonGSLDetail.CapacityNoOfDay is null"
        
        Call Tmp1Table
        
        TmpRs.MoveNext
    Next
    
    ''--4
    tmp = " Select TIMGD.GodownID,(Day(MG.CloseDate) - Day(MG.StartDate)) + 1 'CapacityNoOfDay'"
    tmp = tmp & " From Txn_InvMonGSLDetail  TIMGD"
    tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
    tmp = tmp & " Where TIMGD.YearMonth = '" & mStartDate & "'  And DateDiff(Month,MG.StartDate,MG.CloseDate) = 0"
    tmp = tmp & " And TIMGD.CapacityNoOfDay is null And MG.CloseDate is not null And MG.CloseDate >= '" & mStartDate & "' And TIMGD.GodownID in " & mGodownIDs & ""
    tmp = tmp & " Group By TIMGD.GodownID,MG.CloseDate,MG.StartDate"
    Call TmpTable
    
    
    For x = 1 To TmpRs.RecordCount
    
        tmp = " Update Txn_InvMonGSLDetail Set Txn_InvMonGSLDetail.CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
        tmp = tmp & " Where Txn_InvMonGSLDetail.GodownID = " & TmpRs!GodownID & " "
        tmp = tmp & " And Txn_InvMonGSLDetail.YearMonth = '" & mStartDate & "' And Txn_InvMonGSLDetail.CapacityNoOfDay is null"
        
        Call Tmp1Table
        
        TmpRs.MoveNext
    Next
    
    '--5
    tmp = "Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = " & mTotalDays & " Where TFlag = 'O' And Txn_InvMonGSLDetail.YearMonth = '" & mStartDate & "' And Txn_InvMonGSLDetail.WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
    Call TmpTable
    
    '--6
    tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = (" & mTotalDays & " - Day(Txn_InvMonGSLDetail.CxTFDate)) + 1  Where YearMonth = '" & mStartDate & "' And TFlag = 'D' And WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
    Call TmpTable
    
    '--7
    tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = (Day(Txn_InvMonGSLDetail.CxTFDate))  Where YearMonth = '" & mStartDate & "' And TFlag = 'W' And WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
    Call TmpTable
    
    '--8
    tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = -(Day(Txn_InvMonGSLDetail.CxTFDate))  Where YearMonth = '" & mStartDate & "' And TFlag = 'W' And DateDiff(Month,Txn_InvMonGSLDetail.CDTFDate,Txn_InvMonGSLDetail.CxTFDate) = 0 And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
    Call TmpTable
    
    '--9
    tmp = " Update Txn_InvMonGSLDetail Set WtQty = WtNoOfDay*BalanceWeight Where YearMonth = '" & mStartDate & "' And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
    Call TmpTable
    
l:    mStartDate = Format(DateSerial(Year(mStartDate), Month(mStartDate) + 1, 1), "YYYY-MM-DD")
    ProgressBar1.Value = ProgressBar1.Value + 1
Next I

ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!!"


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

If MSHFlexGrid1.Rows < 3 Then
   MsgBox " No Location populated in Grid  !!!"
   Exit Sub
End If

If ChkSelection = False Then
   MsgBox " No Location selected from Grid  !!!"
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

mNoofDays = ""

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

CmdProcess.Enabled = False

mGenInvoiceLocation = ""

If MSHFlexGrid1.Rows > 2 Then
      mGenInvoiceLocation = "SM_Prefix in ("
      For I = 1 To MSHFlexGrid1.Rows - 1
          If MSHFlexGrid1.TextMatrix(I, 1) = "" Then Exit For
          If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
             If mGenInvoiceLocation = "SM_Prefix in (" Then
                mGenInvoiceLocation = mGenInvoiceLocation & "'" & MSHFlexGrid1.TextMatrix(I, 2) & "'"
             Else
                mGenInvoiceLocation = mGenInvoiceLocation & ",'" & MSHFlexGrid1.TextMatrix(I, 2) & "' "
             End If
          End If
      Next
      mGenInvoiceLocation = mGenInvoiceLocation & " )"
End If



'Commented temprorly for Test server

'If ChkBOD = False Then
'   MsgBox " BOD pending for Location selected from Grid  !!!"
'   CmdProcess.Enabled = True
'   Exit Sub
'End If



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

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   If mShowDTD = False Then
      MsgBox "Please Complete all check points before Processing Opening Balance!!!"
'      CmdProcess.Enabled = True
      Exit Sub
   End If
   If TxtConfirmation.Text = "" Then
      MsgBox "Blank Confirmation message not allowed!!!"
      TxtConfirmation.SetFocus
'      CmdProcess.Enabled = True
      Exit Sub
   End If
End If

Gen_mTimeStamp = GetTimeStamp

'-- Create Temp Table with record of GSL
MsgBox "Wait till next message!!!"

'Commented on 02 July 2011

'tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSL') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
'tmp = tmp & " drop table dbo.TempGSL"
'Call TmpTable

'Commented On 12 Nov 2010
'If CmbPreFix.Text <> "" Then
'Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)
'Else
'   Call TableTxn_CreateTempGSL(mFromDate, "", mGenInvoiceLocation)
'End If


Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)

'-- Find Distinct location for those location which are having Withdral txn.

'tmp = " Select Distinct TCG.SM_Prefix"
'tmp = tmp & " From Txn_CxTF_GSL  TCG"
'tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
''tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'
''Commented on 12 Nov 2010 Parag
''If CmbPreFix.Text <> "" Then
''   tmp = tmp & " Where TCG.SM_Prefix = '" & mSM_Prefix & "' And TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' "
''Else
''   tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And  '" & Format(Date, Gen_DatFormat) & "'"
''End If
'
'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And " & mGenInvoiceLocation
'tmp = tmp & " And TCG." & mGenInvoiceLocation
'tmp = tmp & " Order By TCG.SM_Prefix"
'
'Call TmpTable
'
''If TmpRs.RecordCount > 0 Then
'   ProgressBar1.Value = 0
'   ProgressBar1.Max = TmpRs.RecordCount + 1
'   For i = 1 To TmpRs.RecordCount
'
'       '---fetch Withdrawal record for adding back to opening stock. which is > [To Date]
''
''       tmp = "Select TCG.GSLID,TCG.SM_Prefix,TCG.CxTFNo,TCD.CxTFDate,TCG.NoOfBags,TCG.DematWeight,TCG.TxnType"
''       tmp = tmp & " From Txn_CxTF_GSL  TCG"
''       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
''       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
''       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
''       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '" & mFromDate & "' And  '" & Format(Date, Gen_DatFormat) & "' "
''       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'"
''       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID,TCD.CxTFDate"
''
'
'       tmp = "Select TCG.GSLID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
'       tmp = tmp & " From Txn_CxTF_GSL  TCG"
'       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
'       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
'       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate > '" & mToDate & "' "
'       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'" 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
'       tmp = tmp & " Group By TCG.SM_Prefix,TCG.GSLID "
'       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID "
'
'
'       Call Tmp1Table
'
'       '--Update temp table with Withdrawal transaction record to get opening stock
'
'
'       Call TableTxn_TempGSL(" Where SM_Prefix = '" & TmpRs!SM_Prefix & "'")
'
'
'       If RsTxn_TempGSL.RecordCount > 0 Then
'          For j = 1 To TmpRs1.RecordCount
'              If RsTxn_TempGSL.RecordCount > 0 Then
'                 RsTxn_TempGSL.MoveFirst
'              End If
'
'              RsTxn_TempGSL.Find "GSLID=" & TmpRs1!GSLID
'              If Not RsTxn_TempGSL.EOF Then
'                 RsTxn_TempGSL!BalanceBags = RsTxn_TempGSL!BalanceBags + TmpRs1!NoofBags
'                 RsTxn_TempGSL!BalanceWeight = RsTxn_TempGSL!BalanceWeight + TmpRs1!DematWeight
'                 RsTxn_TempGSL!CDTFBags = RsTxn_TempGSL!CDTFBags - TmpRs1!NoofBags
'                 RsTxn_TempGSL!CDTFWeight = RsTxn_TempGSL!CDTFWeight - TmpRs1!DematWeight
'                 RsTxn_TempGSL!Flag = "O"
'                 RsTxn_TempGSL.Update
'              End If
'           TmpRs1.MoveNext
'          Next
'       End If
'
'
'       '--------Added to add Opening balance of Full Withdrawal GSLs for Reservation with 'P' TFlag
'
'       tmp = "Select TCG.GSLID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
'       tmp = tmp & " From Txn_CxTF_GSL  TCG"
'       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
'       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'       'tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate Between '2010-04-01' And  '" & Format(Date, Gen_DatFormat) & "' "
'       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' And  TCD.CxTFDate <= '" & mToDate & "' "
'       tmp = tmp & " And TCD.ExchangeTypeID=1 And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "' And isnull(MGSL.NContractID,0)>0" 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
'       tmp = tmp & " Group By TCG.SM_Prefix,TCG.GSLID "
'       tmp = tmp & " Order By TCG.SM_Prefix,TCG.GSLID "
'
'       Call Tmp1Table
'
'       '--Update temp table with Withdrawal transaction record to get opening stock
'
'       Call TableTxn_TempGSL(" Where SM_Prefix = '" & TmpRs!SM_Prefix & "'")
'
'       If RsTxn_TempGSL.RecordCount > 0 Then
'          For j = 1 To TmpRs1.RecordCount
'              If RsTxn_TempGSL.RecordCount > 0 Then
'                 RsTxn_TempGSL.MoveFirst
'              End If
'
'              RsTxn_TempGSL.Find "GSLID=" & TmpRs1!GSLID
'              If Not RsTxn_TempGSL.EOF Then
'                 RsTxn_TempGSL!BalanceBags = RsTxn_TempGSL!BalanceBags + TmpRs1!NoofBags
'                 RsTxn_TempGSL!BalanceWeight = RsTxn_TempGSL!BalanceWeight + TmpRs1!DematWeight
'                 RsTxn_TempGSL!CDTFBags = RsTxn_TempGSL!CDTFBags - TmpRs1!NoofBags
'                 RsTxn_TempGSL!CDTFWeight = RsTxn_TempGSL!CDTFWeight - TmpRs1!DematWeight
'                 RsTxn_TempGSL!Flag = "P"
'                 RsTxn_TempGSL.Update
'              End If
'           TmpRs1.MoveNext
'          Next
'       End If
'
'       TmpRs.MoveNext
'       ProgressBar1.Value = ProgressBar1.Value + 1
'   Next
   
   
   ProgressBar1.Value = 0
   
   'If TmpRs.RecordCount < 50 Then
   ProgressBar1.Max = 1000
   'End If
   
   '-- Delete All data from temp table having Flag='E'
   
   tmp = "Delete from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where Flag = 'E' "
   Call TmpTable
   
'   tmp = "Delete from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where Flag = 'E' And isnull(NContractID,0)=0"
'   Call TmpTable
   
   ProgressBar1.Value = ProgressBar1.Value + 1
   
   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Add YearMonth DateTime,BillingCycleID decimal(18,0),BillingUnit Varchar(1),Drno decimal(18,0),Rate decimal(18,4), "
   tmp = tmp & " NoofDays decimal(18,0) ,InvoiceAmount decimal(18,2),TFLag Varchar(1),CxTFNo Decimal(18,0),CxTFDate DateTime,InvoiceDate DateTime,DateOfWithdrawal Datetime ,MonG_UID Varchar(36),CommTrackFlag varchar(1) "
   
   Call TmpTable
   
   
   
   '-- Insert all Temp data into GSLInvoice table for billing
   
'   tmp = " Insert into Txn_InvMonGSLDetail "
'   tmp = tmp & " (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
'   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,Tflag,GrossWt,KRDate,"
'   tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
'   tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo)"
'   tmp = tmp & " Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
'   tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,'O',GrossWt,KRDate,"
'   tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
'   tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo"
'   tmp = tmp & " From TempGSL"
'
'   Call TmpTable
   
   If Gen_DBType = "MDB" Then
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = #" & mFromDate & "#"
   Else
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = '" & mFromDate & "'"
   End If
     
   Call TmpTable
   
   tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set MonG_UID = newID(),YearMonth = '" & mFromDate & "'"
   Call TmpTable
   
   ProgressBar1.Value = ProgressBar1.Value + 1
   
'   tmp = "Drop Table TempGSL"
'   Call TmpTable
   
   
   'YearMonth,BillingCycleID,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount
   
   'ProgressBar1.Value = ProgressBar1.Value + 1
   ' Copy GSL from transaction table for selected month for generation of billing
   
'   tmp = " Insert into Txn_InvMonGSLDetail (GSLID,SM_Prefix,TFlag,CxTFNo,CxTFDate,CMPID,"
   
'   tmp = " Insert into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " (GSLID,SM_Prefix,TFlag,CxTFNo,CxTFDate,CMPID,"
'   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
'   tmp = tmp & " ClientIDRow,BalanceBags,BalanceWeight,Flag,Flag1,GrossWt,"
'   tmp = tmp & " BagSize,CDTFNo,CDTFDate,KRDate,NContractID,YearMonth,G_UID,CreatedBy,CreatedDate,"
'   tmp = tmp & " UpdatedDate , UpdatedBy,KRNo,MonG_UID,MadeUpFlag)"
'   tmp = tmp & " Select TCG.GSLID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTFNo,TCD.CxTFDate,TCD.CMPID,MGSL.GodownID,MGSL.StackNo,"
'   tmp = tmp & " MGSL.LotNo,MGSL.CommodityID,MGSL.ExchangeTypeID,"
'   tmp = tmp & " MGSL.ClientIDRow,TCG.NoOfBags,TCG.DematWeight,MGSL.Flag,MGSL.Flag1,MGSL.GrossWt,"
'   tmp = tmp & " MGSL.BagSize,MGSL.CDTFNo,MGSL.CDTFDate,MGSL.KRDate,MGSL.NContractID,'" & mFromDate & "',MGSL.G_UID,MGSL.CreatedBy,MGSL.CreatedDate,"
'   tmp = tmp & " MGSL.UpdatedDate , MGSL.UpdatedBy,MGSL.KRNo,newid(),MGSL.MadeUpFlag"
'   tmp = tmp & " From Txn_CxTF_GSL  TCG"
'   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix"
'   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
'   tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'   tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
'
'   'Commented on 19 July 2012 after discussion with mayur sir due to leakage in stock data of NCDEX
'   'tmp = tmp & " Where MGSL.SpillageFlag = 0 And (TCG.BillFlag is Null or TCG.BillFlag = '')"
'
'   tmp = tmp & " Where MGSL.SpillageFlag = 0 "
'   tmp = tmp & " and TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' "
'
''Commented on 12 Nov 2010 Parag
''   If CmbPreFix.Text <> "" Then
''      tmp = tmp & " and TCG.SM_Prefix = '" & mSM_Prefix & "'"
''   End If
'
'   tmp = tmp & " And TCG." & mGenInvoiceLocation
'
'   Call TmpTable

   tmp = "Delete From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where TFlag = 'O'"
   Call TmpTable
   
   tmp = " Insert into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " (GSLID,SM_Prefix,TFlag,CxTFNo,CxTFDate,CMPID,"
   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
   tmp = tmp & " ClientIDRow,BalanceBags,BalanceWeight,Flag,Flag1,GrossWt,"
   tmp = tmp & " BagSize,CDTFNo,CDTFDate,KRDate,NContractID,YearMonth,G_UID,CreatedBy,CreatedDate,"
   tmp = tmp & " UpdatedDate , UpdatedBy,KRNo,MonG_UID,MadeUpFlag)"
   tmp = tmp & " Select MGSL.GSLID,MGSL.SM_Prefix,'D',MGSL.CDTFNo,MGSL.CDTFDate,MGSL.CMPID,MGSL.GodownID,MGSL.StackNo,"
   tmp = tmp & " MGSL.LotNo,MGSL.CommodityID,MGSL.ExchangeTypeID,"
   tmp = tmp & " MGSL.ClientIDRow,MGSL.DepositeBags,MGSL.DepositeWeight,MGSL.Flag,MGSL.Flag1,MGSL.GrossWt,"
   tmp = tmp & " MGSL.BagSize,MGSL.CDTFNo,MGSL.CDTFDate,MGSL.KRDate,MGSL.NContractID,'" & mFromDate & "',MGSL.G_UID,MGSL.CreatedBy,MGSL.CreatedDate,"
   tmp = tmp & " MGSL.UpdatedDate , MGSL.UpdatedBy, MGSL.KRNo, newid(), MGSL.MadeUpFlag"
   tmp = tmp & " From Mst_GSL MGSL "
   tmp = tmp & " Where MGSL.SpillageFlag = 0 And MGSL.CdtfNo < 0"
   tmp = tmp & " and MGSL.CDTFDate Between '" & mFromDate & "' And '" & mToDate & "'"
   tmp = tmp & " And MGSL." & mGenInvoiceLocation
   
   Call TmpTable
   
'   Commented on 22 March 2011 order given by Mayur Sir
'   tmp = "Select SM_Prefix,CxTFNo,TFlag,CDTFNo,CDTFDate,CxTFDate From TempGSL "
'   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And ExchangeTypeID = 1 "
'   tmp = tmp & " And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,CxTFDate) = 0 And "
'   tmp = tmp & " (SpillageFlag = 0 Or SpillageFlag is Null) And " & mGenInvoiceLocation & " "
'   tmp = tmp & " Group by SM_Prefix,CxTFNo,TFlag,CDTFNo,CDTFDate,CxTFDate"
'
'   Call TmpTable
'
'
'   For I = 1 To TmpRs.RecordCount
'
'       tmp = " Update TempGSL Set DateofWithdrawal = '" & TmpRs!CxtfDate & "'"
'       tmp = tmp & " Where CxTFNo = " & TmpRs!CdtfNo & "  And SM_Prefix = '" & TmpRs!SM_Prefix & "' And "
'       tmp = tmp & " YearMonth = '" & mFromDate & "' And TFlag = 'D' And ExchangeTypeID = 1 "
'       tmp = tmp & " And (ContractID = 0 Or ContractID is Null) And (SpillageFlag = 0 Or SpillageFlag is Null) "
'
'       Call Tmp1Table
'
'       TmpRs.MoveNext
'   Next
   
   'Implemented on 03 Nov 2010
   'Due to decimal place defination difference in GSL Master table and  GSL Invoice Detail Table
   
   
   tmp = "Select * from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set BagSize = Round(Bagsize,0) "
      Call TmpTable
       
      ProgressBar1.Value = ProgressBar1.Value + 1
       
       ' Flag off all GSL which are used for selected month in transaction table so to avoid duplicate billing
       
       
       
'       tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillFlag = 'I'"
'       tmp = tmp & " Where  GSLDetailID in (Select TCG.GSLDetailID"
'       tmp = tmp & " From Txn_CxTF_GSL  TCG"
'       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTFNo = TCD.CxTFNo And TCG.SM_Prefix = TCD.SM_Prefix"
'       tmp = tmp & " And TCG.TxnType = TCD.TxnType"
'       tmp = tmp & " where  (TCG.BillFlag is Null or TCG.BillFlag = '')  And"
'       tmp = tmp & " TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix) "
'
'       'tmp = tmp & " TCD.CxTFDate <= '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix)"
'
'    'Commented on 12 Nov 2010 Parag
'    '   If CmbPreFix.Text <> "" Then
'    '      tmp = tmp & " and Txn_CxTF_GSL.SM_Prefix = '" & mSM_Prefix & "'"
'    '   End If
'
'       tmp = tmp & " And Txn_CxTF_GSL." & mGenInvoiceLocation
'
'       Call TmpTable
          
       ProgressBar1.Value = ProgressBar1.Value + 1
       
       '--------------------------------------------------
       '| A. Apply Special Rate to GSL Table.                |
       '--------------------------------------------------
        
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingCycleID = "
    '    tmp = tmp & " (Select a.BillingCycleID"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
    '    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
    '    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
    '    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
    '    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    
        ' For Op and With
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
        tmp = tmp & " (Select a.BillingCycleID"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'For Depo
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = "
        tmp = tmp & " (Select a.DepoBillingCycleID"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount) "
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
           
        
        Call TmpTable
        
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingUnit = "
    '    tmp = tmp & " (Select a.BillingUnit"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
    '    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
    '    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
    '    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
    '    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    
    
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = "
        tmp = tmp & " (Select a.BillingUnit"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Drno = "
    '    tmp = tmp & " (Select a.SPRCID"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
    '    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
    '    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
    '    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
    '    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate = "
    '    tmp = tmp & " (Select a.Amount"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
    '    tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And b.GSLID = Txn_InvMonGslDetail.GSLID"
    '    tmp = tmp & " And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
    '    tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
    '    tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Flag1 = Null Where YearMonth = '" & mFromDate & "'"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Flag1 = 'S' Where DRNo is not null And YearMonth = '" & mFromDate & "' "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
    '    tmp = tmp & " (Select  MSRCD.RatePerBag"
    '    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD"
    '    tmp = tmp & " Inner Join Mst_SpecialRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
    '    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
    '    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = Txn_InvMonGslDetail.GSLID"
    '    tmp = tmp & " And TIMGD.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix And TIMGD.BillingUnit = 'B'"
    '    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
    '    tmp = tmp & " Where Txn_InvMonGslDetail.BillingUnit = 'B' And Txn_InvMonGslDetail.Flag1 = 'S' "
    '    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '   '--------------------------------------------------
    '   '| B. Apply  Base Rate to GSL Table.                |
    '   '--------------------------------------------------
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingCycleID ="
    '    tmp = tmp & " (Select a.BillingCycleID"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
    '    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
    '    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
    '    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null)  "
    '    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.BillingUnit ="
    '    tmp = tmp & " (Select a.Flag"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
    '    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
    '    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
    '    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
    '    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.DRNo ="
    '    tmp = tmp & " (Select a.SRCID"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
    '    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
    '    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
    '    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
    '    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
    '
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
    '    tmp = tmp & " (Select a.RateOnQty"
    '    tmp = tmp & " from  Txn_InvMonGslDetail b"
    '    tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
    '    tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
    '    tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
    '    tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
    '    tmp = tmp & " And b.GSLID = Txn_InvMonGslDetail.GSLID And b.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
    '    tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
    '    tmp = tmp & " Where (Txn_InvMonGslDetail.Flag1 is Null) "
    '    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Flag1 = 'N' Where DRNo is not null And YearMonth = '" & mFromDate & "' And Flag1 is null And BillingCycleID is not Null And ExchangeTypeID = 1 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set Txn_InvMonGslDetail.Rate ="
    '    tmp = tmp & " (Select MSRCD.RatePerBag"
    '    tmp = tmp & " From Txn_InvMonGSLDetail TIMGD"
    '    tmp = tmp & " Inner Join Mst_StorageRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
    '    tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
    '    tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
    '    tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
    '    tmp = tmp & " And TIMGD.GSLID = Txn_InvMonGslDetail.GSLID And TIMGD.SM_Prefix = Txn_InvMonGslDetail.SM_Prefix"
    '    tmp = tmp & " And TIMGD.DRNo = Txn_InvMonGslDetail.DRNo"
    '    tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
    '    tmp = tmp & " Where Txn_InvMonGslDetail.BillingUnit = 'B' And Txn_InvMonGslDetail.Flag1 = 'N'"
    '    tmp = tmp & " And Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and Txn_InvMonGslDetail.SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '   '--------------------------------------------------
    '   '| C. Apply  No Of Cycles to GSL Table as per Billing Unit and Transaction Flag.                |
    '   '--------------------------------------------------
    '
    '    'Op Bal and Monthly Cycle
    '    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
    '    tmp = tmp & " And BillingCycleID = 4  And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Op Bal and Forthnightly Cycle
    '    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
    '    tmp = tmp & " And BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Op Bal and Weekly Cycle
    '    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
    '    tmp = tmp & " And BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '    'Op Bal and Daily Cycle
    '    tmp = " Update Txn_InvMonGslDetail Set NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
    '    tmp = tmp & " And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Depo and Withd and Monthly Cycle
    '    tmp = " Update Txn_InvMonGslDetail Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag <> 'O'"
    '    tmp = tmp & " And BillingCycleID = 4 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Depo and Withd and Daily Cycle
    '    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Deposit & Weekly
    '    mNoofdays = GetBillingCycleDays(2)
    '
    ''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Floor((4 - Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)"
    ''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
    ''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
    ''
    ''    Call TmpTable
    '
    '
    ''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = 1 Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID in (2) And"
    ''    tmp = tmp & " Floor ((4 - (Convert(Numeric, Day(CxTFDate)) / 7) + 1)) < 1"
    ''
    ''    Call TmpTable
    '
    '
    '    tmp = "Update Txn_InvMonGslDetail Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
    '    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Deposit & FortNightly
    '    mNoofdays = GetBillingCycleDays(3)
    '
    '
    ''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Floor(((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1))"
    ''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
    ''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
    ''
    ''    Call TmpTable
    '
    '
    ''    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1)) End "
    ''    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
    ''    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
    '
    '
    '    tmp = " Update Txn_InvMonGslDetail Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
    '    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '    'Withdrawal & weekly
    '    mNoofdays = GetBillingCycleDays(2)
    '
    '    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
    '    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And Txn_InvMonGslDetail.BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Withdrawal & Fortnightly
    '    mNoofdays = GetBillingCycleDays(3)
    '
    '
    '    tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End "
    '    tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
    '    tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And Txn_InvMonGslDetail.BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    'Updating Invoice Amount And Invoice Date
    '    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceBags * NoOfDays * Rate, InvoiceDate = '" & mToDate & "' "
    '    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
    '    tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    '    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '" & mToDate & "' "
    '    tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
    '    tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
    '
    '
    '    ProgressBar1.Value = ProgressBar1.Value + 1
    '
    ''    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month(CxTFDate) = Month(CDTFDate) "
    '
    '    tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And DateDiff(MOnth,CDTFDate,CxTFDate) = 0 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
        
        
        '==============================================
        
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Drno = "
        tmp = tmp & " (Select a.SPRCID"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
        tmp = tmp & " (Select a.Amount"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid "
        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "'  And a.SPRCID is not null And a.FLag = 'A' And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.GSLID,b.SM_Prefix,"
        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
         
        
        ProgressBar1.Value = ProgressBar1.Value + 1
         
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = Null Where YearMonth = '" & mFromDate & "'"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
            
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'S' Where DRNo is not null And YearMonth = '" & mFromDate & "' "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
    
        
        ProgressBar1.Value = ProgressBar1.Value + 1
              
              
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
        tmp = tmp & " (Select  MSRCD.RatePerBag"
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
        tmp = tmp & " Inner Join Mst_SpecialRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
        tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W')"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
       
        
        ProgressBar1.Value = ProgressBar1.Value + 1
       
       
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
        tmp = tmp & " (Select  MSRCD.RatePerBag"
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
        tmp = tmp & " Inner Join Mst_SpecialRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SPRCID"
        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID"
        tmp = tmp & " And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix And TIMGD.BillingUnit = 'B'"
        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'S' "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D')"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
       
        
        ProgressBar1.Value = ProgressBar1.Value + 1
       
       
       '--------------------------------------------------
       '| B. Apply  Base Rate to GSL Table.                |
       '--------------------------------------------------
       
       
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
        tmp = tmp & " (Select a.BillingCycleID"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID ="
        tmp = tmp & " (Select a.DepoBillingCycleID"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null "
        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
        tmp = tmp & " a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null)  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit ="
        tmp = tmp & " (Select a.Flag"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo ="
        tmp = tmp & " (Select a.SRCID"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
        
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
    
        
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate = "
        tmp = tmp & " (Select a.RateOnQty"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
        tmp = tmp & " Where a.SRCID is not null And (b.ContractID = 0 Or b.ContractID is Null) And b.YearMonth = '" & mFromDate & "' And a.ExchangeTypeID = 1 And a.DefaultRate = 1 And b.Flag1 is null"
        tmp = tmp & " And b.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.SM_Prefix,b.GSLID,b.ExchangeTypeID,"
        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,a.RateOnQty)"
        tmp = tmp & " Where (TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 is Null) "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
    
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set Flag1 = 'N' Where DRNo is not null And YearMonth = '" & mFromDate & "' And Flag1 is null And BillingCycleID is not Null And ExchangeTypeID = 1 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
       
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
        tmp = tmp & " (Select MSRCD.RatePerBag"
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
        tmp = tmp & " Inner Join Mst_StorageRateCardDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
        tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
        tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O','W') "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Rate ="
        tmp = tmp & " (Select MSRCD.RatePerBag"
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TIMGD"
        tmp = tmp & " Inner Join Mst_StorageRateCardDepoDetail MSRCD On TIMGD.DRNo = MSRCD.SRCID"
        tmp = tmp & " Inner Join Mst_StorageRateCard MSRC On MSRCD.SRCID = MSRC.SRCID"
        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And TIMGD.DRNo is Not null And TIMGD.BagSize >= MSRCD.FromBagSize And"
        tmp = tmp & " TIMGD.BagSize <= MSRCD.ToBagSize  And TIMGD.BillingUnit = 'B' And TIMGD.Flag1 = 'N' And  MSRC.DefaultRate = 1"
        tmp = tmp & " And TIMGD.GSLID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GSLID And TIMGD.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " And TIMGD.DRNo = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DRNo"
        tmp = tmp & " Group By TIMGD.GSLID,TIMGD.SM_Prefix,TIMGD.BagSize,MSRCD.FromBagSize,MSRCD.ToBagSize,MSRCD.RatePerBag)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingUnit = 'B' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Flag1 = 'N' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('D') "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "'  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
        
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
       '--------------------------------------------------
       '| C. Apply  No Of Cycles to GSL Table as per Billing Unit and Transaction Flag.                |
       '--------------------------------------------------
    
        'Op Bal and Monthly Cycle
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
        tmp = tmp & " And BillingCycleID = 4  And (ContractID = 0 Or ContractID is Null)"
    
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
    
        Call TmpTable
        
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Op Bal and Forthnightly Cycle
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
        tmp = tmp & " And BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
        
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Op Bal and Weekly Cycle
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
        tmp = tmp & " And BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
        
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        'Op Bal and Daily Cycle
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
        tmp = tmp & " And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
         
        ProgressBar1.Value = ProgressBar1.Value + 1
         
        'Depo and Withd and Monthly Cycle
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag <> 'O'"
        tmp = tmp & " And BillingCycleID = 4 And (ContractID = 0 Or ContractID is Null)"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
        
        ProgressBar1.Value = ProgressBar1.Value + 1
         
        'Depo and Withd and Daily Cycle
        
        
    '    Commented on 16 March 2011 Parag
    
    '    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null) "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '   Call TmpTable
        
        
        
    ''    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  And DateDiff(Month,CxTFDate,DateofWithdrawal) = 0 "
    '
    '    'tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Day(DateOfWithdrawal)-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  And DateDiff(Month,CxTFDate,DateofWithdrawal) = 0 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
        
        'tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 30-Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  " 'And DateofWithdrawal is Null "
        
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = " & mTotalDays & " - Day(CxTFDate) + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)  " 'And DateofWithdrawal is Null "
        
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        'Withd and Daily Cycle
        
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Day(CxTFDate) Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null)"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
    
        
        
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Deposit & Weekly
        mNoofDays = GetBillingCycleDays(2)
                 
    '    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Floor((4 - Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)"
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
    '
    '    Call TmpTable
           
            
    '    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And TFlag = 'D' And BillingCycleID in (2) And"
    '    tmp = tmp & " Floor ((4 - (Convert(Numeric, Day(CxTFDate)) / 7) + 1)) < 1"
    '
    '    Call TmpTable
            
            
    ' Commented on 16th March 2011 722 PM Parag
    
    '    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null)"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
            
            
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ") + 1)) End "
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) " 'And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".DateOfWithdrawal is null "
    
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
            
           
            
    '    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays =  Case Floor(((Case CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/" & mNoofdays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/" & mNoofdays & ",1)))) End) - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1))  When 0 then 1 else Floor(((Case CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/7,1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(DateOfWithdrawal))/7,1)))) End) - (Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ") + 1)) End "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
        
        ProgressBar1.Value = ProgressBar1.Value + 1
            
        'Deposit & FortNightly
        mNoofDays = GetBillingCycleDays(3)
            
                 
    '    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Floor(((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1))"
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
    '
    '    Call TmpTable
           
           
    '    tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day(CxTFDate))/7) + 1)) End "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
           
           
    ' Commented on 16th March 2011 722 PM Parag
           
    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) "
        
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
           
           
    '    Call TmpTable
           
           
        'Commented on 21 March 2011
    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor(((Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End) - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor(((Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofdays & ",1)))) End) - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '
    '    Call TmpTable
        
        
    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
    '    tmp = tmp & " And Day(DateOfWithdrawal) <= 15 And Day(CxTFDate) <= 15 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 2 "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
    '    tmp = tmp & " And Day(DateOfWithdrawal) >15 And Day(CxTFDate) <= 15"
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
    '
    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1  And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,DateOfWithdrawal) = 0 "
    '    tmp = tmp & " And Day(DateOfWithdrawal) <= 31 And Day(CxTFDate) > 15 "
    '
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
    '
    '    Call TmpTable
            
            
    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day(CxTFDate))/" & mNoofdays & ",1)) - 1) End "
    '    tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
    '    tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) And DateofWithdrawal is null "
        
    '    If CmbPreFix.Text <> "" Then
    '       tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
    '    End If
           
           
        'Case 1 Depo Date  <= 15 Then No Of Cycle = 2
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 2 "
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
        tmp = tmp & " And Day(CxTFDate) <= 15 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
           
        Call TmpTable
              
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Case 1 Depo Date  > 15 Then No Of Cycle = 1
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = 1 "
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'D' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null) " 'And DateofWithdrawal is null "
        tmp = tmp & " And Day(CxTFDate) > 15 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
           
        Call TmpTable
              
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Withdrawal & weekly
        mNoofDays = GetBillingCycleDays(2)
                   
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) End "
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 "
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 2 And (ContractID = 0 Or ContractID is Null) "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        
        Call TmpTable
        
        ProgressBar1.Value = ProgressBar1.Value + 1
           
        'Withdrawal & Fortnightly
        mNoofDays = GetBillingCycleDays(3)
        
          
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = Case CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day(CxTFDate))/" & mNoofDays & ",1)))) End "
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1"
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag = 'W' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".BillingCycleID = 3 And (ContractID = 0 Or ContractID is Null)"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
            
        ProgressBar1.Value = ProgressBar1.Value + 1
        
        'Updating Invoice Amount And Invoice Date
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceBags * NoOfDays * Rate, InvoiceDate = '" & mToDate & "' "
        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null)"
            
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
            
        Call TmpTable
        
          
        ProgressBar1.Value = ProgressBar1.Value + 1
    
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = BalanceWeight * NoOfDays * Rate,InvoiceDate =  '" & mToDate & "' "
        tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
        tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
                
        Call TmpTable
                    
        ProgressBar1.Value = ProgressBar1.Value + 1
        
    '    tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month(CxTFDate) = Month(CDTFDate) "
        
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And DateDiff(Month,CDTFDate,CxTFDate) = 0 "
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        tmp = " UPDATE TI SET TI.CommTrackFlag = MC.CommTrackFlag  From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " TI"
        tmp = tmp & " inner join Mst_Commodity MC on MC.CommodityID=TI.CommodityID"
        tmp = tmp & " Where TI.ExchangeTypeID = 2 and TI.YearMonth = '" & mFromDate & "'"

        Call TmpTable
                    
        ProgressBar1.Value = ProgressBar1.Value + 1
                
        tmp = " Insert into Txn_InvMonGSLDetail "
        tmp = tmp & " (GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
        tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,Tflag,GrossWt,KRDate,"
        tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
        tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate,KRNo,ProcessedBy,MonG_UID,CommTrackFlag,MadeUpFlag)"
        tmp = tmp & " Select GSLID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
        'tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,'O',GrossWt,KRDate,"
        
        tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientIDRow,Flag,Flag1,TFlag,GrossWt,KRDate,"
        tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
        tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTFNo,CxTFDate,InvoiceDate,KRNo,'" & Gen_UserLoginID & "',MonG_UID,CommTrackFlag,MadeUpFlag "
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
    
        Call TmpTable
       
        tmp = " Update Txn_InvMonGSLDetail Set OldBalanceWeight = BalanceWeight Where BillingUnit = 'Q' And BalanceWeight < 0 And YearMonth= '" & mFromDate & "'"
        Call TmpTable
    
        tmp = " Update Txn_InvMonGSLDetail Set BalanceWeight = (BagSize * BalanceBags)/1000 Where BillingUnit = 'Q' And BalanceWeight < 0 And isnull(ContractID,0)=0 And YearMonth= '" & mFromDate & "'"
        Call TmpTable
        
        tmp = " Update Txn_InvMonGSLDetail Set InvoiceAmount = BalanceWeight*Rate*NoofDays Where BillingUnit = 'Q' And OldBalanceWeight < 0 And YearMonth= '" & mFromDate & "'"
        Call TmpTable
        
        tmp = " Select SM_Prefix from Mst_Location"
        tmp = tmp & " Where SM_Prefix not in "
        tmp = tmp & " (Select Distinct SM_Prefix From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ") "
        tmp = tmp & " And " & mGenInvoiceLocation & " "
        
        Call TmpTable
        
        For I = 1 To TmpRs.RecordCount
            
            tmp = " Insert into Txn_InvMonGSLDetail "
            tmp = tmp & " (GSLID,SM_Prefix,CMPID,StackNo,LotNo, "
            tmp = tmp & " YearMonth,G_UID,CreatedBy,CreatedDate,ProcessedBy "
            tmp = tmp & " )"
            tmp = tmp & " Values (-1,'" & TmpRs!SM_Prefix & "',-1,'Dummy-P','Dummy-P','" & mFromDate & "',newid(),'" & Gen_UserLoginID & "','" & Now() & "','" & Gen_UserLoginID & "')"
            
            Call Tmp1Table
            TmpRs.MoveNext
        
        Next
    Else
        
        tmp = "Select SM_prefix from Mst_Location Where " & mGenInvoiceLocation & " "
        Call TmpTable
        
        For I = 1 To TmpRs.RecordCount
            
            tmp = " Insert into Txn_InvMonGSLDetail "
            tmp = tmp & " (GSLID,SM_Prefix,CMPID,StackNo,LotNo, "
            tmp = tmp & " YearMonth,G_UID,CreatedBy,CreatedDate,ProcessedBy "
            tmp = tmp & " )"
            tmp = tmp & " Values (-1,'" & TmpRs!SM_Prefix & "',-1,'Dummy-P','Dummy-P','" & mFromDate & "',newid(),'" & Gen_UserLoginID & "','" & Now() & "','" & Gen_UserLoginID & "')"
            
            Call Tmp1Table
            
            TmpRs.MoveNext
        Next
    End If
    
    tmp = "Drop Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
    Call TmpTable
    
'End If


ProgressBar1.Value = ProgressBar1.Max

ChkCheckAll.Value = 0
Call Grid_Head
MsgBox "Done !!"

CmdProcess.Enabled = True

'TxtConfirmation = ""
'TxtRateCard.Text = ""
'TxtDTD.Text = ""
'TxtGSLMapping.Text = ""
'TxtConfirmation.Locked = True

Exit Sub

msgError:
    
    tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
    tmp = tmp & " drop table dbo.TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
    Call TmpTable
    
    'MsgBox "Some problem occured during process!!!"
    
    MsgBox Err.Description, vbExclamation, "Error"
    
    TxtConfirmation = ""
    TxtRateCard.Text = ""
    TxtDTD.Text = ""
    TxtGSLMapping.Text = ""
    TxtConfirmation.Locked = True

End Sub

Private Sub CmdShowLocation_Click()

Call Grid_Head
Call Grid_HeadDetails
ChkCheckAll.Value = 0
TxtConfirmation = ""

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

'mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")


mFromDate = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")

mYear = Val(CmbYear.Text)
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

mToDate = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")


If CDate(Format(mFromDate, "yyyymm")) >= Format(Date, "yyyymm") Then
   MsgBox "Present and Future Month not Allowed !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

Dim mChooseLocation As Variant

mChooseLocation = ""


'If Gen_DBType = "MDB" Then
'   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = #" & mFromDate & "# "
'Else
'   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' "
'End If


'If Gen_DBType = "MDB" Then
''   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = #" & mFromDate & "# "
'   tmp = "Select Distinct MGSL.SM_Prefix"
'   tmp = tmp & " From Mst_GSL MGSL --On TCG.GSLID = MGSL.GSLID And TCG.SM_Prefix = MGSL.SM_Prefix"
'   tmp = tmp & " Where MGSL.SpillageFlag = 0 And MGSL.CdtfNo < 0"
'   tmp = tmp & " and MGSL.CDTFDate Between #" & mFromDate & "# And #" & mToDate & "#"
'
'
'Else
''   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' "
'
'   tmp = "Select Distinct MGSL.SM_Prefix"
'   tmp = tmp & " From Mst_GSL MGSL "
'   tmp = tmp & " Where MGSL.SpillageFlag = 0 And MGSL.CdtfNo < 0"
'   tmp = tmp & " and MGSL.CDTFDate Between '" & mFromDate & "' And '" & mToDate & "' And MGSL.SM_Prefix Not in (   Select Distinct TGSL.SM_Prefix"
'   tmp = tmp & " From Txn_InvMonGSLDetail TGSL"
'   tmp = tmp & " Where  TGSL.CdtfNo < 0 And TGSL.TFlag = 'D'"
'   tmp = tmp & " and TGSL.CDTFDate Between '" & mFromDate & "' And '" & mToDate & "')"
'
'End If
'
'Call TmpTable
'
'For i = 1 To TmpRs.RecordCount
'    If mChooseLocation = "" Then
'       mChooseLocation = "(" & "'" & TmpRs!SM_Prefix & "'"
'    Else
'       mChooseLocation = mChooseLocation & "," & "'" & TmpRs!SM_Prefix & "'"
'    End If
'
'    If i = TmpRs.RecordCount Then
'       mChooseLocation = mChooseLocation & ")"
'    End If
'    TmpRs.MoveNext
'Next


'Commented on 15 March 2011 744 PM Parag
'tmp = "Select Distinct ML.LocationName,ML.SM_Prefix  "
'tmp = tmp & " From Mst_Location ML "
'tmp = tmp & " Left Join Mst_CMP  MC On ML.LocationID = MC.LocationID "
'tmp = tmp & " Left Join Mst_GSL MG On MC.CMPID = MG.CMPID "
'tmp = tmp & " Where ML.SM_Prefix not in "
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " (Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = #" & mFromDate & "#) "
'Else
'   tmp = tmp & " (Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "') "
'End If
'tmp = tmp & " Order By ML.LocationName "


'tmp = "Select Distinct ML.LocationName,ML.SM_Prefix  "
'tmp = tmp & " From Mst_Location ML "
'tmp = tmp & " Left Join Mst_CMP  MC On ML.LocationID = MC.LocationID "
'tmp = tmp & " Left Join Mst_GSL MG On MC.CMPID = MG.CMPID "
'
''If mChooseLocation <> "" Then
''   tmp = tmp & " Where ML.SM_Prefix not in " & mChooseLocation & ""
''End If
''
''If Gen_WcLocation <> "" Then
''   If mChooseLocation = "" Then
''      tmp = tmp & " Where ML." & Gen_WcLocation
''   Else
''      tmp = tmp & " And ML." & Gen_WcLocation
''   End If
''End If
'
'tmp = tmp & " Order By ML.LocationName "

tmp = "Select Distinct ML.LocationName,MGSL.SM_Prefix"
tmp = tmp & " From Mst_GSL MGSL "
tmp = tmp & " Inner Join Mst_Location ML On MGSL.SM_Prefix = ML.SM_Prefix "
tmp = tmp & " Where MGSL.SpillageFlag = 0 And MGSL.CdtfNo < 0 "
tmp = tmp & " and MGSL.CDTFDate Between '" & mFromDate & "' And '" & mToDate & "' And MGSL.SM_Prefix Not in (   Select Distinct TGSL.SM_Prefix"
tmp = tmp & " From Txn_InvMonGSLDetail TGSL"
tmp = tmp & " Where  TGSL.CdtfNo < 0 And TGSL.TFlag = 'D'"
tmp = tmp & " and TGSL.CDTFDate Between '" & mFromDate & "' And '" & mToDate & "')"


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

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

If LCase(Gen_UserRole) <> "power" Then
   Unload Me
End If

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If


Call Grid_Head
Call Grid_HeadDetails

Frame1.Enabled = True
'CmdDelete.Visible = False
ChkCheckAll.Visible = True
'If LCase(Gen_UserRole) = "power" Then
''   CmdDelete.Visible = True
'   ChkCheckAll.Visible = True
'End If
mCountLocation = 0
TxtConfirmation = ""
ProgressBar1.Value = 0
TxtFromDate.Value = Date
TxtToDate.Value = Date

End Sub


Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If
CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   mLocationID = 0
   mSM_Prefix = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbPreFix.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
mLocationName = RsMst_Location!LocationName
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

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
Private Sub Grid_HeadDetails()
With MSHFlexGrid2
     .Clear
     .Rows = 2
     .Cols = 29 '26 '27 '23 '25
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(0) = 500

     .TextMatrix(0, 0) = "Location"
     .TextMatrix(0, 1) = "GSL ID "
     .TextMatrix(0, 2) = "SM Prefix"
     .TextMatrix(0, 3) = "CMP "
     .TextMatrix(0, 4) = "Commodity "
     .TextMatrix(0, 5) = "Godown "
     .TextMatrix(0, 6) = "Stack No "
     .TextMatrix(0, 7) = "Lot No "
     .TextMatrix(0, 8) = "Balance Bags"
     .TextMatrix(0, 9) = "Balance Weight"
     .TextMatrix(0, 10) = "Exchange Type"
     .TextMatrix(0, 11) = "Client " 'IDRow"
     .TextMatrix(0, 12) = "Flag "
     .TextMatrix(0, 13) = "Flag1 "
     .TextMatrix(0, 14) = "TFlag "
     .TextMatrix(0, 15) = "Contract ID "
     .TextMatrix(0, 16) = "Bag Size "
     .TextMatrix(0, 17) = "Year Month "
     .TextMatrix(0, 18) = "CDTF Date "
     .TextMatrix(0, 19) = "CDTF No"
     .TextMatrix(0, 20) = "Billing Cycle" 'ID"
     .TextMatrix(0, 21) = "Billing Unit"
     .TextMatrix(0, 22) = "DR No"
     .TextMatrix(0, 23) = "Rate "
     .TextMatrix(0, 24) = "No of Days"
     .TextMatrix(0, 25) = "Invoice Amount"
     .TextMatrix(0, 26) = "CxTF No"
     .TextMatrix(0, 27) = "CxTF Date"
     .TextMatrix(0, 28) = "Invoice Date"
        
     .ColWidth(0) = 1600
     .ColWidth(1) = 1000
     .ColWidth(2) = 1000
          
End With
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


Private Sub MSHFlexGrid1_Click()
'
'headcmg


If LCase(Gen_UserRole) <> "power" Then
   If LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
      If mCountLocation = 5 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.Col) = strUnChecked Then
            MsgBox "You can't select more than 5 location. Only 'Power' user can select all location!!!"
            Exit Sub
         End If
      End If
   Else
      If mCountLocation = 1 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.Col) = strUnChecked Then
            MsgBox "You can't select more than 1 location." ' Only 'Power' user can select all location!!!"
            Exit Sub
         End If
      End If
   End If
End If
With MSHFlexGrid1
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        .Col = 0
        If .TextMatrix(.row, .Col) = strUnChecked Then
           .TextMatrix(.row, .Col) = strChecked
           mCountLocation = mCountLocation + 1
        Else
           .TextMatrix(.row, .Col) = strUnChecked
           mCountLocation = mCountLocation - 1
        End If
     End If
End With

End Sub
