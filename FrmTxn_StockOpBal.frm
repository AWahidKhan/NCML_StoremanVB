VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_StockOpBal 
   Caption         =   "Stock Opennig Balance"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9735
   ScaleWidth      =   10695
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   75
      TabIndex        =   17
      Top             =   8775
      Width           =   15090
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   75
         TabIndex        =   18
         Top             =   210
         Width           =   14940
         _ExtentX        =   26353
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8820
      Left            =   75
      TabIndex        =   0
      Top             =   -30
      Width           =   15090
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
         Left            =   6120
         TabIndex        =   32
         Top             =   360
         Width           =   1800
      End
      Begin VB.CommandButton CmdShowDTD 
         Caption         =   "Show DTD Report"
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
         Height          =   585
         Left            =   8400
         TabIndex        =   9
         Top             =   7320
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton CmdShowRateCard 
         Caption         =   "Show Rate Card Report"
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
         Height          =   585
         Left            =   8400
         TabIndex        =   7
         Top             =   4440
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton CmdShowGSL 
         Caption         =   "Show GSL Mapping"
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
         Height          =   585
         Left            =   8400
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox TxtDTD 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   10080
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   7320
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.TextBox TxtRateCard 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   10080
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   4440
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.TextBox TxtGSLMapping 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   480
         Left            =   10080
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2235
         Visible         =   0   'False
         Width           =   4845
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
         Height          =   1650
         Left            =   120
         TabIndex        =   23
         Top             =   6645
         Visible         =   0   'False
         Width           =   8055
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
            TabIndex        =   11
            Top             =   765
            Width           =   7845
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
            TabIndex        =   28
            Top             =   240
            Width           =   7800
         End
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
         TabIndex        =   4
         Top             =   795
         Width           =   1335
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
         ItemData        =   "FrmTxn_StockOpBal.frx":0000
         Left            =   120
         List            =   "FrmTxn_StockOpBal.frx":0002
         TabIndex        =   1
         Top             =   375
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
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1800
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
         TabIndex        =   15
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
         ItemData        =   "FrmTxn_StockOpBal.frx":0004
         Left            =   2595
         List            =   "FrmTxn_StockOpBal.frx":0006
         TabIndex        =   2
         Top             =   375
         Width           =   1500
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
         Left            =   6615
         TabIndex        =   13
         Top             =   8340
         Width           =   1560
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
         Left            =   5055
         TabIndex        =   12
         Top             =   8340
         Width           =   1560
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5505
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   9710
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label11 
         Caption         =   "Enter ""DTD Report Completed"" in below box and press tab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8355
         TabIndex        =   31
         Top             =   6705
         Visible         =   0   'False
         Width           =   6465
      End
      Begin VB.Label Label10 
         Caption         =   "Enter ""Rate Card Report Completed"" in below box and press tab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   30
         Top             =   4080
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Label Label9 
         Caption         =   "Enter ""GSL Mapping Completed"" in below box and press tab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   29
         Top             =   1815
         Visible         =   0   'False
         Width           =   6465
      End
      Begin VB.Label Label6 
         Caption         =   $"FrmTxn_StockOpBal.frx":0008
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   8310
         TabIndex        =   27
         Top             =   5160
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Label Label5 
         Caption         =   $"FrmTxn_StockOpBal.frx":014C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8310
         TabIndex        =   26
         Top             =   2760
         Visible         =   0   'False
         Width           =   6705
      End
      Begin VB.Label Label4 
         Caption         =   $"FrmTxn_StockOpBal.frx":023A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   8310
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   6705
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
         Left            =   8310
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
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
         TabIndex        =   22
         Top             =   780
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
         Left            =   4335
         TabIndex        =   21
         Top             =   780
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
         Left            =   1027
         TabIndex        =   19
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
         TabIndex        =   16
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
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmTxn_StockOpBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth, mFromDate, mMonthListindex, mSM_Prefix, mTotalDays, mToDate, mNoofDays As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mGenInvoiceLocation As Variant
Dim mCountLocation As Variant
Dim mShowDTD As Boolean

Dim mGodownIDs As Variant

Private Sub ChkCheckAll_Click()
If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If


Dim ans As Variant
ans = MsgBox("Do you want to select only Closed Locations . Click 'Yes' for so and 'No' to select all", vbYesNo)
If ans = vbYes Then
   If ChkCheckAll.Value = vbChecked Then
      For I = 1 To MSHFlexGrid1.Rows - 2
          MSHFlexGrid1.row = I
          MSHFlexGrid1.Col = 0
    '       MSHFlexGrid1.TextMatrix(i, 0) = strChecked
          If MSHFlexGrid1.TextMatrix(I, 3) = "" Then
             MSHFlexGrid1.TextMatrix(I, 0) = strChecked
          End If
      Next
      CmdShowGSL.Enabled = True
   Else
      For I = 1 To MSHFlexGrid1.Rows - 2
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
      Next
      CmdShowGSL.Enabled = False
   End If
Else
   If ChkCheckAll.Value = vbChecked Then
      For I = 1 To MSHFlexGrid1.Rows - 2
          MSHFlexGrid1.row = I
          MSHFlexGrid1.Col = 0
    '       MSHFlexGrid1.TextMatrix(i, 0) = strChecked
          If MSHFlexGrid1.TextMatrix(I, 3) <> "" Then
             If LCase(MSHFlexGrid1.TextMatrix(I, 5)) <> "yes" Then
                If Format(MSHFlexGrid1.TextMatrix(I, 3), "yyyymm") > Format(mToDate, "yyyymm") Then
                   MSHFlexGrid1.TextMatrix(I, 0) = strChecked
                ElseIf Format(MSHFlexGrid1.TextMatrix(I, 4), "yyyymmdd") = Format(mToDate, "yyyymmdd") Then
                   MSHFlexGrid1.TextMatrix(I, 0) = strChecked
                End If
             End If
          Else
             MSHFlexGrid1.TextMatrix(I, 0) = strChecked
          End If
      Next
      CmdShowGSL.Enabled = True
   Else
      For I = 1 To MSHFlexGrid1.Rows - 2
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
      Next
      CmdShowGSL.Enabled = False
   End If
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

If ChkBOD = False Then
   MsgBox " BOD pending for Location selected from Grid  !!!"
   CmdProcess.Enabled = True
   Exit Sub
End If


If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   If mShowDTD = False Then
      MsgBox "Please Complete all check points before Processing Opening Balance!!!"
      Exit Sub
   End If
   If TxtConfirmation.Text = "" Then
      MsgBox "Blank Confirmation message not allowed!!!"
      TxtConfirmation.SetFocus
      Exit Sub
   End If
End If

Gen_mTimeStamp = GetTimeStamp

'-- Create Temp Table with record of GSL
 MsgBox "Wait till next message!!!"



Call TableTxn_CreateTempGSL(mFromDate, mSM_Prefix, mGenInvoiceLocation)

'-- Find Distinct location for those location which are having Withdral txn.

tmp = " Alter table TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " Add Mst_GSL_ID_New int null;"
Call TmpTable

tmp = " Update TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " Set Mst_GSL_ID_New = Mst_GSL_ID"
Call TmpTable

tmp = " Alter table TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & " drop column Mst_GSL_ID;"
Call TmpTable

tmp = " EXEC sp_rename 'TempGSl" & Gen_UserLoginID & Gen_mTimeStamp & ".Mst_GSL_ID_New', 'Mst_GSL_ID', 'COLUMN';"
Call TmpTable



tmp = " Select Distinct TCG.SM_Prefix"
tmp = tmp & " From Txn_CxTF_GSL  TCG"
tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' " 'And " & mGenInvoiceLocation
tmp = tmp & " And TCG." & mGenInvoiceLocation
tmp = tmp & " Order By TCG.SM_Prefix"

Call TmpTable

   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount + 1
   For I = 1 To TmpRs.RecordCount
       
       '---fetch Withdrawal record for adding back to opening stock. which is > [To Date]

       tmp = "Select TCG.Mst_GSL_ID,TCG.SM_Prefix,sum(TCG.NoOfBags)'NoOfBags' ,sum(TCG.DematWeight)'DematWeight' "
       tmp = tmp & " From Txn_CxTF_GSL  TCG"
       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix And TCG.TxnType = TCD.TxnType"
       tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate > '" & mToDate & "' "
       tmp = tmp & " And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "'" 'And Month(MGSL.CDTFDate) <> Month(TCD.CxTFDate) "
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
                 RsTxn_TempGSL!Status = "O"
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
       tmp = tmp & " Where TCG.TxnType = 'W' And  TCD.CxTFDate >= '" & mFromDate & "' And  TCD.CxTFDate <= '" & mToDate & "' "
       tmp = tmp & " And TCD.ExchangeTypeID=1 And TCG.SM_Prefix = '" & TmpRs!SM_Prefix & "' And isnull(MGSL.NContractID,0)>0"
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
                 RsTxn_TempGSL!Status = "P"
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
  
   tmp = "Delete from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where Status = 'E' And isnull(NContractID,0)=0"
   Call TmpTable
   
   ProgressBar1.Value = ProgressBar1.Value + 1
   
   
   ''-- Alter Add With three Column Named WtNoOfDay,WtQty,CapacityNoOfDay 03 March 2016 -- Ullhas Patil
   
   
   tmp = "Alter Table TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Add YearMonth DateTime,Drno decimal(18,0),"
   tmp = tmp & " InvoiceAmount decimal(18,2),TFLag Varchar(1),CxTFNo Decimal(18,0),CxTFDate DateTime,InvoiceDate DateTime,DateOfWithdrawal Datetime ,MonG_UID Varchar(36),CommTrackFlag varchar(1),GodownYearlyBillingFlg varchar(1),CommodityYearlyBillingFlg varchar(1), "
   tmp = tmp & " WtNoOfDay Decimal(18,0),CapacityNoOfDay Decimal(18,0) "
   Call TmpTable
   
   
   '-- Insert all Temp data into GSLInvoice table for billing
   
  
   If Gen_DBType = "MDB" Then
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = #" & mFromDate & "#"
   Else
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TFlag = 'O',YearMonth = '" & mFromDate & "'"
   End If
     
   Call TmpTable
   
   tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set MonG_UID = newID(),YearMonth = '" & mFromDate & "'"
   Call TmpTable
   
   ProgressBar1.Value = ProgressBar1.Value + 1
   
   ' Copy GSL from transaction table for selected month for generation of billing
   
   
   tmp = " Insert into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " (Mst_GSL_ID,SM_Prefix,TFlag,CxTF_DetailsID,CxTFDate,CMPID,"
   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
   tmp = tmp & " ClientID,BalanceBags,BalanceWeight,Status,Flag1,GrossWt,"
   tmp = tmp & " BagSize,CDTFNo,CDTFDate,KRDate,NContractID,YearMonth,G_UID,CreatedBy,CreatedDate,"
   tmp = tmp & " UpdatedDate , UpdatedBy,KRNo,MonG_UID,MadeUpFlag,GodownYearlyBillingFlg,CommodityYearlyBillingFlg,RateCardType,BillingCycleID,BillingUnit,Rate,RateForDailyPnL,DepConDID,NoofDays,GSLID,SpillageFlag)"
   tmp = tmp & " Select TCG.Mst_GSL_ID,TCG.SM_Prefix,TCG.TxnType,TCG.CxTF_DetailsID,TCD.CxTFDate,TCD.CMPID,MGSL.GodownID,MGSL.StackNo,"
   tmp = tmp & " MGSL.LotNo,MGSL.CommodityID,MGSL.ExchangeTypeID,"
   tmp = tmp & " MGSL.ClientID,TCG.NoOfBags,TCG.DematWeight,MGSL.Status,MGSL.Flag1,MGSL.GrossWt,"
   tmp = tmp & " MGSL.BagSize,MGSL.CDTFNo,MGSL.CDTFDate,MGSL.KRDate,MGSL.NContractID,'" & mFromDate & "',MGSL.G_UID,MGSL.CreatedBy,MGSL.CreatedDate,"
   tmp = tmp & " MGSL.UpdatedDate , MGSL.UpdatedBy,MGSL.KRNo,newid(),MGSL.MadeUpFlag,MGod.YearlyBillingFlg 'GodownYearlyBillingFlg',MComm.YearlyBillingFlg 'CommodityYearlyBillingFlg',TCG.RateCardType,TCG.BillingCycleID,TCG.BillingUnit,TCG.Rate,TCG.RateForDailyPnL,TCG.DepConDID,TCG.NoofDays,MGSL.GSLID,MGSL.SpillageFlag"
   tmp = tmp & " From Txn_CxTF_GSL  TCG"
   tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix"
   tmp = tmp & " And TCG.TxnType = TCD.TxnType"
   tmp = tmp & " Inner Join Mst_GSL MGSL On TCG.Mst_GSL_ID = MGSL.Mst_GSL_ID And TCG.SM_Prefix = MGSL.SM_Prefix"
   tmp = tmp & " Inner Join Mst_CMP MCMP On TCD.CMPID = MCMP.CMPID"
   tmp = tmp & " Inner Join Mst_Godown MGod On MGSL.GodownID = MGod.GodownID"
   tmp = tmp & " Inner Join Mst_Commodity MComm On MGSL.CommodityID = MComm.CommodityID"
   tmp = tmp & " Where MGSL.SpillageFlag = 0 "
   tmp = tmp & " and TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' "
   tmp = tmp & " And TCG." & mGenInvoiceLocation
   
   Call TmpTable
   

   tmp = " Insert into TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " (Mst_GSL_ID,SM_Prefix,TFlag,CxTF_DetailsID,CxTFDate,CMPID,"
   tmp = tmp & " GodownID,StackNo,LotNo,CommodityID,ExchangeTypeID,"
   tmp = tmp & " ClientID,BalanceBags,BalanceWeight,Status,Flag1,GrossWt,"
   tmp = tmp & " BagSize,CDTFNo,CDTFDate,KRDate,NContractID,YearMonth,G_UID,CreatedBy,CreatedDate,"
   tmp = tmp & " UpdatedDate , UpdatedBy,KRNo,MonG_UID,MadeUpFlag,GodownYearlyBillingFlg,CommodityYearlyBillingFlg,RateCardType,BillingCycleID,BillingUnit,Rate,RateForDailyPnL,DepConDID,NoofDays)"
   tmp = tmp & " Select MGSL.Mst_GSL_ID,MGSL.SM_Prefix,'D',MGSL.CxTF_DetailsID,MGSL.CDTFDate,MGSL.CMPID,MGSL.GodownID,MGSL.StackNo,"
   tmp = tmp & " MGSL.LotNo,MGSL.CommodityID,MGSL.ExchangeTypeID,"
   tmp = tmp & " MGSL.ClientID,MGSL.DepositeBags,MGSL.DepositeWeight,MGSL.Status,MGSL.Flag1,MGSL.GrossWt,"
   tmp = tmp & " MGSL.BagSize,MGSL.CDTFNo,MGSL.CDTFDate,MGSL.KRDate,MGSL.NContractID,'" & mFromDate & "',MGSL.G_UID,MGSL.CreatedBy,MGSL.CreatedDate,"
   tmp = tmp & " MGSL.UpdatedDate , MGSL.UpdatedBy, MGSL.KRNo, newid(), MGSL.MadeUpFlag,MGod.YearlyBillingFlg 'GodownYearlyBillingFlg',MComm.YearlyBillingFlg 'CommodityYearlyBillingFlg',RateCardType,BillingCycleID,BillingUnit,Rate,RateForDailyPnL,DepConDID,NoofDays"
   tmp = tmp & " From Mst_GSL MGSL "
   tmp = tmp & " Inner Join Mst_Godown MGod On MGSL.GodownID = MGod.GodownID"
   tmp = tmp & " Inner Join Mst_Commodity MComm On MGSL.CommodityID = MComm.CommodityID"
   tmp = tmp & " Where MGSL.SpillageFlag = 0 And MGSL.CdtfNo < 0 And MGSL.Status not in ('Z') "
   tmp = tmp & " and MGSL.CDTFDate Between '" & mFromDate & "' And '" & mToDate & "'"
   tmp = tmp & " And MGSL." & mGenInvoiceLocation
   Call TmpTable

   tmp = "Select * from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set BagSize = Round(Bagsize,0) "
       Call TmpTable
       
       ProgressBar1.Value = ProgressBar1.Value + 1
       
       ' Flag off all GSL which are used for selected month in transaction table so to avoid duplicate billing
       
       tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillFlag = 'I'"
       tmp = tmp & " Where  GSLDetailID in (Select TCG.GSLDetailID"
       tmp = tmp & " From Txn_CxTF_GSL  TCG"
       tmp = tmp & " Inner Join Txn_CxTF_Details TCD On TCG.CxTF_DetailsID = TCD.CxTF_DetailsID And TCG.SM_Prefix = TCD.SM_Prefix"
       tmp = tmp & " And TCG.TxnType = TCD.TxnType"
       tmp = tmp & " where  (TCG.BillFlag is Null or TCG.BillFlag = '')  And"
       tmp = tmp & " TCD.CxTFDate Between '" & mFromDate & "' And '" & mToDate & "' And Txn_CxTF_GSL.SM_Prefix = TCG.SM_Prefix) "
       tmp = tmp & " And Txn_CxTF_GSL." & mGenInvoiceLocation
       
       Call TmpTable
          
       ProgressBar1.Value = ProgressBar1.Value + 1
       
    
        ' For Op and With
        
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownYearlyBillingFlg ="
        tmp = tmp & " (Select g.YearlyBillingFlg"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " inner join mst_Godown g on b.godownid = g.godownid"
        tmp = tmp & " where b.YearMonth = '" & mFromDate & "' And b.Mst_GSL_ID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Mst_GSL_ID"
        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,g.YearlyBillingFlg)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O') --And Txn_InvMonGSLDetail.SM_Prefix = 'GUNTR'"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CommodityYearlyBillingFlg ="
        tmp = tmp & " (Select c.YearlyBillingFlg"
        tmp = tmp & " from  TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " b"
        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
        tmp = tmp & " inner join mst_Godown g on b.godownid = g.godownid"
        tmp = tmp & " where b.YearMonth = '" & mFromDate & "' And b.Mst_GSL_ID = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".Mst_GSL_ID"
        tmp = tmp & " And b.SM_Prefix = TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix"
        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,c.YearlyBillingFlg)"
        tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".ExchangeTypeID = 1 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".TFlag in ('O') --And Txn_InvMonGSLDetail.SM_Prefix = 'GUNTR'"
        
        If CmbPreFix.Text <> "" Then
           tmp = tmp & " and TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".SM_Prefix = '" & mSM_Prefix & "'"
        End If
        
        Call TmpTable
        
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set NoOfDays = " & mTotalDays & "  Where TFlag = 'O' And BillingCycleID = 1 "
        Call TmpTable
        
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
        
        ''----- Add New on 3 March 2016 for WtNoOfDay,WtQty,CapacityNoOfDay
        
        ''--1
        tmp = "Select Distinct GodownID from TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where YEarMonth = '" & mFromDate & "' And GodownID is not null "
        Call TmpTable
        
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
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "  TIMGD "
        tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID "
        tmp = tmp & " Where TIMGD.YearMonth = '" & mStartDate & "' And DateDiff(Month,MG.StartDate,TIMGD.YearMonth) = 0 "
        tmp = tmp & " And DateDiff(Month,MG.StartDate,isnull(MG.CloseDate,'2040-01-01')) > 0 And TIMGD.CapacityNoOfDay is null "
        tmp = tmp & " And (MG.CloseDate >= '" & mStartDate & "' Or MG.CloseDate is null) And TIMGD.GodownID in " & mGodownIDs & " "
        tmp = tmp & " Group By TIMGD.GodownID,MG.StartDate "
        Call TmpTable
        
        For x = 1 To TmpRs.RecordCount
            tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
            tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID = " & TmpRs!GodownID & " "
            tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay is null"
            Call Tmp1Table
            TmpRs.MoveNext
        Next

        ''--2
        tmp = " Select TIMGD.GodownID,(Day(MG.CloseDate)) 'CapacityNoOfDay' "
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "  TIMGD "
        tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID "
        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And DateDiff(Month,TIMGD.YearMonth,MG.CloseDate) = 0 "
        tmp = tmp & " And DateDiff(Month,MG.StartDate,isnull(MG.CloseDate,'2040-01-01')) > 0 "
        tmp = tmp & " And TIMGD.CapacityNoOfDay is null And (MG.CloseDate >= '" & mFromDate & "' Or MG.CloseDate is null) And TIMGD.GodownID in " & mGodownIDs & " "
        tmp = tmp & " Group By TIMGD.GodownID,MG.CloseDate"
        Call TmpTable
    
        For x = 1 To TmpRs.RecordCount
            tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
            tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID = " & TmpRs!GodownID & " "
            tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay is null"
            Call Tmp1Table
            TmpRs.MoveNext
        Next

        ''--3
        tmp = " Select TIMGD.GodownID, " & mTotalDays & " 'CapacityNoOfDay'"
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "  TIMGD"
        tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "' And DateDiff(Month,MG.StartDate,TIMGD.YearMonth) > 0"
        tmp = tmp & " And DateDiff(Month,MG.StartDate,isnull(MG.CloseDate,'2040-01-01')) > 0"
        tmp = tmp & " And TIMGD.CapacityNoOfDay is null And (MG.CloseDate >= '" & mFromDate & "' Or MG.CloseDate is null) And TIMGD.GodownID in " & mGodownIDs & ""
        tmp = tmp & " Group By TIMGD.GodownID,MG.CloseDate"
        Call TmpTable
        
        For x = 1 To TmpRs.RecordCount
            tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
            tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID = " & TmpRs!GodownID & " "
            tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay is null"
            Call Tmp1Table
            TmpRs.MoveNext
        Next
        
        ''--4
        tmp = " Select TIMGD.GodownID,(Day(MG.CloseDate) - Day(MG.StartDate)) + 1 'CapacityNoOfDay'"
        tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & "  TIMGD"
        tmp = tmp & " Inner Join Mst_Godown MG On TIMGD.GodownID = MG.GodownID"
        tmp = tmp & " Where TIMGD.YearMonth = '" & mFromDate & "'  And DateDiff(Month,MG.StartDate,MG.CloseDate) = 0"
        tmp = tmp & " And TIMGD.CapacityNoOfDay is null And MG.CloseDate is not null And MG.CloseDate >= '" & mFromDate & "' And TIMGD.GodownID in " & mGodownIDs & ""
        tmp = tmp & " Group By TIMGD.GodownID,MG.CloseDate,MG.StartDate"
        Call TmpTable
        
        For x = 1 To TmpRs.RecordCount
            tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay = " & TmpRs!CapacityNoOfDay & " "
            tmp = tmp & " Where TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID = " & TmpRs!GodownID & " "
            tmp = tmp & " And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CapacityNoOfDay is null"
            Call Tmp1Table
            TmpRs.MoveNext
        Next
        
        '--5
        'tmp = "Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = " & mTotalDays & " Where TFlag = 'O' And Txn_InvMonGSLDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
        tmp = "Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDay = " & mTotalDays & " Where TFlag = 'O' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".YearMonth = '" & mFromDate & "' And WtNoOfDay is null And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
        Call TmpTable
        
        '--6
        'tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = (" & mTotalDays & " - Day(Txn_InvMonGSLDetail.CxTFDate)) + 1  Where YearMonth = '" & mFromDate & "' And TFlag = 'D' And WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDay = (" & mTotalDays & " - Day(TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate)) + 1  Where YearMonth = '" & mFromDate & "' And TFlag = 'D' And WtNoOfDay is null And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
        Call TmpTable
        
        '--7
        'tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = (Day(Txn_InvMonGSLDetail.CxTFDate))  Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And WtNoOfDay is null And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDay = (Day(TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate))  Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And WtNoOfDay is null And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
        Call TmpTable
        
        '--8
        'tmp = " Update Txn_InvMonGSLDetail  Set Txn_InvMonGSLDetail.WtNoOfDay = -(Day(Txn_InvMonGSLDetail.CxTFDate))  Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And DateDiff(Month,Txn_InvMonGSLDetail.CDTFDate,Txn_InvMonGSLDetail.CxTFDate) = 0 And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".WtNoOfDay = -(Day(TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate))  Where YearMonth = '" & mFromDate & "' And TFlag = 'W' And DateDiff(Month,TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CDTFDate,TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".CxTFDate) = 0 And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
        Call TmpTable
        
        '--9
        'tmp = " Update Txn_InvMonGSLDetail Set WtQty = WtNoOfDay*BalanceWeight Where YearMonth = '" & mFromDate & "' And Txn_InvMonGSLDetail.GodownID in " & mGodownIDs & ""
        tmp = " Update TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Set WtQty = WtNoOfDay*BalanceWeight Where YearMonth = '" & mFromDate & "' And TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ".GodownID in " & mGodownIDs & ""
        Call TmpTable
        
        ''-------------------------------
                
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call TmpTable
                
        tmp = "Select Distinct SM_Prefix From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & " Where " & mGenInvoiceLocation & " And YearMonth = '" & mFromDate & "'  "
        Call TmpTable
        
        For I = 1 To TmpRs.RecordCount
        
            tmp = "Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And YearMonth = '" & mFromDate & "'  "
            Call Tmp1Table
            
            If TmpRs1.RecordCount = 0 Then
                    
                tmp = " Set DeadLock_Priority Low Insert into Txn_InvMonGSLDetail "
                tmp = tmp & " (Mst_GSL_ID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
                tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientID,Status,Flag1,Tflag,GrossWt,KRDate,"
                tmp = tmp & " ContractID,BagSize,YearMonth,G_UID,CreatedBy,CreatedDate,"
                tmp = tmp & " UpdatedDate,UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,Drno,Rate,NoofDays,InvoiceAmount,CxTF_DetailsID,CxTFDate,InvoiceDate,KRNo,ProcessedBy,MonG_UID,CommTrackFlag,MadeUpFlag,GodownYearlyBillingFlg,CommodityYearlyBillingFlg,WtNoOfDay, WtQty, CapacityNoOfDay)"
                tmp = tmp & " Select Mst_GSL_ID,SM_Prefix,CMPID,CommodityID,GodownID,StackNo,LotNo,BalanceBags,"
                tmp = tmp & " BalanceWeight,ExchangeTypeID,ClientID,Status,RateCardType,TFlag,GrossWt,KRDate,"
                tmp = tmp & " NContractID,BagSize,'" & mFromDate & "',G_UID,CreatedBy,CreatedDate,"
                tmp = tmp & " UpdatedDate , UpdatedBy,CDTFDate,CDTFNo,BillingCycleID ,BillingUnit,DepConDID,Rate,NoofDays,InvoiceAmount,CxTF_DetailsID,CxTFDate,InvoiceDate,KRNo,'" & Gen_UserLoginID & "',MonG_UID,CommTrackFlag,MadeUpFlag,GodownYearlyBillingFlg,CommodityYearlyBillingFlg,WtNoOfDay , WtQty, CapacityNoOfDay "
                tmp = tmp & " From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ""
                Call Tmp1Table
               
               
                tmp = "SET IMPLICIT_TRANSACTIONS ON"
                Call Tmp1Table
               
                tmp = " Set DeadLock_Priority Low Update Txn_InvMonGSLDetail Set OldBalanceWeight = BalanceWeight Where BillingUnit = 'Q' And BalanceWeight < 0 And YearMonth= '" & mFromDate & "'"
                Call Tmp1Table
            
            
                tmp = "SET IMPLICIT_TRANSACTIONS ON"
                Call Tmp1Table
                
                tmp = " Set DeadLock_Priority Low Update Txn_InvMonGSLDetail Set BalanceWeight = (BagSize * BalanceBags)/1000 Where BillingUnit = 'Q' And BalanceWeight < 0 And isnull(ContractID,0)=0 And YearMonth= '" & mFromDate & "'"
                Call Tmp1Table
                
                tmp = "SET IMPLICIT_TRANSACTIONS ON"
                Call Tmp1Table
                
                tmp = " Set DeadLock_Priority Low Update Txn_InvMonGSLDetail Set InvoiceAmount = BalanceWeight*Rate*NoofDays Where BillingUnit = 'Q' And OldBalanceWeight < 0 And YearMonth= '" & mFromDate & "'"
                Call Tmp1Table
                
                tmp = "SET IMPLICIT_TRANSACTIONS ON"
                Call Tmp1Table
                
                tmp = " Set DeadLock_Priority Low Update Txn_InvMonGSLDetail Set InvoiceAmount = 0 Where GodownYearlyBillingFlg = '1' And CommodityYearlyBillingFlg = '1' And YearMonth= '" & mFromDate & "'"
                Call Tmp1Table
            End If
        
            TmpRs.MoveNext
        Next
        
        
        tmp = " Select SM_Prefix from Mst_Location"
        tmp = tmp & " Where SM_Prefix not in "
        tmp = tmp & " (Select Distinct SM_Prefix From TempGSL" & Gen_UserLoginID & Gen_mTimeStamp & ") "
        tmp = tmp & " And " & mGenInvoiceLocation & " "
        
        Call TmpTable
        
        For I = 1 To TmpRs.RecordCount
            
            tmp = "SET IMPLICIT_TRANSACTIONS ON"
            Call Tmp1Table
            
            tmp = " Set DeadLock_Priority Low Insert into Txn_InvMonGSLDetail "
            tmp = tmp & " (Mst_GSL_ID,SM_Prefix,CMPID,StackNo,LotNo, "
            tmp = tmp & " YearMonth,G_UID,CreatedBy,CreatedDate,ProcessedBy "
            tmp = tmp & " )"
            tmp = tmp & " Values (Null,'" & TmpRs!SM_Prefix & "',Null,'Dummy-P','Dummy-P','" & mFromDate & "',newid(),'" & Gen_UserLoginID & "','" & Now() & "','" & Gen_UserLoginID & "')"
            
            Call Tmp1Table
            TmpRs.MoveNext
        
        Next
        
    Else
        tmp = "SET IMPLICIT_TRANSACTIONS ON"
        Call TmpTable
        
        tmp = "Select SM_prefix from Mst_Location Where " & mGenInvoiceLocation & " "
        Call TmpTable
        
        For I = 1 To TmpRs.RecordCount
                
            tmp = "SET IMPLICIT_TRANSACTIONS ON"
            Call Tmp1Table
            
            tmp = " Set DeadLock_Priority Low Insert into Txn_InvMonGSLDetail "
            tmp = tmp & " (Mst_GSL_ID,SM_Prefix,CMPID,StackNo,LotNo, "
            tmp = tmp & " YearMonth,G_UID,CreatedBy,CreatedDate,ProcessedBy "
            tmp = tmp & " )"
            tmp = tmp & " Values (Null,'" & TmpRs!SM_Prefix & "',Null,'Dummy-P','Dummy-P','" & mFromDate & "',newid(),'" & Gen_UserLoginID & "','" & Now() & "','" & Gen_UserLoginID & "')"
            
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

'CmdProcess.Enabled = True

TxtConfirmation = ""
TxtRateCard.Text = ""
TxtDTD.Text = ""
TxtGSLMapping.Text = ""
TxtConfirmation.Locked = True

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

Private Sub CmdShowDTD_Click()
FrmRpt_DTD.Show
TxtDTD.Enabled = True
End Sub

Private Sub CmdShowGSL_Click()
Gen_ShowGSLMapping = True
FrmTxn_StorageContractGslMapping.Show
TxtGSLMapping.Enabled = True
End Sub

Private Sub CmdShowLocation_Click()

Call Grid_Head

ChkCheckAll.Value = 0

TxtConfirmation = ""
TxtRateCard.Text = ""
TxtDTD.Text = ""
TxtGSLMapping.Text = ""
CmdShowDTD.Enabled = False
 
TxtRateCard.Enabled = False
TxtDTD.Enabled = False
TxtGSLMapping.Enabled = False
TxtConfirmation.Locked = True
CmdShowGSL.Enabled = False
CmdShowRateCard.Enabled = False
mShowDTD = False
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


'-------Commented By Sagar on 26-08-211

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


If Gen_DBType = "MDB" Then
   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = #" & mFromDate & "# And StackNo <> 'Dummy-F' "
Else
   tmp = " Select Distinct SM_Prefix From Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "' And StackNo <> 'Dummy-F' "
End If

Call TmpTable

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

tmp = "Select Distinct ML.LocationName,ML.SM_Prefix  "
tmp = tmp & " From Mst_Location ML "
tmp = tmp & " Left Join Mst_CMP  MC On ML.LocationID = MC.LocationID "
tmp = tmp & " Left Join Mst_GSL MG On MC.CMPID = MG.CMPID "

If mChooseLocation <> "" Then
   tmp = tmp & " Where ML.SM_Prefix not in " & mChooseLocation & ""
End If

If Gen_WcLocation <> "" Then
   If mChooseLocation = "" Then
      tmp = tmp & " Where ML." & Gen_WcLocation
   Else
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

tmp = tmp & " Group by ML.LocationName,ML.SM_Prefix"
tmp = tmp & " Having Min(MC.StartDate) <= '" & mToDate & "'"

tmp = tmp & " Order By ML.LocationName "

Call TmpTable



tmp = "Select ML.LocationName ,Min(MC.CurrentDate) 'BOD'"
tmp = tmp & " From Mst_CMP MC Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
tmp = tmp & " And MC.Closedate is null"
tmp = tmp & " Group By ML.LocationName"

Call Tmp1Table

''tmp = " Select Distinct SM_Prefix"
''tmp = tmp & " From (Select GSL.SM_Prefix"
''tmp = tmp & " From  Mst_GSL GSL"
''tmp = tmp & " Left Join Txn_DepositContract TDC On  TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') "
''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On TDCD.DepConID = TDC.DepConID And GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate > '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
''tmp = tmp & " Union All"
''tmp = tmp & " Select GSL.SM_Prefix"
''tmp = tmp & " From  Mst_GSL GSL"
''tmp = tmp & " Left Join Txn_DepositContract TDC On  TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  "
''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On TDCD.DepConID = TDC.DepConID And GSL.GodownID = TDCD.GodownId"
''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And GSL.CDTFDate > '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is null And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
''tmp = tmp & " Order By SM_Prefix"

'tmp = "Select Distinct SM_Prefix  from Mst_GSL Where ExchangeTypeID = 1 And DepConDID is null and isnull(NContractID,0) = 0 And DepositeBags > 0 And CDTFDate <= '" & CDate(mToDate) & "'"

'Commented on 05 March 2018 by parag as issue occured while fetching location list
'tmp = " Select Distinct SM_Prefix From ("
'tmp = tmp & " Select Distinct l.SM_Prefix"
'tmp = tmp & " From  Mst_GSL GSL"
'tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
'tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
'tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
'tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
'tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
'tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientIDRow = Mclient.ClientIDRow"
'tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And  GSL.CDTFDate <= '" & mFromDate & "' And GSL.Flag not in ('Z') And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1') "
'tmp = tmp & " Union All"
'tmp = tmp & " Select Distinct l.SM_Prefix"
'tmp = tmp & " From  Mst_GSL GSL"
'tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
'tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
'tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
'tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
'tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
'tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientIDRow = Mclient.ClientIDRow"
'tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And GSL.Flag not in ('Z') And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1') "
'tmp = tmp & " And Convert(Varchar,GSL.GSLID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTFNo = b.CxTFNo and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where b.CxTFDate >= '" & mFromDate & "')) A"
'tmp = tmp & " Order By SM_Prefix"


'tmp = " Select Distinct l.SM_Prefix"
'tmp = tmp & " From  Mst_GSL GSL"
'tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
'tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
'tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
'tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
'tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
'tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientIDRow = Mclient.ClientIDRow"
'tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0 and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And  GSL.CDTFDate <= '" & mFromDate & "' And GSL.Flag not in ('Z') And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1') "
'tmp = tmp & " Order By l.SM_Prefix"



tmp = " Select SM_Prefix From ("
tmp = tmp & " Select Distinct gsl.SM_Prefix"
tmp = tmp & " From  Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0 and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And  GSL.CDTFDate <= '" & mToDate & "' And GSL.Status not in ('Z') And Convert(Varchar,MG.YearlyBillingFlg)+'-'+COnvert(Varchar,Mcom.YearlyBillingFlg) not in ('1-1')"
tmp = tmp & " Union All"
tmp = tmp & " Select Distinct a.SM_Prefix"
tmp = tmp & " From  Txn_CxTF_GSL  a"
tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
tmp = tmp & " Inner Join Mst_GSL GSL On a.Mst_GSL_ID = GSL.Mst_GSL_ID and a.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on b.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on b.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Where a.TxnType = 'W' And b.CxTFDate >= '" & mFromDate & "' And b.CxTFDate <= '" & mToDate & "' and GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate > '2017-01-01'   And b.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And GSL.Status in ('E') And Convert(Varchar,MG.YearlyBillingFlg)+'-'+COnvert(Varchar,Mcom.YearlyBillingFlg) not in ('1-1') "
tmp = tmp & " Union All"
tmp = tmp & " Select Distinct a.SM_Prefix"
tmp = tmp & " From  Txn_CxTF_GSL  a"
tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
tmp = tmp & " Inner Join Mst_GSL GSL On a.Mst_GSL_ID = GSL.Mst_GSL_ID and a.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on b.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on b.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Where a.TxnType = 'D' And b.CxTFDate >= '" & mFromDate & "' And b.CxTFDate <= '" & mToDate & "' and GSL.DepositeBags > 0 And  GSL.CDTFDate > '2017-01-01'   And b.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0  and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And Convert(Varchar,MG.YearlyBillingFlg)+'-'+COnvert(Varchar,Mcom.YearlyBillingFlg) not in ('1-1')) A"
tmp = tmp & " Order By SM_Prefix"


Call Tmp2Table

'tmp = " Select Distinct l.SM_Prefix"
'tmp = tmp & " From  Mst_GSL GSL"
'tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
'tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
'tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
'tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
'tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
'tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientIDRow = Mclient.ClientIDRow"
'tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = 0 And GSL.MadeUpFlag = 0 and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And GSL.Flag not in ('Z') And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1') "
'tmp = tmp & " And Convert(Varchar,GSL.GSLID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTFNo = b.CxTFNo and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where b.CxTFDate >= '" & mFromDate & "')"
'tmp = tmp & " Order By l.SM_Prefix"
'
'Call Tmp3Table


ProgressBar1.Value = 0

ProgressBar1.Max = TmpRs.RecordCount + 1
MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
With MSHFlexGrid1
For mR = 1 To TmpRs.RecordCount
    .TextMatrix(mR, 1) = TmpRs!Locationname
    .TextMatrix(mR, 2) = TmpRs!SM_Prefix
    .TextMatrix(mR, 5) = "No"
    If TmpRs1.RecordCount > 0 Then
       TmpRs1.MoveFirst
       TmpRs1.Find "LocationName = '" & .TextMatrix(mR, 1) & "'"
       If TmpRs1.EOF = False Then
          .TextMatrix(mR, 3) = Format(TmpRs1!BOD, "mmm-yyyy") 'TmpRs!SM_Prefix
          .TextMatrix(mR, 4) = Format(TmpRs1!BOD, "DD-mm-yyyy")
       End If
    End If
    If TmpRs2.RecordCount > 0 Then
       TmpRs2.MoveFirst
       TmpRs2.Find "SM_Prefix = '" & .TextMatrix(mR, 2) & "'"
       If TmpRs2.EOF = False Then
          .TextMatrix(mR, 5) = "Yes"
       End If
    End If
'    If .TextMatrix(mR, 5) <> "Yes" Then
'        If TmpRs3.RecordCount > 0 Then
'           TmpRs3.MoveFirst
'           TmpRs3.Find "SM_Prefix = '" & .TextMatrix(mR, 2) & "'"
'           If TmpRs3.EOF = False Then
'              .TextMatrix(mR, 5) = "Yes"
'           End If
'        End If
'    End If
    
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
CmdProcess.Enabled = True

End Sub

Private Sub CmdShowRateCard_Click()
FrmRpt_DepositContract.Show  'FrmRpt_RateCard.Show
TxtRateCard.Enabled = True
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

Call Grid_Head

Frame1.Enabled = True
'CmdDelete.Visible = False
ChkCheckAll.Visible = False
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
'   CmdDelete.Visible = True
   ChkCheckAll.Visible = True
End If
mCountLocation = 0
TxtConfirmation = ""
TxtConfirmation.Locked = True
CmdProcess.Enabled = False
ProgressBar1.Value = 0
mShowDTD = False
Gen_ShowGSLMapping = False

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
    CmbPreFix.AddItem RsMst_Location!Locationname
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
mLocationName = RsMst_Location!Locationname
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 6 '5 '4 '3 '26 '27 '23 '25
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(0) = 500

     .TextMatrix(0, 0) = "Select"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "SM_Prefix"
     .TextMatrix(0, 3) = "BOD(MMM-YYYY)"
     .TextMatrix(0, 4) = "BOD(DD-MM-YYYY)"
     .TextMatrix(0, 5) = "Missing Deposit Contract"
     
     .ColWidth(0) = 1000
     .ColWidth(1) = 3000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1200
     
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
   tmp = tmp & " Where MC.CurrentDate < #" & Format(mToDate, Gen_DatFormat) & "# "
Else
   tmp = tmp & " Where MC.CurrentDate < '" & Format(mToDate, Gen_DatFormat) & "' "
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

'Private Function GetMaxBOD() As Variant
'Dim Retval As Boolean
'Retval = True
'
'tmp = "Select ML.LocationName"
'tmp = tmp & " From Mst_CMP MC Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
''tmp = tmp & " Where MC.CloseDate Is Null "
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Where MC.CurrentDate <= #" & Format(mToDate, Gen_DatFormat) & "# "
'Else
'   tmp = tmp & " Where MC.CurrentDate <= '" & Format(mToDate, Gen_DatFormat) & "' "
'End If
'
'tmp = tmp & " And MC.Closedate is null "
'tmp = tmp & " And ML." & mGenInvoiceLocation
'tmp = tmp & " Order by ML.LocationName,MC.CMPName"
'
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   Retval = False
'End If
'
'GetMaxBOD = Retval
'
'End Function




Private Sub MSHFlexGrid1_Click()
'
'Head CMG

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   If LCase(Gen_UserRole) = "admin" Then 'Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Then
      If mCountLocation = 10 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.Col) = strUnChecked Then
            MsgBox "You can't select more than 5 location. Only 'Power' user can select all location!!!"
            Exit Sub
         End If
      End If
   Else
      If mCountLocation = 1 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.Col) = strUnChecked Then
            MsgBox "You can't select more than 1 location. Only 'Power' user can select all location!!!"
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
           If .TextMatrix(.row, 3) <> "" Then
              If LCase(.TextMatrix(.row, 5)) <> "yes" Then
                 If Format(.TextMatrix(.row, 3), "yyyymm") > Format(mToDate, "yyyymm") Then
                    .TextMatrix(.row, .Col) = strChecked
                    mCountLocation = mCountLocation + 1
                 Else
                    If Format(.TextMatrix(.row, 4), "yyyymmdd") = Format(mToDate, "yyyymmdd") Then
                       .TextMatrix(.row, .Col) = strChecked
                       mCountLocation = mCountLocation + 1
                    End If
                 End If
              End If
           Else
              .TextMatrix(.row, .Col) = strChecked
              mCountLocation = mCountLocation + 1
           End If
        Else
           .TextMatrix(.row, .Col) = strUnChecked
           mCountLocation = mCountLocation - 1
        End If
     End If
End With
If mCountLocation >= 1 Then
   CmdShowGSL.Enabled = True
Else
   CmdShowGSL.Enabled = False
   mShowDTD = False
End If

End Sub

Private Sub TxtConfirmation_LOstFocus()
If Trim(TxtConfirmation.Text) = "" Then Exit Sub

If LCase(Trim(TxtConfirmation.Text)) <> "all checkpoint items are completed" Then
   MsgBox "Confirmation message not correct!!!"
   TxtConfirmation.SetFocus
   CmdProcess.Enabled = False
   Exit Sub
End If

CmdProcess.Enabled = True

End Sub

Private Sub TxtDTD_LostFocus()
If TxtDTD.Text <> "" Then
   If LCase(TxtDTD.Text) = "dtd report completed" Then
      CmdShowDTD.Enabled = False
      mShowDTD = True
      TxtConfirmation.Locked = False
   Else
      MsgBox "Confirmation text for Point 3 mismatch!!!"
      TxtDTD.SetFocus
      CmdShowDTD.Enabled = True
      mShowDTD = False
      TxtConfirmation = ""
      TxtConfirmation.Locked = True
   End If
Else
   CmdShowDTD.Enabled = True
   mShowDTD = False
   TxtConfirmation = ""
   TxtConfirmation.Locked = True
End If
End Sub

Private Sub TxtGSLMapping_LostFocus()
If TxtGSLMapping.Text <> "" Then
   If LCase(TxtGSLMapping.Text) = "gsl mapping completed" Then
      CmdShowGSL.Enabled = False
      CmdShowRateCard.Enabled = True
   Else
      MsgBox "Confirmation text for Point 1 mismatch!!!"
      TxtGSLMapping.SetFocus
      CmdShowGSL.Enabled = True
      CmdShowRateCard.Enabled = False
      mShowDTD = False
   End If
Else
   CmdShowGSL.Enabled = True
   CmdShowRateCard.Enabled = False
   mShowDTD = False
End If
End Sub

Private Sub TxtRateCard_LostFocus()
If TxtRateCard.Text <> "" Then
   If LCase(TxtRateCard.Text) = "rate card report completed" Then
      CmdShowRateCard.Enabled = False
      CmdShowDTD.Enabled = True
   Else
      MsgBox "Confirmation text for Point 2 mismatch!!!"
      TxtRateCard.SetFocus
      CmdShowRateCard.Enabled = True
      CmdShowDTD.Enabled = False
      mShowDTD = False
   End If
Else
   CmdShowRateCard.Enabled = True
   CmdShowDTD.Enabled = False
   mShowDTD = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_StockOpBalance.xls")
Call Xls_Detail_Rpt
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_StockOpBalance.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
