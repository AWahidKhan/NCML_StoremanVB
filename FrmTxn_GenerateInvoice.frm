VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_GenerateInvoice 
   Caption         =   "Generate Invoice"
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
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   9150
      Width           =   15030
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   60
         TabIndex        =   27
         Top             =   165
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   840
      Left            =   120
      TabIndex        =   14
      Top             =   15
      Width           =   15015
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
         Left            =   7665
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   4140
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
         Left            =   13905
         TabIndex        =   7
         Top             =   405
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Contract Mapped"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6480
         TabIndex        =   18
         Top             =   -30
         Visible         =   0   'False
         Width           =   1770
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
         Left            =   12930
         TabIndex        =   6
         Top             =   405
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
         Left            =   11955
         TabIndex        =   5
         Top             =   405
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
         Left            =   3690
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   405
         Width           =   3900
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
         ItemData        =   "FrmTxn_GenerateInvoice.frx":0000
         Left            =   2130
         List            =   "FrmTxn_GenerateInvoice.frx":0002
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
         ItemData        =   "FrmTxn_GenerateInvoice.frx":0004
         Left            =   120
         List            =   "FrmTxn_GenerateInvoice.frx":0006
         TabIndex        =   1
         Top             =   405
         Width           =   1935
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
         Left            =   9150
         TabIndex        =   25
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
         Left            =   5220
         TabIndex        =   17
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
         Left            =   2610
         TabIndex        =   16
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
         Left            =   787
         TabIndex        =   15
         Top             =   165
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6630
      Left            =   120
      TabIndex        =   0
      Top             =   2520
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
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5445
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   9604
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "Yearly Billing flag is on"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2310
         TabIndex        =   45
         Top             =   720
         Width           =   2880
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
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
         Left            =   1590
         TabIndex        =   44
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label14 
         BackColor       =   &H0000FFFF&
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
         Left            =   10320
         TabIndex        =   43
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Rate Updated by Employee on this Screen"
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
         Left            =   10965
         TabIndex        =   42
         Top             =   210
         Width           =   2355
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
         Left            =   10560
         TabIndex        =   24
         Top             =   720
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
         Left            =   12840
         TabIndex        =   23
         Top             =   720
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
         Left            =   1590
         TabIndex        =   22
         Top             =   315
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
         Left            =   2310
         TabIndex        =   21
         Top             =   180
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
         Height          =   480
         Left            =   7755
         TabIndex        =   20
         Top             =   210
         Width           =   2280
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
         Left            =   7110
         TabIndex        =   19
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1635
      Left            =   120
      TabIndex        =   28
      Top             =   795
      Width           =   15015
      Begin VB.ComboBox CmbBillingCycle 
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
         TabIndex        =   8
         Top             =   1185
         Width           =   2670
      End
      Begin VB.ComboBox CmbBillingUnit 
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
         ItemData        =   "FrmTxn_GenerateInvoice.frx":0008
         Left            =   2835
         List            =   "FrmTxn_GenerateInvoice.frx":0012
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1185
         Width           =   2565
      End
      Begin VB.TextBox TxtNewRate 
         Alignment       =   1  'Right Justify
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
         Left            =   7050
         TabIndex        =   10
         Top             =   1185
         Width           =   1575
      End
      Begin VB.CommandButton CmdUpdateList 
         Caption         =   "Update List"
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
         Left            =   8880
         TabIndex        =   11
         Top             =   1155
         Width           =   1545
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   11010
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   465
         Width           =   3930
      End
      Begin VB.TextBox TxtLocationName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   465
         Width           =   2670
      End
      Begin VB.TextBox TxtCMPName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   2835
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   465
         Width           =   2565
      End
      Begin VB.TextBox TxtPresentRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   5430
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1575
      End
      Begin VB.TextBox TxtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   5430
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   465
         Width           =   5550
      End
      Begin VB.Label Label12 
         Caption         =   "Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   855
         TabIndex        =   41
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label11 
         Caption         =   "Billing Unit"
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
         Left            =   3615
         TabIndex        =   40
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label5 
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
         Height          =   240
         Left            =   735
         TabIndex        =   39
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label Label4 
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
         Height          =   240
         Left            =   7605
         TabIndex        =   38
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label15 
         Caption         =   "New Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7387
         TabIndex        =   37
         Top             =   885
         Width           =   900
      End
      Begin VB.Label Label17 
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
         Height          =   240
         Left            =   12435
         TabIndex        =   36
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label18 
         Caption         =   "CMP Name"
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
         Left            =   3577
         TabIndex        =   35
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label20 
         Caption         =   "Present Rate"
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
         Left            =   5617
         TabIndex        =   34
         Top             =   900
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FrmTxn_GenerateInvoice"
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

Private Sub ChkCheckAll_Click()
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   Exit Sub
End If
If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed And MSHFlexGrid1.CellBackColor <> vbGreen And MSHFlexGrid1.CellBackColor <> vbBlue Then
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 0) = strChecked
       End If
   Next
Else
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If

End Sub

Private Sub CmbBillingCycle_GotFocus()
tmp = CmbBillingCycle.Text
Call TableMst_BillingCycle
CmbBillingCycle.Clear
If RsMst_BillingCycle.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_BillingCycle.RecordCount
   CmbBillingCycle.AddItem RsMst_BillingCycle!BillingCycle
   RsMst_BillingCycle.MoveNext
Next
CmbBillingCycle.Text = tmp
End Sub

Private Sub CmbBillingCycle_LostFocus()
If CmbBillingCycle.Text = "" Then
   mBillingCycleID = Null
   Exit Sub
End If
RsMst_BillingCycle.MoveFirst
RsMst_BillingCycle.Find "BillingCycle = '" & CmbBillingCycle.Text & "'"
If RsMst_BillingCycle.EOF Then
   MsgBox "Invalid selection "
   CmbBillingCycle.SetFocus
   Exit Sub
End If
mBillingCycleID = RsMst_BillingCycle!BillingCycleID
mTxtTotDays = RsMst_BillingCycle!NoOfDays

'Call CalNoOfCylce(
End Sub


Private Sub CmbMonth_LostFocus()
Call ClearFrame
Frame4.Enabled = False
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
Call ClearFrame
Call Grid_Head
Frame4.Enabled = False
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

Call ClearFrame

ChkCheckAll.Value = 0
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

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "fagadmin" And LCase(Gen_UserRole) <> "fag" Then
   If CmbLocationID.Text = "" Then
      MsgBox "Blank Location not allowed!!!"
      CmbLocationID.SetFocus
      Exit Sub
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

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(mFromDate, "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

Call Grid_Head

Gen_mTimeStamp = GetTimeStamp



tmp = " Select * from Txn_InvMonGSLDetail Where YearMonth = '" & mFromDate & "'" 'And ExchangeTypeID = 1 "
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix & "' "
End If
Call TmpTable

If TmpRs.RecordCount = 0 Then
   If CmbLocationID.Text <> "" Then
      MsgBox "Opening balance for " & CmbLocationID.Text & " is pending!!!"
   Else
      MsgBox "Opening balance for all location is pending!!!"
   End If
   CmbLocationID.SetFocus
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

tmp = " Insert Into Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & "(Mst_GSL_ID,SM_Prefix ,ExchangeTypeID ,CMPID ,ClientID,CommodityID,BalanceBags,BalanceWeight ,Rate,InvoiceAmount ,"
tmp = tmp & " TFlag,BillingCycleID,BillingUnit,NoOfDays,InvoiceDate,PINo,Bagsize,ContractID,CxTFDate,CDTFDate,YearMonth,CxTF_DetailsID,GodownYearlyBillingFlg,CommodityYearlyBillingFlg,FranchClaimID,FranchBillingUnit,MonG_UID,GodownID)"
tmp = tmp & " Select Mst_GSL_ID,SM_Prefix ,ExchangeTypeID ,CMPID ,ClientID,CommodityID,BalanceBags,BalanceWeight ,Rate,InvoiceAmount ,"
tmp = tmp & " TFlag , BillingCycleID, BillingUnit, NoOfDays, InvoiceDate, PINo, Bagsize, ContractID, CxTFDate, CDTFDate, YearMonth,CxTF_DetailsID,GodownYearlyBillingFlg,CommodityYearlyBillingFlg ,FranchClaimID,FranchBillingUnit,MonG_UID,GodownID "
tmp = tmp & " From Txn_InvMonGslDetail"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where YearMonth = #" & mFromDate & "# And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
Else
   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And ExchangeTypeID = 1 And (ContractID = 0 Or ContractID is Null)"
End If

If CmbLocationID.Text <> "" Then
   tmp = tmp & " And SM_Prefix = '" & mSM_Prefix & "'"
End If

Call TmpTable

'tmp = "Select ML.LocationName,TINV.SM_PreFix,MCMP.CMPName,MClient.ClientName,MCG.CommodityGroup,"
'tmp = tmp & " Me.ExchangeType , Sum(TInv.BalanceBags) 'BalanceBags',Sum (TInv.BalanceWeight) 'BalanceWeight',"
'tmp = tmp & " TINV.Bagsize,TINV.YearMonth,TINV.CxTFDate ,TINV.TFlag,"
'tmp = tmp & " ML.LocationID , MCMP.CMPID, MClient.ClientIDRow, MC.CommodityID, MC.CommodityGroupID, Me.ExchangeTypeID,TINV.GSLID,TINV.CDTFDate,MC.Commodity,MST.NAVCode"
'tmp = tmp & " From Txn_InvMonGslDetail TINV"
'tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
'tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
'tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_Client MClient On TINV.ClientIDRow = MClient.ClientIDRow"
'tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
'tmp = tmp & " Left Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"

'Commented on 08 Oct 2010 1145 pm Parag

'tmp = "Select ML.LocationName,TINV.SM_PreFix,MCMP.CMPName,MClient.ClientName,MCG.CommodityGroup, Me.ExchangeType , "
'tmp = tmp & " Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight', TINV.Bagsize,TINV.YearMonth,"
'tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, ML.LocationID , MCMP.CMPID, MClient.ClientIDRow, MC.CommodityID, MC.CommodityGroupID,"
'tmp = tmp & " Me.ExchangeTypeID,TINV.GSLID,TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.NoofDays,TINV.Rate,"
'tmp = tmp & " TInv.Bagsize , Sum(TInv.InvoiceAmount) 'InvoiceAmount',TINV.BillingCycleID,TINV.BillingUnit"
'tmp = tmp & " From Txn_InvMonGslDetail TINV"
'tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
'tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
'tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_Client MClient On TINV.ClientIDRow = MClient.ClientIDRow"
'tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
'tmp = tmp & " Left Join Mst_BillingCycle MBC On TINV.BillingCycleID = MBC.BillingCycleID"
'tmp = tmp & " Left Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"


tmp = "Select ML.LocationName,TINV.SM_PreFix,MCMP.CMPName,MClient.ClientName,MCG.CommodityGroup, Me.ExchangeType , "
tmp = tmp & " Sum (TInv.BalanceBags) 'BalanceBags',Sum (TINV.BalanceWeight) 'BalanceWeight', TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, TINV.ClientID, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TINV.ExchangeTypeID,TINV.Mst_GSL_ID,TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.NoofDays,TINV.Rate,"
tmp = tmp & " TInv.Bagsize , Sum(TInv.InvoiceAmount) 'InvoiceAmount',TINV.BillingCycleID,TINV.BillingUnit,TINV.CxTF_DetailsID,TINV.GodownYearlyBillingFlg,TINV.CommodityYearlyBillingFlg,TINV.FranchClaimID,TINV.FranchBillingUnit,TINV.MonG_UID,MC.HSNNo,MC.AccountType,MGLA.HSNNo 'GLHSNNo',MG.GodownNo,MG.GodownID "
tmp = tmp & " From Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & " TINV"
tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_CommodityGroup MCG On MC.CommodityGroupID = MCG.CommodityGroupID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MClient On TINV.ClientID = MClient.ClientID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_BillingCycle MBC On TINV.BillingCycleID = MBC.BillingCycleID"
tmp = tmp & " Inner Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"
tmp = tmp & " Inner Join Mst_GLAccType MGLA On MGLA.AccountType = MC.AccountType And MGLA.Flag = 'A' "
tmp = tmp & " Inner Join Mst_GLAccountIncomeTypeMapping MGLAI On MGLA.GLAccTypeID = MGLAI.GLAccTypeID "
tmp = tmp & " Inner Join Mst_Godown MG On TINV.GodownID = MG.GodownID "
If Check1.Value = 0 Then
   tmp = tmp & " Where (TINV.ContractID = 0 Or TINV.ContractID Is Null)  " ' And TCG.BillFlag Is Null"
Else
   tmp = tmp & " Where (TINV.ContractID <> 0 And TINV.ContractID Is Not Null)  " ' And TCG.BillFlag Is Null"
End If
'tmp = tmp & " And ME.ExchangeTypeID = 1 And TINV.YearMonth='" & mFromDate & "' And (TINV.PINo is null or TINV.PINo = 0) "

tmp = tmp & " And TINV.ExchangeTypeID = 1 And TINV.YearMonth='" & mFromDate & "' And (TINV.PINo is null or TINV.PINo = 0) And MGLAI.IETypeID in (7) "
tmp = tmp & " And TINV.MonG_UID not in (Select Distinct MonG_UID from Txn_MonthwiseInvoiceData )"
If CmbLocationID.Text <> "" Then
   'tmp = tmp & " And ML.LocationID = " & mLocationID & ""
   tmp = tmp & " And TINV.SM_Prefix = '" & mSM_Prefix & "'"
   'mSM_Prefix
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

If CmbNonNcdexClient.Text <> "" Then
   tmp = tmp & " And MClient.ClientName = '" & CmbNonNcdexClient & "'"
End If

'tmp = tmp & " Group by ML.LocationName,TINV.SM_PreFix,MClient.ClientName,MCG.CommodityGroup,MCMP.CMPName,"
'tmp = tmp & " ME.ExchangeType,ME.ExchangeTypeID,TINV.YearMonth,TINV.Bagsize,TINV.TFlag,"
'tmp = tmp & " ML.LocationID,MClient.ClientIDRow,MCMP.CMPID,MC.CommodityID,MC.CommodityGroupID,TINV.CxTFDate,TINV.GSLID,TINV.CDTFDate,MC.Commodity,MST.NAVCode ,"
'tmp = tmp & " MBC.BillingCycle,TINV.NoofDays,TINV.Rate,TINV.Bagsize,TINV.BillingCycleID,TINV.BillingUnit  "

tmp = tmp & " Group by ML.LocationName , TInv.SM_PreFix, MCMP.CMPName, MClient.ClientName, MCG.CommodityGroup, Me.ExchangeType, "
tmp = tmp & " TINV.Bagsize,TINV.YearMonth,"
tmp = tmp & " TINV.CxTFDate ,TINV.TFlag, MCMP.LocationID , TINV.CMPID, TINV.ClientID, TINV.CommodityID, MC.CommodityGroupID,"
tmp = tmp & " TINV.ExchangeTypeID,TINV.Mst_GSL_ID,TINV.CDTFDate,MC.Commodity,MST.NAVCode,MBC.BillingCycle,TINV.NoofDays,TINV.Rate,"
tmp = tmp & " TInv.Bagsize ,TINV.BillingCycleID,TINV.BillingUnit,TINV.CxTF_DetailsID,TINV.GodownYearlyBillingFlg,TINV.CommodityYearlyBillingFlg,TINV.FranchClaimID,TINV.FranchBillingUnit,TINV.MonG_UID,MC.HSNNo,MC.AccountType,MGLA.HSNNo,MG.GodownNo,MG.GodownID"

tmp = tmp & " Order By ML.LocationName,MClient.ClientName,MCG.CommodityGroup"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = TmpRs.RecordCount
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
   With MSHFlexGrid1
   For mR = 2 To TmpRs.RecordCount + 1   'For mr = 1 To TmpRs.RecordCount
'       mNoofdays = 0
'       .TextMatrix(mr, 1) = TmpRs!LocationName
'       .TextMatrix(mr, 2) = TmpRs!CMPName
'       .TextMatrix(mr, 3) = TmpRs!ClientName
'       .TextMatrix(mr, 4) = TmpRs!CommodityGroup
'       .TextMatrix(mr, 5) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
'       .TextMatrix(mr, 6) = TmpRs!tFlag
'       .TextMatrix(mr, 7) = Format(IIf(IsNull(TmpRs!YearMonth), "", TmpRs!YearMonth), "dd-MMM-yyyy")
'       .TextMatrix(mr, 8) = Format(IIf(IsNull(TmpRs!CxTFDate), "", TmpRs!CxTFDate), "dd-MMM-yyyy")
'       .TextMatrix(mr, 9) = IIf(IsNull(TmpRs!CDTFDate), "", TmpRs!CDTFDate)
'       .TextMatrix(mr, 9) = Format(.TextMatrix(mr, 9), "dd-mmm-yyyy")
'
'       .TextMatrix(mr, 10) = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
'       .TextMatrix(mr, 11) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
'
'       .TextMatrix(mr, 15) = Format(IIf(IsNull(TmpRs!Bagsize), 0, TmpRs!Bagsize), "########0.00")
'
'       .TextMatrix(mr, 18) = TmpRs!GSLID
'       .TextMatrix(mr, 19) = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
'       .TextMatrix(mr, 20) = IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow)
'       .TextMatrix(mr, 21) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
'       .TextMatrix(mr, 22) = IIf(IsNull(TmpRs!ExchangeTypeID), "", TmpRs!ExchangeTypeID)
'       .TextMatrix(mr, 23) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
'       .TextMatrix(mr, 24) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
'
       
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

       .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle) '"Billing Cycle"
       .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!NoOfDays), "", TmpRs!NoOfDays) '"No of Cycles"
       .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!Rate), "", TmpRs!Rate) '"Rate"

       
       .TextMatrix(mR, 15) = Format(IIf(IsNull(TmpRs!BagSize), 0, TmpRs!BagSize), "########0.00")
       .TextMatrix(mR, 16) = Format(IIf(IsNull(TmpRs!InvoiceAmount), 0, TmpRs!InvoiceAmount), "########0.0000") '"Total Amount"
       .TextMatrix(mR, 17) = TmpRs!Mst_GSL_ID
       .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
       .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
       .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
       .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!ExchangeTypeID), "", TmpRs!ExchangeTypeID)
       .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
       .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
       .TextMatrix(mR, 24) = IIf(IsNull(TmpRs!BillingCycleID), "", TmpRs!BillingCycleID) '"BillingCycleID"
       .TextMatrix(mR, 25) = IIf(IsNull(TmpRs!BillingUnit), "", TmpRs!BillingUnit) '"BillingUnit"
       .TextMatrix(mR, 26) = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID) '"BillingUnit"
       .TextMatrix(mR, 27) = GetADhocInvoiceAmount(TmpRs!LocationID, TmpRs!ClientID) 'IIf(IsNull(TmpRs!CxTFNo), "", TmpRs!CxTFNo) '"BillingUnit"
       
       .TextMatrix(mR, 28) = IIf(IsNull(TmpRs!GodownYearlyBillingFlg), 0, TmpRs!GodownYearlyBillingFlg) '"BillingUnit"
       .TextMatrix(mR, 29) = IIf(IsNull(TmpRs!CommodityYearlyBillingFlg), 0, TmpRs!CommodityYearlyBillingFlg) '"BillingUnit"
       
       .TextMatrix(mR, 30) = IIf(IsNull(TmpRs!FranchCLaimID), 0, TmpRs!FranchCLaimID) '"BillingUnit"
       .TextMatrix(mR, 31) = IIf(IsNull(TmpRs!FranchBillingUnit), 0, TmpRs!FranchBillingUnit) '"BillingUnit"
       .TextMatrix(mR, 32) = IIf(IsNull(TmpRs!MonG_UID), 0, TmpRs!MonG_UID) '"BillingUnit"
       
       .TextMatrix(mR, 33) = IIf(IsNull(TmpRs!HSNNo), "", TmpRs!HSNNo) '"Commodity HSN No"
       If Trim(.TextMatrix(mR, 33)) = "" Then
          .TextMatrix(mR, 33) = IIf(IsNull(TmpRs!GLHSNNo), "", TmpRs!GLHSNNo) '"G/L HSN No"
       End If
       .TextMatrix(mR, 34) = IIf(IsNull(TmpRs!AccountType), "", TmpRs!AccountType) '"Agri / NAgri"
       
       .TextMatrix(mR, 35) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo) '"Agri / NAgri"
       .TextMatrix(mR, 36) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID) '"Agri / NAgri"
       
       If LCase(.TextMatrix(mR, 25)) = "b" Then
          .TextMatrix(mR, 25) = "Bags"
       ElseIf LCase(.TextMatrix(mR, 25)) = "q" Then
          .TextMatrix(mR, 25) = "Quantity"
       ElseIf LCase(.TextMatrix(mR, 25)) = "s" Then
          .TextMatrix(mR, 25) = "Sq Ft"
       End If
       
       'if Rate = 0 then mark as red
         If (Val(.TextMatrix(mR, 14)) = 0) Then
            For I = 0 To MSHFlexGrid1.Cols - 1
                .Col = I
                 .row = mR
                .CellBackColor = vbRed
            Next
         End If
         
         
         If (Val(.TextMatrix(mR, 28)) = 1) And (Val(.TextMatrix(mR, 29)) = 1) Then
            For I = 0 To MSHFlexGrid1.Cols - 1
                .Col = I
                 .row = mR
                .CellBackColor = vbBlue 'vbRed
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
         .TextMatrix(0, 16) = Val(.TextMatrix(0, 16)) + Val(.TextMatrix(mR, 16))
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
CmdSave.Enabled = True

Exit Sub

mErrorMsg:

tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & "') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
tmp = tmp & " drop table dbo.Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""

Call TmpTable

'tmp = "Drop Table Txn_TempGenInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""
'Call TmpTable

MsgBox "Some error occured while executing process!!!"


CmdGo.Enabled = False
CmdSave.Enabled = False
CmdExcel.Enabled = False

End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 3 '2
     .Cols = 37 '35 '33 '32 '30 '28 '27 '26 '27 '23 '25
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     .RowHeight(1) = 700 '.RowHeight(0) = 700

'     .TextMatrix(0, 0) = "Select"
'     .TextMatrix(0, 1) = "Location Name"
'     .TextMatrix(0, 2) = "CMP Name"
'     .TextMatrix(0, 3) = "Client Name"
'     .TextMatrix(0, 4) = "Commodity"
'     .TextMatrix(0, 5) = "Exchange Type"
'     .TextMatrix(0, 6) = "Txn Type"
'     .TextMatrix(0, 7) = "Billing Month" '"Year Month"
'     .TextMatrix(0, 8) = "Txn Date" '"CxTF Date"
'     .TextMatrix(0, 9) = "CDTF Date" '.TextMatrix(0, 19) = "CDTF Date"
'     .TextMatrix(0, 10) = "No Of Bags" '"Balance Bags"
'     .TextMatrix(0, 11) = "Weight in MT" '"Balance Weight"
'     .TextMatrix(0, 12) = "Billing Cycle"
'     .TextMatrix(0, 13) = "No of Cycles"
'     .TextMatrix(0, 14) = "Rate"
'     .TextMatrix(0, 15) = "Bagsize"
'     .TextMatrix(0, 16) = "Total Amount"
'     .TextMatrix(0, 17) = "GSLID"
'     .TextMatrix(0, 18) = "ST NAV Code"
'     .TextMatrix(0, 19) = "ClientIDRow"
'     .TextMatrix(0, 20) = "CommodityID"
'     .TextMatrix(0, 21) = "ExchangeTypeID"
'     .TextMatrix(0, 22) = "LocationID"
'     .TextMatrix(0, 23) = "SM_Prefix"
'     .TextMatrix(0, 24) = "BillingCycleID"
'     .TextMatrix(0, 25) = "Billing Unit"
'     .TextMatrix(0, 26) = "CxTFNo"
'     .TextMatrix(0, 27) = "Adhoc Invoice Amount (Only for information)"

     
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
     .TextMatrix(1, 12) = "Billing Cycle"
     .TextMatrix(1, 13) = "No of Cycles"
     .TextMatrix(1, 14) = "Rate"
     .TextMatrix(1, 15) = "Bagsize"
     .TextMatrix(1, 16) = "Total Amount"
     .TextMatrix(1, 17) = "GSLID"
     .TextMatrix(1, 18) = "GST Group Code" '"ST NAV Code"
     .TextMatrix(1, 19) = "ClientID"
     .TextMatrix(1, 20) = "CommodityID"
     .TextMatrix(1, 21) = "ExchangeTypeID"
     .TextMatrix(1, 22) = "LocationID"
     .TextMatrix(1, 23) = "SM_Prefix"
     .TextMatrix(1, 24) = "BillingCycleID"
     .TextMatrix(1, 25) = "Billing Unit"
     .TextMatrix(1, 26) = "CxTFNo"
     .TextMatrix(1, 27) = "Adhoc Invoice Amount (Only for information)"
     .TextMatrix(1, 28) = "Godown Yearly Flag"
     .TextMatrix(1, 29) = "Commodity Yearly Flag"
     .TextMatrix(1, 30) = "Franch Claim ID"
     .TextMatrix(1, 31) = "Franch Billing Unit"
     .TextMatrix(1, 32) = "MonG_UID"
     .TextMatrix(1, 33) = "HSN No"
     .TextMatrix(1, 34) = "Account Type"
     .TextMatrix(1, 35) = "Godown No"
     .TextMatrix(1, 36) = "Godown ID"
     
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
     .ColWidth(20) = 0
     .ColWidth(21) = 0
     .ColWidth(22) = 0
     .ColWidth(23) = 0
     .ColWidth(24) = 0
     .ColWidth(25) = 1000
     .ColWidth(27) = 1000
     
End With
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
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
   Call ClearFrame
   Call Grid_Head
   Frame4.Enabled = False
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
Private Sub CmdSave_Click()

'On Error GoTo msgError

MsgBox "Wait till next message!!!"

If MSHFlexGrid1.Rows <= 2 Then
   MsgBox "No data in grid!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

 
mBillingMonth = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")


'tmp = " if exists (select * from dbo.sysobjects where id = object_id(N'dbo.Txn_TempProInvoice') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
'tmp = tmp & " drop table dbo.Txn_TempProInvoice"
'Call TmpTable


Call TableTxn_CreateTempProInvoice
Call TableTxn_TempProInvoice

'If ChkRate = False Then
'   MsgBox "One or more GSL Storage Rate not availabe in 'Storage Rate Master' or " & _
'          "'Special Rate Master'. Please update rate masters. "
'   Exit Sub
'End If

CmdSave.Enabled = False
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows
For I = 2 To MSHFlexGrid1.Rows - 1 'For I = 1 To MSHFlexGrid1.Rows - 1
       
    If MSHFlexGrid1.TextMatrix(I, 1) = "" Then Exit For
    MSHFlexGrid1.Col = 0
    MSHFlexGrid1.row = I
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       RsTxn_TempProInvoice.AddNew
       RsTxn_TempProInvoice!LocationID = MSHFlexGrid1.TextMatrix(I, 22)
       RsTxn_TempProInvoice!SM_Prefix = MSHFlexGrid1.TextMatrix(I, 23)
       RsTxn_TempProInvoice!Mst_GSL_ID = MSHFlexGrid1.TextMatrix(I, 17)
       RsTxn_TempProInvoice!ClientID = MSHFlexGrid1.TextMatrix(I, 19)
       RsTxn_TempProInvoice!CommodityID = MSHFlexGrid1.TextMatrix(I, 20)
       RsTxn_TempProInvoice!ExchangeTypeID = MSHFlexGrid1.TextMatrix(I, 21)
       RsTxn_TempProInvoice!ServiceTaxNAVCode = MSHFlexGrid1.TextMatrix(I, 18)
       RsTxn_TempProInvoice!tFlag = MSHFlexGrid1.TextMatrix(I, 6)
       RsTxn_TempProInvoice!NoofBags = Val(MSHFlexGrid1.TextMatrix(I, 10))
       RsTxn_TempProInvoice!Weight = Val(MSHFlexGrid1.TextMatrix(I, 11))
       RsTxn_TempProInvoice!Rate = Val(MSHFlexGrid1.TextMatrix(I, 14))
       RsTxn_TempProInvoice!BillingCycleID = Val(MSHFlexGrid1.TextMatrix(I, 24))
       RsTxn_TempProInvoice!BillingUnit = Left(MSHFlexGrid1.TextMatrix(I, 25), 1)
       RsTxn_TempProInvoice!NoOfDays = Val(MSHFlexGrid1.TextMatrix(I, 13))
       RsTxn_TempProInvoice!InvoiceDate = CDate(MSHFlexGrid1.TextMatrix(I, 7))
       RsTxn_TempProInvoice!Rate = Val(MSHFlexGrid1.TextMatrix(I, 14))
       RsTxn_TempProInvoice!Amount = Val(MSHFlexGrid1.TextMatrix(I, 16))
       If LCase(MSHFlexGrid1.TextMatrix(I, 6)) <> "o" Then
          RsTxn_TempProInvoice!CxTFDate = CDate(MSHFlexGrid1.TextMatrix(I, 8))
       End If
       RsTxn_TempProInvoice!MonG_UID = Trim(MSHFlexGrid1.TextMatrix(I, 32))
       RsTxn_TempProInvoice!HSNNo = Trim(MSHFlexGrid1.TextMatrix(I, 33))
       RsTxn_TempProInvoice!AccountType = Trim(MSHFlexGrid1.TextMatrix(I, 34))
       RsTxn_TempProInvoice.Update
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
   
ProgressBar1.Value = ProgressBar1.Max
   
Call TableTxn_ProInvoice
   

'tmp = "Select LocationID,SM_Prefix,ExchangeTypeID,ServiceTaxNAVCode,ClientIDRow, "
'tmp = tmp & " Sum(NoOfBags) as 'NoOfBags',Sum(Weight) as 'Weight',Rate,Sum(Amount) as 'Amount' "
'tmp = tmp & " From Txn_TempProInvoice "
'tmp = tmp & " Group By LocationID,SM_Prefix,ExchangeTypeID,ServiceTaxNAVCode,ClientIDRow,Rate "


tmp = "Select LocationID,SM_Prefix,ExchangeTypeID,ServiceTaxNAVCode,ClientID,HSNNo,AccountType, "
tmp = tmp & " Sum(Amount) as 'Amount' "
tmp = tmp & " From Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & " "
tmp = tmp & " Group By LocationID,SM_Prefix,ExchangeTypeID,ServiceTaxNAVCode,ClientID,HSNNo,AccountType "

Call Tmp2Table
   
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs2.RecordCount + 1

For I = 1 To TmpRs2.RecordCount
       
    RsTxn_ProInvoice.AddNew
'    RsTxn_ProInvoice!InvoiceNo = GetMaxInvoiceNo(TmpRs2!SM_Prefix)
'    mInvoiceNo = RsTxn_ProInvoice!InvoiceNo
    RsTxn_ProInvoice!G_UID = GetGUID
    mInvG_UID = RsTxn_ProInvoice!G_UID
    RsTxn_ProInvoice!BillingMonth = CDate(mBillingMonth)
    RsTxn_ProInvoice!InvoiceDate = CDate(mToDate)
    RsTxn_ProInvoice!LocationID = TmpRs2!LocationID
    RsTxn_ProInvoice!SM_Prefix = TmpRs2!SM_Prefix
    RsTxn_ProInvoice!ClientID = TmpRs2!ClientID
    RsTxn_ProInvoice!ExchangeTypeID = TmpRs2!ExchangeTypeID
    RsTxn_ProInvoice!ServiceTaxNAVCode = TmpRs2!ServiceTaxNAVCode
    RsTxn_ProInvoice!NoofBags = 0 'Val(TmpRs2!NoofBags)
    RsTxn_ProInvoice!Weight = 0 'Val(TmpRs2!Weight)
    RsTxn_ProInvoice!Rate = 0 'Val(TmpRs2!Rate)
    RsTxn_ProInvoice!Amount = Val(TmpRs2!Amount)
    RsTxn_ProInvoice!Flag = "P"
    RsTxn_ProInvoice!HSNNo = TmpRs2!HSNNo
    RsTxn_ProInvoice!AccountType = TmpRs2!AccountType
    If IsNull(RsTxn_ProInvoice!CreatedBy) = True Or RsTxn_ProInvoice!CreatedBy = "" Then
       RsTxn_ProInvoice!CreatedBy = Gen_UserLoginID
       RsTxn_ProInvoice!CreatedDate = Now
    Else
       RsTxn_ProInvoice!UpdatedBy = Gen_UserLoginID
       RsTxn_ProInvoice!UpdatedDate = Now
    End If
    RsTxn_ProInvoice!Concurrency = Now
    
'    tmp = "Select * from Txn_ProInvoice Where InvoiceNo = " & mInvoiceNo & " "
'    Call Tmp1Table
    
'    If TmpRs1.RecordCount > 0 Then
'       RsTxn_ProInvoice!InvoiceNo = GetMaxInvoiceNo(TmpRs2!SM_Prefix)
'       mInvoiceNo = RsTxn_ProInvoice!InvoiceNo
'    End If
    
    RsTxn_ProInvoice.Update
    
    
    
       
    tmp = "Select InvoiceNo From Txn_ProInvoice Where G_UID='" & mInvG_UID & "'"
    Call TmpTable

    If TmpRs.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
       mInvoiceNo = TmpRs!InvoiceNo
    End If



    tmp = "Select a.InvoiceDate,a.LocationID,a.ClientID,st.NAVCode,Sum(a.Amount) 'Amount'"
    tmp = tmp & " from Txn_TempInvoice a"
    tmp = tmp & " Inner Join Mst_ServiceTax st on a.ServiceTaxID = st.ServiceTaxID"
    tmp = tmp & " Where a.InvoiceDate = '" & CDate(mBillingMonth) & "' And Isnull(a.ReverseFlag,'N') = 'N' And a.Flag = 'A' And Isnull(a.BatchID,0) <> 0 And a.IETypeID = 7 And a.LocationID = " & TmpRs2!LocationID & " And a.ClientID = " & TmpRs2!ClientID & " And st.NAVCode = '" & TmpRs2!ServiceTaxNAVCode & "'"
    tmp = tmp & " Group  By a.InvoiceDate,a.LOcationID,a.ClientID,st.NAVCode"
    
    Call TmpTable
    
    For xy = 1 To TmpRs.RecordCount
        mAdhocAmt = TmpRs!Amount
        tmp = "Select * from Txn_ProInvoice Where InvoiceNo = " & mInvoiceNo & " And BillingMonth = '" & CDate(mBillingMonth) & "' And LocationID = " & TmpRs!LocationID & " And ClientID = " & TmpRs!ClientID & " And ServiceTaxNAVCode = '" & TmpRs!NavCode & "'"
        Call Tmp1Table
         
        For pqr = 1 To TmpRs1.RecordCount
            If TmpRs1!Amount >= Val(mAdhocAmt) Then
               TmpRs1!Discount = Val(TmpRs1!Discount) + Val(mAdhocAmt)
               mAdhocAmt = 0
            Else
               TmpRs1!Discount = Val(TmpRs1!Discount) + Val(TmpRs1!Amount)
               mAdhocAmt = Val(mAdhocAmt) - Val(TmpRs1!Amount)
            End If
            TmpRs1.Update
            If mAdhocAmt = 0 Then Exit For
            TmpRs1.MoveNext
        Next pqr
        TmpRs.MoveNext
    Next

    TmpRs2.MoveNext
      
       
    tmp = "Update Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & " Set PINo = " & mInvoiceNo & " " & _
          " Where SM_Prefix = '" & RsTxn_ProInvoice!SM_Prefix & "' And " & _
          " ClientID = " & RsTxn_ProInvoice!ClientID & " And " & _
          " ExchangeTypeID = " & RsTxn_ProInvoice!ExchangeTypeID & " And " & _
          " ServiceTaxNAVCode = '" & RsTxn_ProInvoice!ServiceTaxNAVCode & "' And " & _
          " HSNNo = '" & RsTxn_ProInvoice!HSNNo & "' And " & _
          " AccountType = '" & RsTxn_ProInvoice!AccountType & "' "
    Call Tmp1Table
    
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
   
tmp = "Insert Into Txn_MonthwiseInvoiceData (YearMonth,Mst_GSL_ID,SM_Prefix,MonG_UID,PINo,CreatedBy,Concurrency) Select '" & mBillingMonth & "',Mst_GSL_ID,SM_Prefix,MonG_UID,PINo,'" & Gen_UserLoginID & "',getdate() From Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""
Call Tmp3Table

ProgressBar1.Value = ProgressBar1.Max

''Commented on 05 Oct 2016
'''''''tmp = "Select * from Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & " Where PINo is Not Null"
'''''''Call Tmp3Table
'''''''
'''''''ProgressBar1.Value = 0
'''''''ProgressBar1.Max = TmpRs3.RecordCount + 1
'''''''
'''''''For I = 1 To TmpRs3.RecordCount
'''''''    If LCase(TmpRs3!tFlag) <> "o" Then
'''''''          tmp = "SET IMPLICIT_TRANSACTIONS ON"
'''''''          Call Tmp2Table
'''''''
'''''''
'''''''       tmp = "Select * from Txn_InvMonGSLDetail Where GSLID = " & TmpRs3!gslid & " And SM_Prefix = '" & TmpRs3!SM_Prefix & "' And YearMonth = '" & mBillingMonth & "' And TFlag = '" & TmpRs3!tFlag & "' And CxTFDate = '" & TmpRs3!CxTFDate & "' And PINo is Not Null "
'''''''       Call Tmp2Table
'''''''
'''''''       If TmpRs2.RecordCount = 0 Then
'''''''
'''''''          tmp = "SET IMPLICIT_TRANSACTIONS ON"
'''''''          Call Tmp1Table
'''''''
'''''''          tmp = "  Set DeadLock_Priority Low  Update Txn_InvMonGSLDetail Set PINo = " & TmpRs3!PINo & ", " & _
'''''''                "InvoiceDate = '" & mBillingMonth & "'  Where GSLID = " & TmpRs3!gslid & " And SM_Prefix = '" & TmpRs3!SM_Prefix & "' And YearMonth = '" & mBillingMonth & "' And TFlag = '" & TmpRs3!tFlag & "' And CxTFDate = '" & TmpRs3!CxTFDate & "'"
'''''''          Call Tmp1Table
'''''''       End If
'''''''    Else
'''''''       tmp = "Select * from Txn_InvMonGSLDetail Where GSLID = " & TmpRs3!gslid & " And SM_Prefix = '" & TmpRs3!SM_Prefix & "' And YearMonth = '" & mBillingMonth & "' And TFlag = '" & TmpRs3!tFlag & "' And PINo is Not Null "
'''''''       Call Tmp2Table
'''''''
'''''''       If TmpRs2.RecordCount = 0 Then
'''''''
'''''''          tmp = "SET IMPLICIT_TRANSACTIONS ON"
'''''''          Call TmpTable
'''''''
'''''''          tmp = "  Set DeadLock_Priority Low  Update Txn_InvMonGSLDetail Set PINo = " & TmpRs3!PINo & ", " & _
'''''''                "InvoiceDate = '" & mBillingMonth & "'  Where GSLID = " & TmpRs3!gslid & " And SM_Prefix = '" & TmpRs3!SM_Prefix & "' And YearMonth = '" & mBillingMonth & "' And TFlag = '" & TmpRs3!tFlag & "'"
'''''''
'''''''          Call Tmp1Table
'''''''       End If
'''''''    End If
'''''''    TmpRs3.MoveNext
'''''''    ProgressBar1.Value = ProgressBar1.Value + 1
'''''''Next
'''''''ProgressBar1.Value = ProgressBar1.Max

tmp = "Drop Table Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""
Call TmpTable
'
'5   GST with Tax    GSTS18
'6   GST without Tax GSTS00



MsgBox "Done!!!"

Call Grid_Head

Exit Sub

msgError:

   tmp = "Drop Table Txn_TempProInvoice" & Gen_UserLoginID & Gen_mTimeStamp & ""
   Call TmpTable

   'MsgBox "Some error occured while process!!!"
   MsgBox Err.Description, vbExclamation, "Error"
   
   Call Grid_Head

End Sub
Private Function GetMaxInvoiceNo(SM_Prefix_ As Variant) As Double
Dim Retval As Double
tmp = "Select max(InvoiceNo) from Txn_ProInvoice " 'Where SM_Prefix = '" & SM_Prefix_ & "' "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxInvoiceNo = Retval
End Function
Private Function ChkRate() As Boolean
Dim Retval As Boolean
Retval = True
For I = 1 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.Col = 0
    MSHFlexGrid1.row = I
    If MSHFlexGrid1.CellBackColor = vbRed Then
       Retval = False
       Exit For
    End If
Next

ChkRate = Retval

End Function

Private Sub CmdUpdateList_Click()
Dim x As Variant

If CmbBillingCycle.Text = "" Then
   MsgBox "Blank Billing Cycle not allowed!!!"
   CmbBillingCycle.SetFocus
   Exit Sub
End If

If CmbBillingUnit.Text = "" Then
   MsgBox "Blank Billing Unit not allowed!!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If TxtNewRate.Text = "" Then
   MsgBox "Please Enter New Rate !!"
   TxtNewRate.SetFocus
   Exit Sub
End If

x = MSHFlexGrid1.RowSel

MSHFlexGrid1.TextMatrix(x, 14) = Val(TxtNewRate.Text)
MSHFlexGrid1.TextMatrix(x, 12) = CmbBillingCycle.Text
MSHFlexGrid1.TextMatrix(x, 25) = CmbBillingUnit.Text

'Update Details in InvMonGSL table
Call UpdateTxn_InvMonGSL(x)

MSHFlexGrid1.row = x
For C = 0 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Col = C
    MSHFlexGrid1.CellBackColor = vbYellow
Next
Call ClearFrame
CmdGo.SetFocus
Frame4.Enabled = False
End Sub
Private Sub UpdateTxn_InvMonGSL(mROw_ As Variant)
Dim mChoice As Variant
Dim mNoofDays, mNoOfCycle As Variant

mChoice = MSHFlexGrid1.TextMatrix(mROw_, 6) & Left(CmbBillingCycle.Text, 1)
mSM_Prefix = MSHFlexGrid1.TextMatrix(mROw_, 23)
''Find No Of Days
mNoofDays = GetBillingCycleDays(mBillingCycleID)

'' Year Month
'mFromDate = "01-" & CmbMonth.Text & "-" & CmbYear.Text
'mFromDate = Format(mFromDate, "yyyy-mm-dd")

'Transaction Date To Calculate No OF Cycle
mDate_ = Format(MSHFlexGrid1.TextMatrix(mROw_, 8), "yyyy-mm-dd")

tmp = "SET IMPLICIT_TRANSACTIONS ON"
Call TmpTable

tmp1 = "  Set DeadLock_Priority Low Update Txn_InvMonGslDetail Set OldRate=Rate,OldBillingCycleID=BillingCycleID,OldBillingUnit=BillingUnit,"
tmp1 = tmp1 & " Rate=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 14)) & ",BillingCycleID=" & mBillingCycleID & " ,BillingUnit='" & mFlag & "',UpdatedDate = Getdate(), UpdatedBy = " & Gen_UserLoginID & ","

Select Case (mChoice)

Case "OM"
   mNoOfCycle = 1
   tmp = "  NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
   tmp = tmp & " And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (BillingCycleID = 4 Or BillingCycleID is null)  And
Case "OF"
   mNoOfCycle = 2
   tmp = " NoofDays = 2 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
   tmp = tmp & " And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (BillingCycleID = 3 Or BillingCycleID is null) And
Case "OW"
   mNoOfCycle = 4
   tmp = " NoofDays = 4 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
   tmp = tmp & " And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (BillingCycleID = 2 Or BillingCycleID is null) And
Case "OD"
   mNoOfCycle = mTotalDays
   tmp = " NoofDays = " & mTotalDays & " Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'O'"
   tmp = tmp & " And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (BillingCycleID = 1 Or BillingCycleID is null) And
Case "DM" 'DM & WM
   mNoOfCycle = 1
   tmp = " NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D'"
   tmp = tmp & " And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (BillingCycleID = 4 Or BillingCycleID is null) And
Case "WM"
   mNoOfCycle = 1
   tmp = " NoofDays = 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W'"
   tmp = tmp & " And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (BillingCycleID = 4 Or BillingCycleID is null) And
Case "DD" 'DD & WD

'Commented on 18 June 2011 By Parag
'   mNoOfCycle = 30 - Format(mDate_, "dd")
'   tmp = " NoOfDays = 30-Day('" & mDate_ & "') + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D'  And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 And (BillingCycleID = 1 Or BillingCycleID is null)

    mNoOfCycle = mTotalDays - Format(mDate_, "dd") + 1 'Day(" & mDate_ & ") + 1
    tmp = " NoOfDays = " & mTotalDays & " - Day('" & mDate_ & "') + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'D'  And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 And (BillingCycleID = 1 Or BillingCycleID is null)


Case "WD"
   mNoOfCycle = Format(mDate_, "dd")
   'tmp = " NoOfDays = Day('" & mDate_ & "') + 1 Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W'  And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 And (BillingCycleID = 1 Or BillingCycleID is null)
   
   tmp = " NoOfDays = Day('" & mDate_ & "') Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W'  And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 And (BillingCycleID = 1 Or BillingCycleID is null)
   
Case "DW"
   tmp = "Select Case Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1)) End As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   tmp = " NoOfDays =  Case Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1))  When 0 then 1 else Floor((4 - (Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ") + 1)) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (Txn_InvMonGslDetail.BillingCycleID = 2 Or Txn_InvMonGslDetail.BillingCycleID is null) And
Case "DF"
   tmp = "Select Case Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) End As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   tmp = " NoOfDays = Case Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) When 0 then 1 else Floor((4 - Round(Convert(Decimal,Day('" & mDate_ & "'))/" & mNoofDays & ",1)) - 1) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'D' And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (Txn_InvMonGslDetail.BillingCycleID = 3 Or Txn_InvMonGslDetail.BillingCycleID is null) And
Case "WW"
   tmp = "Select Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End  As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   tmp = " NoOfDays = Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 5 Then 4 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 "
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (Txn_InvMonGslDetail.BillingCycleID = 2 Or Txn_InvMonGslDetail.BillingCycleID is null) And
Case "WF"
   tmp = "Select Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End  As NoOfCycle"
   Call TmpTable
   mNoOfCycle = TmpRs!NoofCycle
   
   tmp = " NoOfDays = Case CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) When 3 Then 2 Else CEILING(((Round(Convert(Numeric,Day('" & mDate_ & "'))/" & mNoofDays & ",1)))) End "
   tmp = tmp & " Where Txn_InvMonGslDetail.YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1"
   tmp = tmp & " And Txn_InvMonGslDetail.TFlag = 'W' And  (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' " ' Commented on 25 Jan 2011 (Txn_InvMonGslDetail.BillingCycleID = 3 Or Txn_InvMonGslDetail.BillingCycleID is null) And
Case ""
   Exit Sub
End Select


If tmp <> "" Then
   tmp = tmp & " And Mst_GSL_ID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 17)) & ""
   If LCase(mTFlagOnUpdate) = "w" Or LCase(mTFlagOnUpdate) = "d" Then
      tmp = tmp & " And CxTF_DetailsID =" & Val(MSHFlexGrid1.TextMatrix(mROw_, 26)) & ""
   End If
   tmp = tmp1 & tmp
   Call TmpTable
End If

'Op Bal and Monthly Cycle** OM
'Op Bal and Forthnightly Cycle ** OF
'Op Bal and Weekly Cycle ** OW
'Op Bal and Daily Cycle ** OD
'Depo and Withd and Monthly Cycle ** DWM
'Depo and Withd and Daily Cycle ** DWD
'Dont Know ==============
'''tmp = "Update Txn_InvMonGslDetail Set NoOfDays = Day('" & mDate_ & "') Where YearMonth = '" & mFromDate & "'  And ExchangeTypeID = 1 And TFlag = 'W' And BillingCycleID = 1 And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' "
'''Call TmpTable
'Deposit & Weekly  ** DW
'Deposit & FortNightly  ** DF
'Withdrawal & weekly ** WW
'Withdrawal & Fortnightly ** WF
'Updating Invoice Amount And Invoice Date

tmp = "SET IMPLICIT_TRANSACTIONS ON"
Call TmpTable

If LCase(MSHFlexGrid1.TextMatrix(mROw_, 25)) = "bags" Then
   mNoofBagsNQuantity = MSHFlexGrid1.TextMatrix(mROw_, 10)
   tmp = "  Set DeadLock_Priority Low  Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceBags * NoOfDays * Rate "
   tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
   tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'B' And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' "
ElseIf LCase(MSHFlexGrid1.TextMatrix(mROw_, 25)) = "quantity" Then
   mNoofBagsNQuantity = MSHFlexGrid1.TextMatrix(mROw_, 11)
   tmp = " Set DeadLock_Priority Low  Update Txn_InvMonGslDetail Set InvoiceAmount = BalanceWeight * NoOfDays * Rate "
   tmp = tmp & " Where YearMonth = '" & mFromDate & "'"
   tmp = tmp & " And Txn_InvMonGslDetail.ExchangeTypeID = 1 And BillingUnit = 'Q' And (ContractID = 0 Or ContractID is Null) And SM_Prefix = '" & mSM_Prefix & "' "
End If
tmp = tmp & " And Mst_GSL_ID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 17)) & ""
Call TmpTable

If LCase(MSHFlexGrid1.TextMatrix(mROw_, 31)) = "p" Then
   tmp = " Update Txn_InvMonGSLDetail Set FranchInvoiceAmount =  (InvoiceAmount)* FranchiseeRate /100 ,InvoiceDate =  '" & mToDate & "' "
   tmp = tmp & " Where YearMonth = '" & mFromDate & "' And InvoiceAmount > 0 "
   tmp = tmp & " And FranchBillingUnit = 'P' And SM_Prefix = '" & mSM_Prefix & "' and Mst_GSL_ID=" & Val(MSHFlexGrid1.TextMatrix(mROw_, 17)) & " "

   Call TmpTable
End If

'Invoice Amount
MSHFlexGrid1.TextMatrix(mROw_, 16) = mNoofBagsNQuantity * Val(MSHFlexGrid1.TextMatrix(mROw_, 14)) * mNoOfCycle


'Condition for below qry updated by Parag Sir
'tmp = " Update Txn_InvMonGslDetail Set InvoiceAmount = 0,InvoiceDate = Null Where YearMonth = '" & mFromDate & "' And Txn_InvMonGslDetail.ExchangeTypeID = 1 And TFlag = 'W' And (ContractID = 0 Or ContractID is Null) And Month('" & mDate_ & "') = Month(CDTFDate) And SM_Prefix = '" & mSM_Prefix & "' "
'Call TmpTable



MSHFlexGrid1.TextMatrix(mROw_, 13) = mNoOfCycle

End Sub

Private Sub ClearFrame()
TxtLocationName.Text = ""
TxtCMPName.Text = ""
TxtClientName.Text = ""
TxtCommodity.Text = ""
CmbBillingCycle.Text = ""
CmbBillingUnit.Text = ""
TxtPresentRate.Text = ""
TxtNewRate.Text = ""
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
Frame4.Enabled = False

Call Grid_Head

For I = 1 To 12
    CmbMonth.AddItem MonthName(I), I - 1
Next
CmdGo.Enabled = True
CmdExcel.Enabled = True
CmdSave.Enabled = False

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "account" Then
   CmdSave.Visible = False
Else
   CmdSave.Visible = True
End If

ProgressBar1.Value = 0

End Sub

Public Function RoundUp(dblValue As Double) As Double
On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
   RoundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
Else
   RoundUp = dblValue
End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox Err.Description, vbInformation, "Round Up"
End Function


Private Sub MSHFlexGrid1_Click()

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "account" Then
   Exit Sub
End If

With MSHFlexGrid1
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
           .Col = 0
        If (.CellBackColor) <> vbRed And (.CellBackColor) <> vbGreen And (.CellBackColor) <> vbBlue Then
           If .TextMatrix(.row, .Col) = strUnChecked Then
              .TextMatrix(.row, .Col) = strChecked
           Else
              .TextMatrix(.row, .Col) = strUnChecked
           End If
        Else
           If .CellBackColor = vbRed Then
              MsgBox "Storage Rate not availabe in Storage Rate Master / Yearly Billing flag is on !!!" 'Storage Rate Not Availabe in Storage Rate Master !!!"
           ElseIf .CellBackColor = vbGreen Then
              MsgBox "Deposit and Withdrawal of selected lot done in same month so can't generate invoice for this withdrawal!!!"
           ElseIf .CellBackColor = vbBlue Then
              MsgBox "Yearly Billing Flag is On!!!"
           ElseIf .CellBackColor = vbYellow Then
              MsgBox "Service Tax not mapped with Commodity Master"
           End If
        End If
''     Else
''        'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
''        '   Exit Sub
''        'End If
''        If .TextMatrix(.RowSel, 1) = "" Then Exit Sub
''        If .CellBackColor = vbGreen Then Exit Sub
''        If Val(.TextMatrix(.RowSel, 30)) <> 0 And LCase(.TextMatrix(.RowSel, 31)) = "p" Then
''           MsgBox "Franchisee process executed and Bill Generated by CMG ,GSL Invoice data can't be edited!!!!"
''           Exit Sub
''        End If
''        Frame4.Enabled = True
''        TxtLocationName.Text = .TextMatrix(.RowSel, 1)
''        TxtCMPName.Text = .TextMatrix(.RowSel, 2)
''        TxtClientName.Text = .TextMatrix(.RowSel, 3)
''        TxtCommodity.Text = .TextMatrix(.RowSel, 4)
''        CmbBillingCycle.Text = .TextMatrix(.RowSel, 12)
''        CmbBillingUnit.Text = .TextMatrix(.RowSel, 25)
''        TxtPresentRate.Text = .TextMatrix(.RowSel, 14)
''        mTFlagOnUpdate = .TextMatrix(.RowSel, 6)
''        mTxnNo = .TextMatrix(.RowSel, 26)
''        Call CmbBillingCycle_GotFocus
''        Call CmbBillingCycle_LostFocus
''        Call CmbBillingUnit_LostFocus
''        If CmbBillingCycle.Text = "" Then
''           CmbBillingCycle.SetFocus
''        Else
''           TxtNewRate.SetFocus
''        End If

     End If
End With

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_InvoiceDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
        If C <> 35 Then
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
        Else
           oWS.Cells(mRow, C) = "'" & MSHFlexGrid1.TextMatrix(I, C)
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
Private Sub CmbBillingUnit_LostFocus()
If CmbBillingUnit.Text = "" Then
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) <> "bags" And LCase(CmbBillingUnit.Text) <> "quantity" Then
   MsgBox "Invalid Selection!!!"
   CmbBillingUnit.SetFocus
   Exit Sub
End If

If LCase(CmbBillingUnit.Text) = "bags" Then
   mFlag = "B"
ElseIf LCase(CmbBillingUnit.Text) = "quantity" Then
   mFlag = "Q"
Else
   mFlag = Null
End If
End Sub

Private Sub TxtNewRate_LostFocus()
If TxtNewRate.Text = "" Then Exit Sub

If IsNumeric(TxtNewRate.Text) = False Then
   MsgBox "Invalid Numeric Format."
   TxtNewRate.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) = "account" Then
   If Val(TxtNewRate) < Val(TxtPresentRate) Then
      MsgBox "You can not give rate less then present rate!!!"
      TxtNewRate.SetFocus
      Exit Sub
   End If
End If
End Sub

Private Function GetADhocInvoiceAmount(mLocID_ As Variant, mClientID_ As Variant)
Dim Retval As Variant
Retval = 0

tmp = "Select LocationID,ClientID,Sum(Amount) As 'Amount' From Txn_TempInvoice "
tmp = tmp & " Where IETypeID = 7 And LocationID = " & mLocID_ & " And ClientID = " & mClientID_ & " And Month(InvoiceDate) = " & mMonth & "  And Year(InvoiceDate) = " & CmbYear.Text & " And Isnull(ReverseFlag,'N') = 'N' And Flag = 'A' And Isnull(BatchID,0) <> 0 "
tmp = tmp & " Group By LocationID,ClientID "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1!Amount
End If

GetADhocInvoiceAmount = Retval

End Function
