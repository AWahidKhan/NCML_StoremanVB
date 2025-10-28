VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_DepositContractChangeFlag 
   Caption         =   "Deposit Contract Change"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   15030
      Begin VB.ComboBox cmbNewStatus 
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
         ItemData        =   "FrmTxn_DepositContractChangeFlag.frx":0000
         Left            =   4920
         List            =   "FrmTxn_DepositContractChangeFlag.frx":0007
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   3720
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.TextBox TxtDepositContractID 
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
         Height          =   310
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   195
         Width           =   1635
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "Bag Wise"
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   195
         TabIndex        =   24
         Top             =   810
         Width           =   1290
      End
      Begin VB.TextBox TxtRemarks 
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
         Height          =   525
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3080
         Width           =   12885
      End
      Begin VB.TextBox TxtEmailId 
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
         Height          =   310
         Left            =   8115
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1845
         Width           =   6810
      End
      Begin VB.TextBox TxtContactNo 
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
         Height          =   310
         Left            =   8115
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1420
         Width           =   2655
      End
      Begin VB.TextBox TxtContactDetails 
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
         Height          =   310
         Left            =   12285
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1420
         Width           =   2655
      End
      Begin VB.TextBox TxtAddress 
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
         Height          =   720
         Left            =   8115
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2250
         Width           =   6810
      End
      Begin VB.TextBox TxtStatus 
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
         Height          =   310
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2670
         Width           =   4770
      End
      Begin VB.TextBox TxtInsurance 
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
         Height          =   310
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2250
         Width           =   4770
      End
      Begin VB.TextBox TxtBillingCycle 
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
         Height          =   310
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1420
         Width           =   4770
      End
      Begin VB.TextBox TxtClient 
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
         Height          =   310
         Left            =   8115
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   6810
      End
      Begin VB.TextBox TxtState 
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
         Height          =   310
         Left            =   8115
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   195
         Width           =   2655
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
         Height          =   310
         Left            =   8115
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1000
         Width           =   6810
      End
      Begin VB.TextBox TxtLocation 
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
         Height          =   310
         Left            =   12285
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   195
         Width           =   2655
      End
      Begin VB.TextBox TxtDepositBillingCycle 
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
         Height          =   310
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1845
         Width           =   4770
      End
      Begin VB.TextBox TxtDepositContractDate 
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
         Height          =   310
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   195
         Width           =   2010
      End
      Begin VB.TextBox TxtStartDate 
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
         Height          =   310
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   2010
      End
      Begin VB.TextBox TxtEndDate 
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
         Height          =   310
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1000
         Width           =   2010
      End
      Begin VB.TextBox TxtRemark 
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
         Height          =   525
         Left            =   2040
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4200
         Width           =   12885
      End
      Begin VB.CommandButton CmbClose 
         Caption         =   "Close"
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
         Left            =   13800
         TabIndex        =   7
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   12720
         TabIndex        =   6
         Top             =   4800
         Width           =   1095
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
         Left            =   11640
         TabIndex        =   4
         Top             =   4800
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker TxtNewStartDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   3720
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   661
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
         Format          =   111935489
         CurrentDate     =   39850
      End
      Begin VB.Label lblNewStatus 
         Caption         =   "New Status"
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
         Left            =   3600
         TabIndex        =   44
         Top             =   3773
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "Contact Details"
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
         Left            =   10875
         TabIndex        =   43
         Top             =   1448
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Email ID "
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
         Left            =   6945
         TabIndex        =   42
         Top             =   1873
         Width           =   1440
      End
      Begin VB.Label Label35 
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
         Height          =   255
         Left            =   105
         TabIndex        =   41
         Top             =   2698
         Width           =   720
      End
      Begin VB.Label Label33 
         Caption         =   "Remarks"
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
         Left            =   105
         TabIndex        =   40
         Top             =   3215
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Client"
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
         Left            =   6945
         TabIndex        =   39
         Top             =   628
         Width           =   1440
      End
      Begin VB.Label Label28 
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
         Height          =   255
         Index           =   0
         Left            =   6915
         TabIndex        =   38
         Top             =   1028
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Insurance By"
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
         Left            =   105
         TabIndex        =   37
         Top             =   2263
         Width           =   1440
      End
      Begin VB.Label Label16 
         Caption         =   "Deposit/ Withdrawal Billing Cycle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   36
         Top             =   1753
         Width           =   1800
      End
      Begin VB.Label Label15 
         Caption         =   "Contact No"
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
         Left            =   6945
         TabIndex        =   35
         Top             =   1448
         Width           =   1440
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
         Height          =   255
         Left            =   105
         TabIndex        =   34
         Top             =   1448
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Deposit Contract End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2040
         TabIndex        =   33
         Top             =   990
         Width           =   2640
      End
      Begin VB.Label Label13 
         Caption         =   "Deposit Contract Start Date"
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
         Index           =   0
         Left            =   2040
         TabIndex        =   32
         Top             =   630
         Width           =   2640
      End
      Begin VB.Label Label6 
         Caption         =   "Address"
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
         Left            =   6975
         TabIndex        =   31
         Top             =   2483
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   10905
         TabIndex        =   30
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label Label3 
         Caption         =   "State"
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
         Left            =   6945
         TabIndex        =   29
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
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
         Left            =   3840
         TabIndex        =   28
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Deposit Contract ID"
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
         Left            =   120
         TabIndex        =   27
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label lblNewStartDate 
         Caption         =   "New Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   3802
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
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
         TabIndex        =   25
         Top             =   4335
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmTxn_DepositContractChangeFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCommodityID, mClientIDRow As Variant, mSM_Prefix As Variant
Dim mFlagChangeFlag As Variant, mDateChangeFlag As Variant
Dim mFlagChange As Boolean, mDateChange As Boolean
Dim mGodownList As Variant, mMinDate As Variant, mMaxDate As Variant, mDCGodownDetailID As Variant

Private Sub CmbClose_Click()
 Unload Me
End Sub

Sub Clear()
  'TxtDepositContractID = ""
  TxtDepositContractDate = ""
  TxtState = ""
  TxtLocation = ""
  Check4.Value = vbUnchecked
  TxtStartDate = ""
  TxtEndDate = ""
  TxtClient = ""
  TxtCommodity = ""
  TxtBillingCycle = ""
  TxtDepositBillingCycle = ""
  TxtContactNo = ""
  TxtContactDetails = ""
  TxtEmailId = ""
  TxtAddress = ""
  TxtInsurance = ""
  TxtStatus = ""
  TxtRemarks = ""
  TxtRemark = ""
  
  If mFlagChange = True Then
    cmbNewStatus.Text = ""
  End If
  If mDateChange = True Then
    TxtNewStartDate.Value = CDate(Date)
  End If
  
If mDateChange = True Then
    lblNewStartDate.Visible = True
    TxtNewStartDate.Visible = True: TxtNewStartDate.Value = CDate(Date)
Else
    lblNewStartDate.Visible = False
    TxtNewStartDate.Visible = False
End If

If mFlagChange = True Then
    lblNewStatus.Visible = True
    cmbNewStatus.Visible = True: cmbNewStatus.Text = ""
Else
    lblNewStatus.Visible = False
    cmbNewStatus.Visible = False
End If
  
End Sub

Private Sub cmbNewStatus_LostFocus()
    If cmbNewStatus.Text = "" Then Exit Sub
    If LCase(cmbNewStatus.Text) <> "forwarded" Then
       MsgBox "Invalid Selection !!!"
       cmbNewStatus.SetFocus
       Exit Sub
    End If
End Sub

Private Sub CmdClear_Click()
    TxtDepositContractID = ""
    Call Clear
    CmdSave.Enabled = False
End Sub

Private Sub CmdSave_Click()
Dim mColumnData As Variant
Dim mR As Double

If mFlagChange = False And mDateChange = False Then Exit Sub

If TxtDepositContractID.Text <> "" Then

    mGodownList = "0"
    mDCGodownDetailID = "0"
    
    ''-- List of Connect Godown and Detail ID
    tmp = " Select * from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID.Text)
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        For mR = 1 To TmpRs.RecordCount
            mGodownList = mGodownList & "," & TmpRs!GodownID
            mDCGodownDetailID = mDCGodownDetailID & "," & TmpRs!DepConDID
            TmpRs.MoveNext
        Next mR
    End If
    
    ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ''-- Check weather the Detail ID is attach in CXTF if yes then stop
    If mDCGodownDetailID <> "" Then
        
        ''tmp = " Select DepConDID from Txn_CxTF_GSL Where SM_Prefix = '" & mSM_Prefix & "' And TxnType = 'D' And DepConDID In (" & mDCGodownDetailID & ") "
        If cmbNewStatus.Text <> "" Then
''-- Azure Changes By Ullhas On 30 Jan 2020
''            tmp = " Select DGSl.GSLDetailID,DGSL.SM_Prefix,DGSL.CxTFNo,DGSL.DepConDID,DDetail.CxTFDate from Txn_CxTF_GSL DGSL "
''            tmp = tmp & " Inner Join Txn_CxTF_Details DDetail On DGSL.CxTFNo = DDetail.CxTFNo And DGSL.TxnType = DDetail.TxnType And DGSL.SM_Prefix = DDetail.SM_Prefix "
''            tmp = tmp & " Where DGSL.SM_Prefix = '" & mSM_Prefix & "' And DGSL.TxnType = 'D' And DDetail.CxTFDate > '2017-11-30' And DGSL.DepConDID In (" & mDCGodownDetailID & ") "
            tmp = " Select DGSl.GSLDetailID,DGSL.SM_Prefix,DGSL.CxTFNo,DGSL.DepConDID,DDetail.CxTFDate from Txn_CxTF_GSL DGSL "
            tmp = tmp & " Inner Join Txn_CxTF_Details DDetail On DGSL.CxTF_DetailsID = DDetail.CxTF_DetailsID And DGSL.TxnType = DDetail.TxnType And DGSL.SM_Prefix = DDetail.SM_Prefix "
            tmp = tmp & " Where DGSL.SM_Prefix = '" & mSM_Prefix & "' And DGSL.TxnType = 'D' And DDetail.CxTFDate > '2017-11-30' And DGSL.DepConDID In (" & mDCGodownDetailID & ") "
            Call TmpTable
            If TmpRs.RecordCount > 0 Then
                MsgBox "Already deposit entry present for above DC . Can't change status"
                'Exit Sub
                mFlagChange = False
            End If
        End If
    End If
    ''-- Check weather the Detail ID is attach in GSL if yes then stop
    If mDCGodownDetailID <> "" Then
        If cmbNewStatus.Text <> "" Then
            tmp = " Select Distinct DepConDID from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' And CommodityID = " & mCommodityID & " And DepConDID In (" & mDCGodownDetailID & ") "
            Call TmpTable
            If TmpRs.RecordCount > 0 Then
                MsgBox "Already above DC mapped to GSL(s) . Can't change status"
                mFlagChange = False
                'Exit Sub
            End If
        End If
    End If
    ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    tmp = " Select YearMonth from Txn_InvMonGSLDetail Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & Format(TxtNewStartDate.Value, "YYYY-MM-01") & "' And YearMonth > '2017-11-01' Group By YearMonth "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        MsgBox "Opening Balance process is executed for given month, so system not allowed for changing Start date and Status.", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    tmp = " Select YearMonth from Txn_InvMonGSLDetail2016 Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & Format(TxtNewStartDate.Value, "YYYY-MM-01") & "' And YearMonth > '2017-11-01' Group By YearMonth "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        MsgBox "Opening Balance process is executed for given month, so system not allowed for changing Start date and Status.", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    tmp = " Select YearMonth from Txn_InvMonGSLDetail2014to2015 Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & Format(TxtNewStartDate.Value, "YYYY-MM-01") & "' And YearMonth > '2017-11-01' Group By YearMonth "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        MsgBox "Opening Balance process is executed for given month, so system not allowed for changing Start date and Status.", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    tmp = " Select YearMonth from Txn_InvMonGSLDetailUpto2013 Where SM_Prefix = '" & mSM_Prefix & "' And YearMonth = '" & Format(TxtNewStartDate.Value, "YYYY-MM-01") & "' And YearMonth > '2017-11-01' Group By YearMonth "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        MsgBox "Opening Balance process is executed for given month, so system not allowed for changing Start date and Status.", vbOKOnly + vbCritical
        Exit Sub
    End If
    ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    If mFlagChange = True Or mDateChange = True Then
        If mFlagChange = True Then
           If LCase(Left(TxtStatus, 1)) <> "a" And LCase(Left(TxtStatus, 1)) <> "s" And LCase(Left(TxtStatus, 1)) <> "f" Then
              MsgBox "Only 'System Approved' / 'Approved' / 'Forwarded' Deposit Contract is allow to change."
              Exit Sub
           End If
        End If
        If mDateChange = True Then
            If CDate(TxtNewStartDate.Value) > Date Then
                MsgBox "Future Date is not allowed."
                TxtNewStartDate.SetFocus
                Exit Sub
            End If
            If CDate(TxtNewStartDate.Value) > CDate(TxtStartDate.Text) Then
                MsgBox "Date should be less then previous start date!!!"
                TxtNewStartDate.SetFocus
                Exit Sub
            End If
        End If
        If Trim(TxtRemark) = "" Then
            MsgBox "Blank Remark not allowed."
            TxtRemark.SetFocus
            Exit Sub
        End If
        
        
''        ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''        ''-- if Same Deposit Contarct is available in backdate with same godown list then not allow to change or overlap not allow
''        mColumnData = ""
''        mMinDate = ""
''        mMaxDate = ""
''
''        tmp = " Select Min(TDC.StartDate)Min From Txn_DepositContractGodownDetail TDCD Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID "
''        tmp = tmp & " Where TDC.ApprovalFlag in ('F','S','A')  And TDC.LocationID = " & mLocationID & " And TDC.ClientID = " & mClientIDRow & " And TDC.CommodityID = " & mCommodityID & " "
''        tmp = tmp & " And TDC.DepConID <> " & Val(TxtDepositContractID.Text) & " And TDCD.GodownId In (" & mGodownList & ") "
''        Call TmpTable
''        If TmpRs.RecordCount > 0 Then
''            mMinDate = IIf(IsNull(TmpRs!Min), "", TmpRs!Min)
''        Else
''            mMinDate = ""
''        End If
''
''        tmp = " Select Max(TDC.EndDate)Max From Txn_DepositContractGodownDetail TDCD Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID"
''        tmp = tmp & " Where TDC.ApprovalFlag in ('F','S','A')  And TDC.LocationID = " & mLocationID & " And TDC.ClientID = " & mClientIDRow & " And TDC.CommodityID = " & mCommodityID & " "
''        tmp = tmp & " And TDC.DepConID <> " & Val(TxtDepositContractID.Text) & " And TDCD.GodownId In (" & mGodownList & ") "
''        Call TmpTable
''        If TmpRs.RecordCount > 0 Then
''            mMaxDate = IIf(IsNull(TmpRs!Max), "", TmpRs!Max)
''        Else
''            mMaxDate = ""
''        End If
''        ''-- if Date Change allow and Start date is take as backdated then
''        If mDateChange = True Then
''            If CDate(TxtNewStartDate.Value) < CDate(TxtStartDate.Text) Then
''
''                If mMinDate <> "" And mMaxDate <> "" Then
''                    If Format(TxtNewStartDate, "YYYY-MM-DD") >= Format(mMinDate, "YYYY-MM-DD") And Format(TxtNewStartDate, "YYYY-MM-DD") <= Format(mMaxDate, "YYYY-MM-DD") Then
''                        MsgBox "Deposit Contract is already available in between given New Start Date.", vbOKOnly + vbInformation
''                        Exit Sub
''                    End If
''                    If Format(TxtNewStartDate, "YYYY-MM-DD") < Format(mMinDate, "YYYY-MM-DD") Then
''                        MsgBox "Deposit Contract is already available in between given New Start Date.", vbOKOnly + vbInformation
''                        Exit Sub
''                    End If
''                End If
''
''                mColumnData = " StartDate = '" & Format(TxtNewStartDate.Value, "YYYY-MM-DD") & "', "
''
''            End If
''        End If
''        ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
        ''-- 02 Feb 2018
        ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        If mDateChange = True Then
            tmp = " SELECT TDCGD.GodownId,TDCGD.DepConID,TDC.StartDate,TDC.EndDate FROM"
            tmp = tmp & " Txn_DepositContractGodownDetail AS TDCGD Inner Join Txn_DepositContract TDC On TDC.DepConID = TDCGD.DepConID"
            tmp = tmp & " where TDC.ApprovalFlag Not In ('R','C') And TDC.DepConID<>" & Val(TxtDepositContractID) & " And TDC.LocationID=" & mLocationID & " And TDC.CommodityID=" & mCommodityID & " And TDC.ClientID=" & mClientIDRow & ""
            ''tmp = tmp & " And (('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= EndDate) Or '" & TxtEndDate & "' >= StartDate And '" & TxtEndDate & "' <= EndDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtEndDate & "' >= EndDate))"
            tmp = tmp & " And (('" & TxtNewStartDate & "' >= StartDate And '" & TxtNewStartDate & "' <= EndDate) Or '" & TxtEndDate & "' >= StartDate And '" & TxtEndDate & "' <= EndDate Or ('" & TxtNewStartDate & "' <= StartDate And '" & TxtEndDate & "' >= EndDate))"
            tmp = tmp & " AND TDCGD.GodownID IN (" & mGodownList & ")"
            Call TmpTable
            If TmpRs.RecordCount > 0 Then
               MsgBox "Deposit Contract already exist for selected godowns on given start date and end date!!!"
               Exit Sub
            End If
            mColumnData = " StartDate = '" & Format(TxtNewStartDate.Value, "YYYY-MM-DD") & "', "
        End If
        ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

        ''-- if Flag Change allow and Flag are not blank then
        If mFlagChange = True Then
            If cmbNewStatus.Text <> "" Then
                If LCase(Left(TxtStatus, 1)) <> "f" Then
                    If mColumnData <> "" Then
                        mColumnData = mColumnData & " ApprovalFlag = 'F',Verified = 'N', "
                    Else
                        mColumnData = " ApprovalFlag = 'F',Verified = 'N', "
                    End If
                End If
            End If
        End If
        
        If mColumnData <> "" Then
           tmp = " UPDATE Txn_DepositContract SET " & mColumnData & " ApprovalRemarks = isnull(ApprovalRemarks,'')+' " & TxtRemark & "' ,Concurrency = GETDATE(),UpdatedDate = GETDATE(),UpdatedBy = '" & Gen_UserLoginID & "' Where DepConID = " & Val(TxtDepositContractID) & " And ApprovalFlag In ('" & LCase(Left(TxtStatus, 1)) & "') "
           Call TmpTable
           MsgBox "Done."
           Call Clear
        Else
           MsgBox "Change the Flag / Date to update."
        End If
        
    End If
   
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" Then
    mFlagChangeFlag = "": mDateChangeFlag = ""
    ''-- Permission are read from config table
    mFlagChangeFlag = Gen_DepositContractCalculationFactor("DepositContractChangeFlag-")
    mDateChangeFlag = Gen_DepositContractCalculationFactor("DepositContractChangeDate-")
    If LCase(mFlagChangeFlag) = "" Or LCase(mFlagChangeFlag) = "n" Then
        mFlagChange = False
    Else
        mFlagChange = True
    End If
    If LCase(mDateChangeFlag) = "" Or LCase(mDateChangeFlag) = "n" Then
        mDateChange = False
    Else
        mDateChange = True
    End If
Else
    mFlagChange = True
    mDateChange = True
End If

If mDateChange = True Then
    lblNewStartDate.Visible = True
    TxtNewStartDate.Visible = True: TxtNewStartDate.Value = CDate(Date)
Else
    lblNewStartDate.Visible = False
    TxtNewStartDate.Visible = False
End If

If mFlagChange = True Then
    lblNewStatus.Visible = True
    cmbNewStatus.Visible = True: cmbNewStatus.Text = ""
Else
    lblNewStatus.Visible = False
    cmbNewStatus.Visible = False
End If

Frame1.Enabled = True
CmdSave.Enabled = False
End Sub


Private Sub TxtDepositContractID_LostFocus()
If Trim(TxtDepositContractID.Text) = "" Then Call Clear: Exit Sub
If IsClean(TxtDepositContractID) = False Then
    MsgBox "Invalid Inpout."
    TxtDepositContractID.SetFocus
End If
If IsNumeric(TxtDepositContractID) = False Then
    MsgBox "Invalid Inpout."
    TxtDepositContractID.SetFocus
End If
If Val(TxtDepositContractID) > 0 Then
    
    tmp = " Select TDC.DepConID,TDC.Date,MS.StateName,ML.SM_Prefix,TDC.LocationID,ML.LocationName,Mco.Commodity,TDC.CommodityID,MC.ClientName,MC.ClientID,MC.ClientIDRow,MC.Address,MC.TelephoneNo,TDC.BagSizeFlag, "
    tmp = tmp & " MC.MobileNo,MC.EmailID,TDC.QuantumMT,MIB.InsuranceByName,TDC.StartDate,TDC.EndDate,MBC.BillingCycle As BillingCycle, MBCDep.BillingCycle As BillingCycleDep,TDC.ApprovalFlag,TDC.ApprovalRemarks"
    tmp = tmp & " From Txn_DepositContract TDC"
    tmp = tmp & " Inner Join Mst_State MS on TDC.StateID=MS.StateID"
    tmp = tmp & " Inner Join MST_Location ML on TDC.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join MST_Commodity Mco on TDC.CommodityID=Mco.CommodityID"
    ''-- Azure Changes By Ullhas On 30 Jan 2020
    ''tmp = tmp & " Inner Join MST_Client MC on TDC.ClientID=MC.ClientIDRow"
    tmp = tmp & " Inner Join MST_Client MC on TDC.ClientID=MC.ClientID"
    tmp = tmp & " Inner Join MST_BillingCycle MBC on TDC.BillingCycleID=MBC.BillingCycleID"
    tmp = tmp & " Inner Join MST_BillingCycle MBCDep on TDC.DepBillingCycleID=MBCDep.BillingCycleID"
    tmp = tmp & " Inner Join Mst_InsuranceBy MIB on TDC.InsuranceByID=MIB.InsuranceByID"
    tmp = tmp & " Where TDC.DepConID = " & TxtDepositContractID & " "
    Call TmpTable
    Call Clear
    CmdSave.Enabled = False
    If TmpRs.RecordCount > 0 Then
        TxtDepositContractDate = IIf(IsNull(TmpRs!Date), "", Format(TmpRs!Date, "DD-MMM-YYYY"))
        TxtState = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        TxtLocation = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
        mLocationID = IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID)
        mSM_Prefix = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
        TxtCommodity = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
        TxtStartDate = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "DD-MMM-YYYY"))
        TxtEndDate = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "DD-MMM-YYYY"))
        TxtClient = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName & "#" & TmpRs!ClientID & "#" & TmpRs!ClientIDRow)
        ''-- Azure Changes By Ullhas On 30 Jan 2020
        ''mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
        mClientIDRow = IIf(IsNull(TmpRs!ClientID), 0, TmpRs!ClientID)
        TxtAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
        TxtContactNo = IIf(IsNull(TmpRs!TelephoneNo), "", TmpRs!TelephoneNo)
        TxtContactDetails = IIf(IsNull(TmpRs!MobileNo), "", TmpRs!MobileNo)
        TxtEmailId = IIf(IsNull(TmpRs!EmailID), "", TmpRs!EmailID)
        TxtBillingCycle = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)
        TxtDepositBillingCycle = IIf(IsNull(TmpRs!BillingCycleDep), "", TmpRs!BillingCycleDep)
        TxtInsurance = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)
        TxtRemarks = IIf(IsNull(TmpRs!ApprovalRemarks), "", TmpRs!ApprovalRemarks)
        Check4.Value = IIf(IsNull(TmpRs!BagSizeFlag), vbUnchecked, IIf(LCase(TmpRs!BagSizeFlag) = "n", vbUnchecked, vbChecked))
        
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "a" Then TxtStatus = "Approved"
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "s" Then TxtStatus = "System Approved"
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "f" Then TxtStatus = "Forwarded"
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "r" Then TxtStatus = "Rejected"
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "c" Then TxtStatus = "Closed"
        
        If mFlagChange = True Then
            cmbNewStatus.Text = ""
        End If
        If mDateChange = True Then
            TxtNewStartDate.Value = CDate(TxtStartDate.Text)
        End If
        
        If mFlagChange = False And mDateChange = False Then
            CmdSave.Enabled = False
        Else
            CmdSave.Enabled = True
        End If
        
        If LCase(Left(TxtStatus, 1)) = "f" Then
            If mDateChange = True Then
                cmbNewStatus = "": cmbNewStatus.Enabled = False
            End If
        Else
            If mDateChange = True Then
                cmbNewStatus = "": cmbNewStatus.Enabled = True
            End If
        End If
        
    Else
        CmdSave.Enabled = False
        MsgBox "Detail Not Found."
        Exit Sub
    End If
  
End If
End Sub
