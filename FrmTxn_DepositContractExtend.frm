VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_DepositContractExtend 
   Caption         =   "Extend Deposit Contract"
   ClientHeight    =   9270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
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
      Height          =   4845
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   15030
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
         TabIndex        =   3
         Top             =   4320
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
         TabIndex        =   4
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton CmbExcel 
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
         TabIndex        =   5
         Top             =   4320
         Width           =   1095
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
         Left            =   4800
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3720
         Width           =   10125
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1000
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
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   600
         Width           =   2010
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   195
         Width           =   2010
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1845
         Width           =   4770
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
         TabIndex        =   19
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1000
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   195
         Width           =   2655
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   6810
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1420
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
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2250
         Width           =   4770
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2670
         Width           =   4770
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2250
         Width           =   6810
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1420
         Width           =   2655
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1420
         Width           =   2655
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1845
         Width           =   6810
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3080
         Width           =   12885
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
         TabIndex        =   7
         Top             =   810
         Width           =   1290
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
      Begin MSComCtl2.DTPicker TxtNewEndDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   3840
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   111083521
         CurrentDate     =   39850
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
         Left            =   3600
         TabIndex        =   42
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Deposit Contract New End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   3720
         Width           =   1800
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
         TabIndex        =   37
         Top             =   225
         Width           =   1815
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
         Left            =   4200
         TabIndex        =   36
         Top             =   225
         Width           =   735
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
         TabIndex        =   35
         Top             =   225
         Width           =   735
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
         TabIndex        =   34
         Top             =   225
         Width           =   1440
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
         TabIndex        =   33
         Top             =   2483
         Width           =   1215
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
         Left            =   2640
         TabIndex        =   32
         Top             =   630
         Width           =   2160
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
         Height          =   210
         Index           =   0
         Left            =   2640
         TabIndex        =   31
         Top             =   1050
         Width           =   2160
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
         TabIndex        =   30
         Top             =   1448
         Width           =   1335
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
         TabIndex        =   29
         Top             =   1448
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
         TabIndex        =   28
         Top             =   1753
         Width           =   1800
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
         TabIndex        =   27
         Top             =   2263
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
         TabIndex        =   26
         Top             =   1028
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
         TabIndex        =   25
         Top             =   628
         Width           =   1440
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
         TabIndex        =   24
         Top             =   3215
         Width           =   1095
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
         TabIndex        =   23
         Top             =   2698
         Width           =   720
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
         TabIndex        =   22
         Top             =   1873
         Width           =   1440
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
         TabIndex        =   21
         Top             =   1448
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmTxn_DepositContractExtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCommodityID, mClientIDRow As Variant
Dim mGodownList As Variant, mMinDate As Variant, mMaxDate As Variant, mDCGodownDetailID As Variant

Private Sub CmbExcel_Click()
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
  TxtNewEndDate.Value = Date
  TxtRemark = ""
End Sub

Private Sub CmdClear_Click()
    TxtDepositContractID = ""
    Call Clear
    CmdSave.Enabled = False
End Sub

Private Sub CmdSave_Click()
If TxtDepositContractID.Text <> "" Then
    If CDate(TxtNewEndDate.Value) <= CDate(TxtEndDate.Text) Then
        MsgBox "Past Date and Same date not allowed."
        TxtNewEndDate.SetFocus
        Exit Sub
    End If
    
    If LCase(Left(TxtStatus, 1)) <> "a" And LCase(Left(TxtStatus, 1)) <> "s" Then
        MsgBox "Only 'System Approved' / 'Approved' Deposit Contract is allow to Extend."
        TxtNewEndDate.SetFocus
        Exit Sub
    End If
    
    If Trim(TxtRemark) = "" Then
        MsgBox "Blank Remark not allowed."
        TxtRemark.SetFocus
        Exit Sub
    End If
    
    If TxtDepositContractID.Text <> "" Then
    
        ''-- 02 Feb 2018
        mGodownList = "0"
        mDCGodownDetailID = "0"
        ''-- List of Connect Godown and Detail ID
        tmp = " Select Distinct GodownID from Txn_DepositContractGodownDetail Where DepConID = " & Val(TxtDepositContractID.Text)
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            For mR = 1 To TmpRs.RecordCount
                mGodownList = mGodownList & "," & TmpRs!GodownID
                'mDCGodownDetailID = mDCGodownDetailID & "," & TmpRs!DepConDID
                TmpRs.MoveNext
            Next mR
        End If
        
        ''-- 02 Feb 2018
        ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        tmp = " SELECT TDCGD.GodownId,TDCGD.DepConID,TDC.StartDate,TDC.EndDate FROM"
        tmp = tmp & " Txn_DepositContractGodownDetail AS TDCGD Inner Join Txn_DepositContract TDC On TDC.DepConID = TDCGD.DepConID"
        tmp = tmp & " where TDC.ApprovalFlag Not In ('R','C') And TDC.DepConID<>" & Val(TxtDepositContractID) & " And TDC.LocationID=" & mLocationID & " And TDC.CommodityID=" & mCommodityID & " And TDC.ClientID=" & mClientIDRow & ""
        ''tmp = tmp & " And (('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= EndDate) Or '" & TxtEndDate & "' >= StartDate And '" & TxtEndDate & "' <= EndDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtEndDate & "' >= EndDate))"
        tmp = tmp & " And (('" & TxtStartDate & "' >= StartDate And '" & TxtStartDate & "' <= EndDate) Or '" & TxtNewEndDate & "' >= StartDate And '" & TxtNewEndDate & "' <= EndDate Or ('" & TxtStartDate & "' <= StartDate And '" & TxtNewEndDate & "' >= EndDate))"
        tmp = tmp & " AND TDCGD.GodownID IN (" & mGodownList & ")"
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
           MsgBox "Deposit Contract already exist for selected godowns on given start date and end date!!!"
           Exit Sub
        End If
        ''------->>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            
        ''-- Finally change the date
        tmp = " UPDATE Txn_DepositContract SET EndDate = '" & Format(TxtNewEndDate.Value, "YYYY-MM-DD") & "',ApprovalRemarks = isnull(ApprovalRemarks,'')+'" & TxtRemark & "' ,Concurrency = GETDATE(),UpdatedDate = GETDATE(),UpdatedBy = '" & Gen_UserLoginID & "' Where DepConID = " & Val(TxtDepositContractID) & " And ApprovalFlag In ('A','S') "
        Call TmpTable
        MsgBox "Deposit Contract end date is extended."
        Call Clear
        
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
Frame1.Enabled = True
CmdSave.Enabled = False
End Sub

Private Sub TxtDepositContractID_LostFocus()
If Trim(TxtDepositContractID.Text) = "" Then Call Clear: Exit Sub
If IsClean(TxtDepositContractID) = False Then
    MsgBox "Invalid Input."
    TxtDepositContractID.SetFocus
End If
If IsNumeric(TxtDepositContractID) = False Then
    MsgBox "Invalid Input."
    TxtDepositContractID.SetFocus
End If
If Val(TxtDepositContractID) > 0 Then
    
    tmp = " Select TDC.DepConID,TDC.Date,MS.StateName,TDC.LocationID,ML.LocationName,Mco.Commodity,TDC.CommodityID,MC.ClientName,MC.ClientID,MC.ClientIDRow,MC.Address,MC.TelephoneNo,TDC.BagSizeFlag, "
    tmp = tmp & " MC.MobileNo,MC.EmailID,TDC.QuantumMT,MIB.InsuranceByName,TDC.StartDate,TDC.EndDate,MBC.BillingCycle As BillingCycle, MBCDep.BillingCycle As BillingCycleDep,TDC.ApprovalFlag,TDC.ApprovalRemarks"
    tmp = tmp & " From Txn_DepositContract TDC"
    tmp = tmp & " Inner Join Mst_State MS on TDC.StateID=MS.StateID"
    tmp = tmp & " Inner Join MST_Location ML on TDC.LocationID=ML.LocationID"
    tmp = tmp & " Inner Join MST_Commodity Mco on TDC.CommodityID=Mco.CommodityID"
    tmp = tmp & " Inner Join MST_Client MC on TDC.ClientID=MC.ClientIDRow"
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
        TxtCommodity = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        mCommodityID = IIf(IsNull(TmpRs!CommodityID), 0, TmpRs!CommodityID)
        TxtStartDate = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "DD-MMM-YYYY"))
        TxtEndDate = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "DD-MMM-YYYY"))
        TxtClient = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName & "#" & TmpRs!ClientID & "#" & TmpRs!ClientIDRow)
        mClientIDRow = IIf(IsNull(TmpRs!ClientIDRow), 0, TmpRs!ClientIDRow)
        TxtAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
        TxtContactNo = IIf(IsNull(TmpRs!TelephoneNo), "", TmpRs!TelephoneNo)
        TxtContactDetails = IIf(IsNull(TmpRs!MobileNo), "", TmpRs!MobileNo)
        TxtEmailId = IIf(IsNull(TmpRs!EmailID), "", TmpRs!EmailID)
        TxtBillingCycle = IIf(IsNull(TmpRs!BillingCycle), "", TmpRs!BillingCycle)
        TxtDepositBillingCycle = IIf(IsNull(TmpRs!BillingCycleDep), "", TmpRs!BillingCycleDep)
        TxtInsurance = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)
        TxtRemarks = IIf(IsNull(TmpRs!ApprovalRemarks), "", TmpRs!ApprovalRemarks)
        Check4.Value = IIf(IsNull(TmpRs!BagSizeFlag), vbUnchecked, IIf(LCase(TmpRs!BagSizeFlag) = "n", vbUnchecked, vbChecked))
        TxtNewEndDate.Value = CDate(TxtEndDate.Text) + 1
        
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "a" Then TxtStatus = "Approved": CmdSave.Enabled = True
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "s" Then TxtStatus = "System Approved": CmdSave.Enabled = True
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "f" Then TxtStatus = "Forwarded": CmdSave.Enabled = False
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "r" Then TxtStatus = "Rejected": CmdSave.Enabled = False
        If LCase(IIf(IsNull(TmpRs!ApprovalFlag), "", TmpRs!ApprovalFlag)) = "c" Then TxtStatus = "Closed": CmdSave.Enabled = False
        
    Else
        MsgBox "Detail Not Found."
        Exit Sub
    End If
End If
End Sub
