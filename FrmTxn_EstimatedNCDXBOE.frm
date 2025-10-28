VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_EstimatedNCDXBOE 
   Caption         =   "Estimated Storage Charges For NCDEX Withdrawal"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4995
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   4995
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   45
      Top             =   8520
      Width           =   14775
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11790
         TabIndex        =   47
         Top             =   195
         Width           =   2835
      End
      Begin VB.CommandButton GenerateCmd 
         Caption         =   "&Generate Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   8955
         TabIndex        =   46
         Top             =   195
         Width           =   2835
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4395
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   14820
      Begin VB.OptionButton OptOldDelivery 
         Caption         =   "Old Delivery"
         Height          =   375
         Left            =   9120
         TabIndex        =   56
         Top             =   165
         Width           =   1215
      End
      Begin VB.OptionButton OptNewDelivery 
         Caption         =   "New Delivery"
         Height          =   375
         Left            =   7560
         TabIndex        =   55
         Top             =   165
         Width           =   1335
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   1800
         TabIndex        =   51
         Top             =   1005
         Width           =   2880
      End
      Begin VB.TextBox TxtDPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2880
      End
      Begin VB.ComboBox CmbISINNo 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   2910
      End
      Begin VB.TextBox TxtWithdrawalQty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   3
         Top             =   3840
         Width           =   1965
      End
      Begin VB.CommandButton CmdTTDAdd2List 
         Caption         =   "Add to List"
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
         Left            =   7440
         TabIndex        =   4
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton CmdDeleteList 
         Caption         =   "Remove from List"
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
         Height          =   375
         Left            =   8535
         TabIndex        =   42
         Top             =   3840
         Width           =   1740
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CommTrack RRN"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   41
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox TxtDeliveredQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2880
         Width           =   1365
      End
      Begin VB.TextBox TxtRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   420
         Left            =   1800
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3360
         Width           =   8475
      End
      Begin VB.TextBox TxtOutOfSystemQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   5520
         TabIndex        =   8
         Top             =   2880
         Width           =   1800
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   37
         Top             =   4440
         Visible         =   0   'False
         Width           =   5580
      End
      Begin VB.TextBox TxtRRNQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   1800
         TabIndex        =   7
         Top             =   2880
         Width           =   1200
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3120
      End
      Begin VB.TextBox TxtISINDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   5520
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3120
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2520
         Width           =   8445
      End
      Begin VB.TextBox TxtCMPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2160
         Width           =   8445
      End
      Begin VB.TextBox TxtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3120
      End
      Begin VB.TextBox TxtDPName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1425
         Width           =   4125
      End
      Begin VB.TextBox TxtISINID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         TabIndex        =   6
         Top             =   1800
         Width           =   2880
      End
      Begin VB.TextBox TxtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1035
         Width           =   4125
      End
      Begin VB.TextBox TxtBalanceQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2880
         Width           =   3120
      End
      Begin VB.TextBox TxtRRNID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   2880
      End
      Begin VB.TextBox TxtRRNNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   0
         Top             =   180
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker RRNDate 
         Height          =   360
         Left            =   6120
         TabIndex        =   5
         Top             =   630
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   113967105
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker CITFDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   3840
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Format          =   113967105
         CurrentDate     =   39864
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client ID"
         Height          =   225
         Left            =   240
         TabIndex        =   53
         Top             =   1073
         Width           =   1800
      End
      Begin VB.Label lblNameDid 
         Caption         =   "ClientDPID"
         Height          =   225
         Left            =   255
         TabIndex        =   52
         Top             =   1475
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Withdrawal Quantity "
         Height          =   240
         Left            =   3840
         TabIndex        =   44
         Top             =   3907
         Width           =   1560
      End
      Begin VB.Label Label15 
         Caption         =   "CITF Date"
         Height          =   240
         Left            =   240
         TabIndex        =   43
         Top             =   3900
         Width           =   1800
      End
      Begin VB.Label Label12 
         Caption         =   "Remarks"
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   3480
         Width           =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "Qty Delived before StoreMan Starting Date"
         Height          =   405
         Left            =   3840
         TabIndex        =   38
         Top             =   2880
         Width           =   1680
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "Exchange Type"
         Height          =   300
         Left            =   10305
         TabIndex        =   34
         Top             =   2172
         Width           =   1440
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   2570
         Width           =   1200
      End
      Begin VB.Label LblCMPID 
         Caption         =   "CMP"
         Height          =   210
         Left            =   240
         TabIndex        =   32
         Top             =   2217
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Grade"
         Height          =   180
         Left            =   10320
         TabIndex        =   31
         Top             =   2592
         Width           =   900
      End
      Begin VB.Label LblCommodity 
         Caption         =   "Commodity"
         Height          =   240
         Left            =   10320
         TabIndex        =   30
         Top             =   1842
         Width           =   1140
      End
      Begin VB.Label Label8 
         Caption         =   "Client Name "
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   1088
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "DP ID"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1455
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Balance Qty in MT"
         Height          =   405
         Left            =   10320
         TabIndex        =   19
         Top             =   2840
         Width           =   840
      End
      Begin VB.Label LblRRNID 
         Caption         =   "RRN ID"
         Height          =   225
         Left            =   240
         TabIndex        =   11
         Top             =   668
         Width           =   1800
      End
      Begin VB.Label LblRRNNo 
         Caption         =   "RRN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   255
         Width           =   1800
      End
      Begin VB.Label LblRRNDate 
         Caption         =   "RRN Date"
         Height          =   180
         Left            =   5160
         TabIndex        =   14
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label LblISINID 
         Caption         =   "ISIN No"
         Height          =   225
         Left            =   240
         TabIndex        =   15
         Top             =   1845
         Width           =   1800
      End
      Begin VB.Label LblRRNQty 
         Caption         =   "RRN Qty in MT"
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   2948
         Width           =   1560
      End
      Begin VB.Label LblDeliveredQty 
         Caption         =   "Qty Delivered in StoreMan"
         Height          =   405
         Left            =   7620
         TabIndex        =   17
         Top             =   2858
         Width           =   1200
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "DP Name "
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   1460
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   3600
      Left            =   240
      TabIndex        =   48
      Top             =   4800
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   6350
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selected RRN Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   49
      Top             =   4560
      Width           =   14775
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   54
      Top             =   9360
      Visible         =   0   'False
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmTxn_EstimatedNCDXBOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mISINID, mClientID As String
Dim mISIN_LotID As String
Dim mClientIDRow As Double
Dim mCMPName, mCMPID, mLocationID, mCommodityID, mCommodityGroupId, mExchangeTypeID, mRRNID, mServiceTaxID As Double
Dim mSLocationID As Double
Dim x As Integer
Dim mFlag As Boolean
Dim mRecord As Boolean
Dim mEntryDate, mFedDate, mNE_Date As Variant
Dim mSumCITFQty, mSumSCToBerecovered As Double
Dim mLastRRNNo As Variant
Dim mDelRows, mCommodityType As Variant

Private Sub Check1_Click()
'Call Grid_Head3
'TxtRRNQty = ""
'If Check1.Value = 1 Then
'   CmbISINNo.Visible = True
'Else
'   CmbISINNo.Visible = False
'End If
End Sub



Private Sub CITFDate_LostFocus()
If OptNewDelivery.Value = True Then
    If CITFDate.Value < Date Then
        MsgBox "Back Date is not allowed"
        CITFDate.SetFocus
        Exit Sub
    End If
ElseIf OptOldDelivery.Value = True Then
    If CITFDate.Value > Date Then
        MsgBox "Future Date is not allowed"
        CITFDate.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub CmbISINNo_LostFocus()
If CmbISINNo = "" Then Exit Sub
TxtWithdrawalQty.Text = ""
If mRecord = True Then
   Call TableMst_RRN(" Where RRNNo = '" & Trim(TxtRRNNo) & "' And ISINID = '" & Trim(CmbISINNo) & "'")
End If
Call ShowRRNData(TxtRRNNo, CmbISINNo)
End Sub
Private Sub CmdDeleteList_Click()
Dim RO, C As Integer
'Dim mEnable As Boolean
mSumCITFQty = 0
mSumSCToBerecovered = 0
'mEnable = True
With MSHFlexGrid2
If mDelRows = 0 Then Exit Sub
If mDelRows > 0 Then
   If .TextMatrix(mDelRows, 0) <> "" Then
      For RO = mDelRows To .Rows - 2
          .row = RO
          For C = 0 To .Cols - 1
              .Col = C
              .TextMatrix(RO, C) = .TextMatrix(RO + 1, C)
          Next
      Next
      .Rows = .Rows - 1
   End If
   
End If
End With

For rv = 2 To MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.row = rv
    For cv = 0 To MSHFlexGrid2.Cols - 1
        MSHFlexGrid2.Col = cv
        If MSHFlexGrid2.CellBackColor = vbRed Then
           MSHFlexGrid2.CellBackColor = vbWhite
        End If
    Next
Next

For rv = 2 To MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.row = rv
    For cv = 0 To MSHFlexGrid2.Cols - 1
        MSHFlexGrid2.Col = cv
        If MSHFlexGrid2.TextMatrix(rv, 15) = "0" Then
           MSHFlexGrid2.CellBackColor = vbRed
           'mEnable = False
        End If
    Next
Next

'If mEnable = True Then GenerateCmd.Enabled = True

If MSHFlexGrid2.Rows > 3 Then
   For I = 2 To MSHFlexGrid2.Rows - 2
       mSumCITFQty = mSumCITFQty + MSHFlexGrid2.TextMatrix(I, 8)
       mSumSCToBerecovered = mSumSCToBerecovered + MSHFlexGrid2.TextMatrix(I, 10)
   Next
   MSHFlexGrid2.TextMatrix(0, 8) = mSumCITFQty
   MSHFlexGrid2.TextMatrix(0, 10) = mSumSCToBerecovered
Else
   MSHFlexGrid2.TextMatrix(0, 8) = "0.00"
   MSHFlexGrid2.TextMatrix(0, 10) = "0.00"
End If
CmdDeleteList.Enabled = False

End Sub


Private Sub CmdTTDAdd2List_Click()
Dim r As Variant
Dim mTotDays As Integer
Dim mGodown As Variant
Dim mSCToBerecovered, mPWRFED, mFEDCITF, mRRNFED As Long
Dim mMaxRate As Double
Dim mRateExist As Boolean

mRateExist = True
mMaxRate = 0
mSumCITFQty = 0
mSumSCToBerecovered = 0
mSCToBerecovered = 0
'GenerateCmd.Enabled = True
If TxtRRNNo = "" Then
   MsgBox "Blank RRN No!!! "
   TxtRRNNo.SetFocus
   Exit Sub
End If

If mFlag = True Then
   If CmbISINNo.Text = "" Then
      MsgBox "Blank ISIN No!!! "
      CmbISINNo.SetFocus
      Exit Sub
   End If
End If

If TxtWithdrawalQty = "" Then
   MsgBox "Blank Withdrawal Quantity is not allowed!!!"
   TxtWithdrawalQty.SetFocus
   Exit Sub
End If

If OptNewDelivery.Value = True Then
    If CITFDate.Value < Date Then
        MsgBox "Back Date is not Allowed"
        CITFDate.SetFocus
        Exit Sub
    End If
ElseIf OptOldDelivery.Value = True Then
    If CITFDate.Value > Date Then
        MsgBox "Future Date is not Allowed"
        CITFDate.SetFocus
        Exit Sub
    End If
End If

If MSHFlexGrid2.Rows > 3 Then
   For j = 2 To MSHFlexGrid2.Rows - 1
       If mFlag = False Then
          If UCase(MSHFlexGrid2.TextMatrix(j, 0)) = UCase(TxtRRNNo.Text) And MSHFlexGrid2.TextMatrix(j, 1) = TxtISINID.Text Then
             MsgBox "RRNo " & TxtRRNNo & " is Already exist in Grid!!!"
             'TxtRRNNo.SetFocus
            ' Call ClearControl
             Exit Sub
          End If
       Else
          If UCase(MSHFlexGrid2.TextMatrix(j, 0)) = UCase(TxtRRNNo.Text) And MSHFlexGrid2.TextMatrix(j, 1) = CmbISINNo.Text Then
             MsgBox "RRNo " & TxtRRNNo & " is Already exist in Grid!!!"
             'Call ClearControl
             Exit Sub
          End If
       End If
   Next
   If MSHFlexGrid2.TextMatrix(2, 16) <> mCommodityType Then
      If MSHFlexGrid2.TextMatrix(2, 16) = "1" Then
        MsgBox "Commtrack Commodities Data Exists in the Grid" & vbCrLf & "You have selected NComm Commodities"
        Exit Sub
      End If
      If MSHFlexGrid2.TextMatrix(2, 16) <> "1" Then
        MsgBox "NComm  Commodities Data Exists in the Grid" & vbCrLf & "You have selected Commtrack Commodities"
        Exit Sub
      End If
   End If
End If

'If CmdTTDAdd2List.Caption = "Update List" Then
'   r = MSHFlexGrid2.RowSel
''   MSHFlexGrid4.TextMatrix(0, 3) = Val(MSHFlexGrid4.TextMatrix(0, 3)) - Val(MSHFlexGrid4.TextMatrix(r, 3))
''   MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) - Val(MSHFlexGrid4.TextMatrix(r, 4))
''   MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) - Val(MSHFlexGrid4.TextMatrix(r, 5))
''   MSHFlexGrid4.TextMatrix(0, 2) = Val(MSHFlexGrid4.TextMatrix(0, 2)) - Val(MSHFlexGrid4.TextMatrix(r, 2))
'Else
'
'End If

'tmp = "Select  Max(Rate) as MaxRate From MSt_NcDEXStorageRateCardDetail Where NSRCID in("
'tmp = tmp & " Select NSRCID From MSt_NcDEXStorageRateCard Where LocationID = " & mLocationID & " And CommodityGroupID =" & mCommodityGroupId & ") And"
'tmp = tmp & " GodownID in(Select Distinct GodownID From Mst_GSL as GSL"
'tmp = tmp & " Where GSL.ExchangeTypeID = 2 And"
'tmp = tmp & " GSL.Flag Not in('E','Z','B') And"
'tmp = tmp & " GSL.CommodityID = " & mCommodityID & " And GSL.ISINID = '" & TxtISINID & "') "
'tmp = tmp & "Group by GodownID"
tmp = "Select  Max(Rate) as MaxRate From MSt_NcDEXStorageRateCardDetail Where NSRCID in("
tmp = tmp & " Select NSRCID From MSt_NcDEXStorageRateCard Where LocationID = " & mLocationID & " And CommodityGroupID =" & mCommodityGroupId & ")"
Call Tmp3Table
If Not TmpRs3.EOF Then
   mMaxRate = TmpRs3!MaxRate
Else
   mRateExist = False
   'GenerateCmd.Enabled = False
End If

r = MSHFlexGrid2.Rows - 1
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
MSHFlexGrid2.TextMatrix(r, 0) = TxtRRNNo
MSHFlexGrid2.TextMatrix(r, 1) = IIf(mFlag = True, CmbISINNo.Text, TxtISINID.Text)
MSHFlexGrid2.TextMatrix(r, 2) = Format(RRNDate.Value, "dd-MMM-yyyy")
MSHFlexGrid2.TextMatrix(r, 3) = TxtRRNQty
MSHFlexGrid2.TextMatrix(r, 4) = TxtBalanceQty
MSHFlexGrid2.TextMatrix(r, 5) = Format(mFedDate, "dd-MMM-yyyy")
MSHFlexGrid2.TextMatrix(r, 6) = Format(mNE_Date, "dd-MMM-yyyy")
MSHFlexGrid2.TextMatrix(r, 7) = Format(CITFDate.Value, "dd-MMM-yyyy")
MSHFlexGrid2.TextMatrix(r, 8) = TxtWithdrawalQty
MSHFlexGrid2.TextMatrix(r, 9) = TxtCommodity
MSHFlexGrid2.TextMatrix(r, 10) = ""
MSHFlexGrid2.TextMatrix(r, 11) = mLocationID
MSHFlexGrid2.TextMatrix(r, 12) = mClientID
MSHFlexGrid2.TextMatrix(r, 13) = mCommodityID
MSHFlexGrid2.TextMatrix(r, 14) = mServiceTaxID
MSHFlexGrid2.TextMatrix(r, 15) = mMaxRate
MSHFlexGrid2.TextMatrix(r, 16) = mCommodityType 'IIf(IsNull(mCommodityType), 0, 1)



If Check1.Value = 1 Then '''Commtrack Commodities
   If CITFDate.Value >= CDate(mFedDate) Then
      'RRNDate To FED Date Ò Single Charges From RRNDate to FEDDate
      'FEDDate to CITFDate Ò Double Charges From FEDDate to CITFDate
'''''       mTotDays = DateDiff("d", CDate(mFedDate), CITFDate.Value) + 1
'''''       mSCToBerecovered = (mTotDays * mMaxRate * Val(TxtWithdrawalQty)) * 2
'''''       MSHFlexGrid2.TextMatrix(r, 10) = Format(mSCToBerecovered, "#0.000")
      mTotDays = DateDiff("d", RRNDate.Value, CDate(mFedDate)) + 1
      mRRNFED = mTotDays * mMaxRate * Val(TxtWithdrawalQty)
         
      mTotDays = DateDiff("d", CDate(mFedDate), CITFDate.Value) + 1#
      
      'Introduced this line for month comparison on 13 Dec 2013 after mail from Mayur Sir and Dharmendra Sir
      If DateDiff("m", CDate(mFedDate), CITFDate.Value) > 0 Then
         mFEDCITF = (mTotDays * mMaxRate * Val(TxtWithdrawalQty)) * 2
      Else
         mFEDCITF = (mTotDays * mMaxRate * Val(TxtWithdrawalQty))
      End If
      mSCToBerecovered = mRRNFED + mFEDCITF
      MSHFlexGrid2.TextMatrix(r, 10) = Format(mSCToBerecovered, "#0.000")
   ElseIf CIFTDate <= CDate(mFedDate) Then
   'Single Charges From RRNDate to CITFDate then
      mTotDays = DateDiff("d", RRNDate.Value, CITFDate.Value) + 1
      mSCToBerecovered = (mTotDays * mMaxRate * Val(TxtWithdrawalQty))
      MSHFlexGrid2.TextMatrix(r, 10) = Format(mSCToBerecovered, "#0.000")
'   Else
'      MSHFlexGrid2.TextMatrix(r, 10) = Format(mSCToBerecovered, "#0.000")
   End If
Else 'NComm Commodities
   If CITFDate.Value <= CDate(mFedDate) Then
      MSHFlexGrid2.TextMatrix(r, 10) = "0.000" '''No Billing
   ElseIf CITFDate.Value >= CDate(mFedDate) Then
      tmp = "Select * From Txn_RRNPWR Where RRNID =" & mRRNID
      Call Tmp2Table
      If Not TmpRs2.EOF Then
         mTotDays = DateDiff("d", TmpRs2!PWRDate, CDate(mFedDate)) + 1
         mPWRFED = mTotDays * mMaxRate * Val(TxtWithdrawalQty)
         
         mTotDays = DateDiff("d", TmpRs2!PWRDate, CITFDate.Value) + 1
         If DateDiff("m", TmpRs2!PWRDate, CITFDate.Value) > 0 Then
            mFEDCITF = (mTotDays * mMaxRate * Val(TxtWithdrawalQty)) * 2
         Else
            mFEDCITF = (mTotDays * mMaxRate * Val(TxtWithdrawalQty))
         End If
         
         mSCToBerecovered = mPWRFED + mFEDCITF
         MSHFlexGrid2.TextMatrix(r, 10) = Format(mSCToBerecovered, "#0.000")
      Else
         'mTotDays = DateDiff("d", Format(mFedDate, Gen_DatFormat), CITFDate.Value) + 1
         mTotDays = DateDiff("d", CDate(mFedDate), CITFDate.Value) + 1
         mSCToBerecovered = mTotDays * mMaxRate * Val(TxtWithdrawalQty)
         MSHFlexGrid2.TextMatrix(r, 10) = Format(mSCToBerecovered, "#0.000")
      End If
   End If
End If

If mRateExist = False Then
   MSHFlexGrid2.row = r
   For Col = 0 To MSHFlexGrid2.Cols - 1
       MSHFlexGrid2.Col = Col
       MSHFlexGrid2.CellBackColor = vbRed
   Next
End If

For I = 2 To MSHFlexGrid2.Rows - 2
    mSumCITFQty = mSumCITFQty + MSHFlexGrid2.TextMatrix(I, 8)
    mSumSCToBerecovered = mSumSCToBerecovered + MSHFlexGrid2.TextMatrix(I, 10)
Next

MSHFlexGrid2.TextMatrix(0, 8) = mSumCITFQty
MSHFlexGrid2.TextMatrix(0, 10) = mSumSCToBerecovered
'If mFlag = False Then
Call ClearControl
'Else
'    TxtClientID.Text = ""
'    TxtClientName.Text = ""
'    TxtDPID.Text = ""
'    TxtDPName.Text = ""
'    TxtISINID.Text = ""
'    CmbISINNo.Text = ""
'    TxtISINDesc.Text = ""
'    TxtCMPID.Text = ""
'    TxtLocationID.Text = ""
'    TxtRRNQty.Text = ""
'    TxtOutOfSystemQty.Text = ""
'    TxtDeliveredQty.Text = ""
'    TxtRemarks.Text = ""
'    TxtWithdrawalQty.Text = ""
'    TxtCommodity.Text = ""
'    TxtExchangeTypeID.Text = ""
'    TxtGrade.Text = ""
'    TxtBalanceQty.Text = ""
'    Check1.Value = 0
'    RRNDate.Value = Date
'    CITFDate.Value = Date
'End If
CmdDeleteList.Enabled = False

End Sub



Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head3
'CITFDate.Value = Date
'RRNDate.Value = Date
OptNewDelivery.Value = True
mFlag = False
Call ClearControl

'If Check1.Value = 1 Then
'   CmbISINNo.Visible = True
'Else
'   CmbISINNo.Visible = False
'End If

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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "FrmTxn_EstimatedNCDXBOE.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        If MSHFlexGrid2.ColWidth(C) <> 0 Then
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



Private Sub GenerateCmd_Click()
Dim mFlag As Boolean
mFlag = False
If MSHFlexGrid2.Rows < 4 Then
   MsgBox "No Data Exists"
   Exit Sub
End If

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 15) = "0" Then
       MSHFlexGrid2.row = I
       For Col = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.Col = Col
           MSHFlexGrid2.CellBackColor = vbRed
           mFlag = True
       Next
    End If
Next

If mFlag = True Then
   MsgBox "Rate Not Found For Highlighted Entries"
   Exit Sub
End If

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FrmTxn_EstimatedNCDXBOE.xls")
Call Xls_Detail_Rpt



End Sub

Private Function GetMaxID() As Double
Dim Retval As Double
tmp = "Select max(PINo) as ID from Txn_TempInvoice"
Call Tmp0Table
Retval = IIf(IsNull(TmpRs0.Fields(0)), 0, TmpRs0.Fields(0))
Retval = Retval + 1

GetMaxID = Retval
End Function

Private Sub MSHFlexGrid2_Click()
mDelRows = MSHFlexGrid2.RowSel
CmdDeleteList.Enabled = True
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub TxtISINID_GotFocus()
tmp = TxtISINID.Text
End Sub

Private Sub TxtRRNNo_GotFocus()
If TxtRRNNo = "" Then Exit Sub
mFlag = False
Call ClearControl
End Sub

Private Sub TxtRRNNo_LostFocus()
'If TxtRRNNo <> mLastRRNNo Then ClearControl
If TxtRRNNo = "" Then
   Call ClearControl
   'TxtRRNNo.SetFocus
   Exit Sub
End If
mLastRRNNo = TxtRRNNo
TxtRRNNo = UCase(Trim(TxtRRNNo))
Call TableMst_RRN(" Where RRNNo = '" & Trim(TxtRRNNo) & "'")

If RsMst_RRN.RecordCount = 0 Then
   MsgBox "No RRN No Found"
   Call ClearControl
   TxtRRNNo.SetFocus
   Exit Sub
End If

mFlag = False
mRecord = False
TxtWithdrawalQty.Text = ""
'RsMst_RRN.MoveFirst

If RsMst_RRN!CommTrackFlag = 1 Then
   CmbISINNo.Visible = True
   CmbISINNo.Clear
   TableMst_RRN_LotDetail (" Where RRNID = " & RsMst_RRN!RRNID)
   If RsMst_RRNLotDetail.RecordCount > 0 Then
      For I = 1 To RsMst_RRNLotDetail.RecordCount
          CmbISINNo.AddItem RsMst_RRNLotDetail!ISINID
          RsMst_RRNLotDetail.MoveNext
      Next
      mFlag = True
    'CmbISINNo.SetFocus
   End If
Else
   If RsMst_RRN.RecordCount > 1 Then
      CmbISINNo.Visible = True
      CmbISINNo.Clear
      For I = 1 To RsMst_RRN.RecordCount
          CmbISINNo.AddItem RsMst_RRN!ISINID
          RsMst_RRN.MoveNext
      Next
      mFlag = True
      mRecord = True
      'CmbISINNo.SetFocus
   End If
End If
If mFlag = False Then
   CmbISINNo.Visible = False
   CmbISINNo.Clear
End If
If mFlag = True Then Exit Sub
Call ShowRRNData(TxtRRNNo)
End Sub
Private Sub ShowRRNData(ByVal mRRNo As String, Optional ByVal mISINNo As String = "")
Dim wc As String

wc = ""
mEntryDate = ""
mFedDate = ""
mNE_Date = ""
mLocationID = ""
mCommodityID = ""
mCommodityType = ""

tmp = "Select RRN.RRNID, RRN.RRNNo, RRN.RRNDate, RRN.ISINID, ISIN.ISINDescription,"
tmp = tmp & " RRN.RRNQty, RRN.DeliveredQty,  (RRN.RRNQty - RRN.DeliveredQty) As BalanceRRNQty, RRN.OutOfSystemQty,"
tmp = tmp & " ISIN.FED, ISIN.NE_Date, ISIN.Grade,ISIN.ExchangeTypeID, RRN.Remarks, RRN.EntryDate, RRN.Flag, RRN.CommTrackFlag,"
tmp = tmp & " ML.LocationID, ML.LocationName, CMP.CMPName, COM.CommodityID, COM.CommodityGroupID, COM.Commodity, COM.ServiceTaxID,"
tmp = tmp & " MC.ClientID , MC.ClientName, DP.DepositoryID, DP.DepositoryName"
If RsMst_RRN!CommTrackFlag = "1" Then
   tmp = tmp & " ,RRND.ISINID"
End If
tmp = tmp & " From Mst_RRN as RRN"
tmp = tmp & " Inner Join Mst_ISIN as ISIN on ISIN.ISINID = RRN.ISINID"
tmp = tmp & " Inner Join Mst_CMP as CMP on CMP.CMPID = ISIN.CMPID"
tmp = tmp & " Inner Join mst_location as  ML on ML.LocationID = CMP.LocationID"
tmp = tmp & " Inner Join Mst_Commodity as COM on COM.CommodityID = ISIN.CommodityID"
tmp = tmp & " Inner Join Mst_Client as MC on  MC.ClientIDRow = RRN.ClientIDRow"
tmp = tmp & " Inner Join Mst_Depository as DP on DP.DepositoryID = MC.DepositoryID and DP.ExchangeTypeID = 2"
If RsMst_RRN!CommTrackFlag = "1" Then
   tmp = tmp & " Inner Join Mst_RRN_LotDetail as RRND on RRND.RRNID = RRN.RRNID"
End If

If mRRNo <> "" Then
   wc = wc & " RRN.RRNNo = '" & mRRNo & "'"
End If

If mISINNo <> "" And RsMst_RRN!CommTrackFlag <> "1" Then
   wc = wc & " And RRN.ISINID = '" & mISINNo & "'"
End If

If mISINNo <> "" And RsMst_RRN!CommTrackFlag <> "0" Then
   wc = wc & " And RRND.ISINID = '" & mISINNo & "'"
End If

tmp = tmp & " Where " & wc
tmp = tmp & " Order By RRN.RRNNo, CMP.CMPName, COM.Commodity"

Call Tmp1Table

If Not TmpRs1.EOF Then
   TxtRRNID = TmpRs1!RRNID
   mRRNID = TmpRs1!RRNID
   TxtClientID = TmpRs1!ClientID
   mClientID = TxtClientID
   TxtClientName = TmpRs1!ClientName
   TxtDPID = TmpRs1!DepositoryID
   TxtDPName = TmpRs1!DepositoryName
   TxtISINID = TmpRs1!ISINID
   TxtISINDesc = TmpRs1!ISINDescription
   TxtCMPID = TmpRs1!CMPName
   TxtLocationID = TmpRs1!LocationName
   mLocationID = TmpRs1!LocationID
   TxtCommodity = TmpRs1!Commodity
   mCommodityID = TmpRs1!CommodityID
   mCommodityGroupId = TmpRs1!CommodityGroupID
   mServiceTaxID = TmpRs1!ServiceTaxID
   TxtExchangeTypeID = TmpRs1!ExchangeTypeID
   TxtGrade = TmpRs1!Grade
   TxtBalanceQty = TmpRs1!BalanceRRNQty
   TxtRRNQty = TmpRs1!RRNQty
   TxtOutOfSystemQty = TmpRs1!OutOfSystemQty
   TxtDeliveredQty = TmpRs1!DeliveredQty
   TxtRemarks = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
   Check1.Value = TmpRs1!CommTrackFlag
   mEntryDate = Format(TmpRs1!EntryDate, Gen_DatFormat)
   mFedDate = Format(TmpRs1!FED, Gen_DatFormat)
   mNE_Date = Format(TmpRs1!NE_Date, Gen_DatFormat)
   RRNDate = Format(TmpRs1!RRNDate, Gen_DatFormat)
   TxtDPName = TmpRs1!DepositoryName
   mCommodityType = TmpRs1!CommTrackFlag
Else
   MsgBox "Data Not Found"
   Call ClearControl
   Exit Sub
End If
End Sub

Private Sub Grid_Head3()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedCols = 0
MSHFlexGrid2.Cols = 17

MSHFlexGrid2.TextMatrix(1, 0) = "RRN No"
MSHFlexGrid2.TextMatrix(1, 1) = "ISIN No"
MSHFlexGrid2.TextMatrix(1, 2) = "RRN Date"
MSHFlexGrid2.TextMatrix(1, 3) = "RRN Qty"
MSHFlexGrid2.TextMatrix(1, 4) = "Balance RRN Qty"
MSHFlexGrid2.TextMatrix(1, 5) = "FED Date"
MSHFlexGrid2.TextMatrix(1, 6) = "NE Date"
MSHFlexGrid2.TextMatrix(1, 7) = "CITF Date"
MSHFlexGrid2.TextMatrix(1, 8) = "CITF Qty"
MSHFlexGrid2.TextMatrix(1, 9) = "Commodity"
MSHFlexGrid2.TextMatrix(1, 10) = "Storage Charge to be recovered"
MSHFlexGrid2.TextMatrix(1, 11) = "LocationID"
MSHFlexGrid2.TextMatrix(1, 12) = "ClientID"
MSHFlexGrid2.TextMatrix(1, 13) = "CommodityID"
MSHFlexGrid2.TextMatrix(1, 14) = "ServiceTaxID"
MSHFlexGrid2.TextMatrix(1, 15) = "MaxRate"
MSHFlexGrid2.TextMatrix(1, 16) = "Flag"


'MSHFlexGrid2.TextMatrix(1, 9) = "Storage Charges"
'MSHFlexGrid2.TextMatrix(1, 10) = "FED Charge"
'MSHFlexGrid2.TextMatrix(1, 11) = "Amount"
    
MSHFlexGrid2.ColWidth(0) = 1200
MSHFlexGrid2.ColWidth(1) = 1200
MSHFlexGrid2.ColWidth(2) = 1000
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 1500
MSHFlexGrid2.ColWidth(5) = 1000
MSHFlexGrid2.ColWidth(6) = 1000
MSHFlexGrid2.ColWidth(7) = 1000
MSHFlexGrid2.ColWidth(8) = 1200
MSHFlexGrid2.ColWidth(9) = 2500
MSHFlexGrid2.ColWidth(10) = 2500
'MSHFlexGrid2.ColWidth(10) = 1200
'MSHFlexGrid2.ColWidth(11) = 1200
'MSHFlexGrid2.ColWidth(12) = 1500
'MSHFlexGrid2.ColWidth(13) = 1500
'MSHFlexGrid2.ColWidth(14) = 1500
'MSHFlexGrid2.ColWidth(15) = 1000
MSHFlexGrid2.ColWidth(11) = 0
MSHFlexGrid2.ColWidth(12) = 0
MSHFlexGrid2.ColWidth(13) = 0
MSHFlexGrid2.ColWidth(14) = 0
MSHFlexGrid2.ColWidth(15) = 0
MSHFlexGrid2.ColWidth(16) = 0

End Sub

Private Sub ClearControl()
If mFlag = False Then
   TxtRRNNo.Text = ""
   TxtRRNID.Text = ""
End If
TxtClientID.Text = ""
TxtClientName.Text = ""
TxtDPID.Text = ""
TxtDPName.Text = ""
TxtISINID.Text = ""
CmbISINNo.Text = ""
TxtISINDesc.Text = ""
TxtCMPID.Text = ""
TxtLocationID.Text = ""
TxtRRNQty.Text = ""
TxtOutOfSystemQty.Text = ""
TxtDeliveredQty.Text = ""
TxtRemarks.Text = ""
TxtWithdrawalQty.Text = ""
TxtCommodity.Text = ""
TxtExchangeTypeID.Text = ""
TxtGrade.Text = ""
TxtBalanceQty.Text = ""
Check1.Value = 0
RRNDate.Value = Date
CITFDate.Value = Date
'mFlag = False
'CmbISINNo.Visible = False
End Sub

Private Sub TxtWithdrawalQty_LostFocus()

If TxtWithdrawalQty = "" Then Exit Sub

If IsNumeric(TxtWithdrawalQty) = False Or Val(TxtWithdrawalQty) < 0 Then
   MsgBox "Invalid Numeric value !!!"
   TxtWithdrawalQty.SetFocus
   Exit Sub
End If

If OptNewDelivery.Value = True Then
    If Val(TxtWithdrawalQty) > Val(TxtBalanceQty) Then
       MsgBox "WithDrawal Quantity should not be greater than Balance Quantity " & TxtBalanceQty & " !!!"
       TxtWithdrawalQty.SetFocus
       Exit Sub
    End If
End If

End Sub
