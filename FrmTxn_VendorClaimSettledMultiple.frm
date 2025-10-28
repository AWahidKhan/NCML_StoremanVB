VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_VendorClaimSettledMultiple 
   Caption         =   "Vendor Claim Settel Multiple"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Frame0 
      BorderStyle     =   0  'None
      Height          =   9660
      Left            =   60
      ScaleHeight     =   9660
      ScaleWidth      =   15090
      TabIndex        =   8
      Top             =   30
      Width           =   15090
      Begin VB.OptionButton OptFumigationBill 
         Caption         =   "Fumigation Bill"
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
         Left            =   6120
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton OptSecurityBill 
         Caption         =   "Security Bill"
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
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton OptFranchiseeFee 
         Caption         =   "Franchisee Fee"
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
         Left            =   2040
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin VB.OptionButton OptGodownRent 
         Caption         =   "Godown Rent"
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
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox TxtErrMsg 
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
         Height          =   840
         Left            =   1050
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   8385
         Width           =   13980
      End
      Begin VB.ComboBox CmbModeOfPay 
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
         ItemData        =   "FrmTxn_VendorClaimSettledMultiple.frx":0000
         Left            =   240
         List            =   "FrmTxn_VendorClaimSettledMultiple.frx":0010
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox CmbFAGBank 
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
         Left            =   4545
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox TxtDisbuAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   4200
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtDeductionRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2880
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtDisbuRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1680
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker TxtDisbuDate 
         Height          =   360
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2145
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   48300033
         CurrentDate     =   36494
      End
      Begin VB.CheckBox ChkCheckAll 
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
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   165
         TabIndex        =   23
         Top             =   1245
         Width           =   210
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
         Height          =   350
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1212
         Width           =   1000
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   45
         TabIndex        =   22
         Top             =   9345
         Width           =   14970
         _ExtentX        =   26405
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
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
         Height          =   350
         Left            =   14010
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1212
         Width           =   1000
      End
      Begin VB.ComboBox CmbSVendorID 
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
         Left            =   3090
         Sorted          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   800
         Width           =   6000
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
         Height          =   350
         Left            =   10995
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1212
         Width           =   1000
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
         Height          =   350
         Left            =   13005
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1212
         Width           =   1000
      End
      Begin VB.TextBox TxtSClaimID 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   4050
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1425
      End
      Begin VB.ComboBox CmbSForwardTo 
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
         Left            =   9135
         Sorted          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   800
         Width           =   5880
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6660
         Left            =   45
         TabIndex        =   14
         Top             =   1650
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   11748
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         WordWrap        =   -1  'True
         AllowUserResizing=   3
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
      Begin MSComCtl2.DTPicker STo 
         Height          =   360
         Left            =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   800
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   48300033
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   800
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   48300033
         CurrentDate     =   36494
      End
      Begin VB.Label Label1 
         Caption         =   "Error Message"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   60
         TabIndex        =   32
         Top             =   8505
         Width           =   915
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   465
         TabIndex        =   24
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "--"
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
         Left            =   8370
         TabIndex        =   21
         Top             =   1260
         Width           =   165
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
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
         Left            =   375
         TabIndex        =   20
         Top             =   485
         Width           =   945
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
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
         Left            =   1965
         TabIndex        =   19
         Top             =   485
         Width           =   735
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Name"
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
         Left            =   5460
         TabIndex        =   18
         Top             =   485
         Width           =   1260
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
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
         Left            =   5760
         TabIndex        =   17
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Label19 
         Caption         =   "Claim ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3090
         TabIndex        =   16
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label Label31 
         Caption         =   "Forwarded To"
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
         Left            =   11385
         TabIndex        =   15
         Top             =   485
         Width           =   1395
      End
   End
End
Attribute VB_Name = "FrmTxn_VendorClaimSettledMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSVendorID, mSForwardToID, tmp1 As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mSelRow, mSelCol, mFAGBankID As Variant

Private Sub ChkCheckAll_Click()

If MSHFlexGrid1.Rows = 2 And MSHFlexGrid1.TextMatrix(1, 3) = "" Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       If MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked Then
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

Private Sub CmdClear_Click()
Call Grid_Head1
SFrom.Value = Date - 30
STo.Value = Date 'To_GenDate
CmbSVendorID.Text = ""
CmbSForwardTo.Text = ""
TxtSClaimID = ""
Label22.Caption = ""
ProgressBar1.Value = 0
End Sub

Private Sub CmdSave_Click()

TxtErrMsg.Text = ""
mErrorMsg = ""

If ChkGridSelection = False Then
   MsgBox "No Claim selected !!!"
   Exit Sub
End If

If OptGodownRent.Value = True Then
    ''-- Godown Rent
    ''-------------------------------
    Call TableTxn_GodownRent(" Where BatchID > 0 And Flag = 'A' ")
    If RsTxn_GodownRent.RecordCount = 0 Then
       MsgBox "No Claim present for settelment!!!"
       Exit Sub
    End If
    CmdSave.Enabled = False
    With MSHFlexGrid1
        For I = 2 To .Rows - 2
            If .TextMatrix(I, 0) = strChecked Then
               If .TextMatrix(I, 20) = "" Then
                    mErrorMsg = mErrorMsg & "Blank Bank-A/C No not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                    GoTo NextGodownRow
               End If
               If .TextMatrix(I, 17) = "" Then
                  mErrorMsg = mErrorMsg & "Blank  Disbursement Remark not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextGodownRow
               End If
               If .TextMatrix(I, 19) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Mode of Payment not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextGodownRow
               End If
               If LCase(.TextMatrix(I, 19)) = "bank transfer" Or LCase(.TextMatrix(I, 19)) = "neft" Or LCase(.TextMatrix(I, 19)) = "rtgs" Then
                  If .TextMatrix(I, 23) = "" Then
                     mErrorMsg = mErrorMsg & "Blank Account Details not allowed for 'Bank Transfer', Contact FAG to update Vendor Account details for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                     GoTo NextGodownRow
                  End If
               End If
               If .TextMatrix(I, 18) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Disbursment Amount not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextGodownRow
               End If
               If Val(.TextMatrix(I, 18)) > 0 Then
                  If Val(.TextMatrix(I, 18)) < Val(.TextMatrix(I, 6)) Then
                     If Trim(.TextMatrix(I, 21)) = "" Then
                        mErrorMsg = mErrorMsg & "As disbursment amount is less than claim amount please provide deduction remarks for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                        GoTo NextGodownRow
                     End If
                  End If
               End If
            
               If RsTxn_GodownRent.RecordCount > 0 Then RsTxn_GodownRent.MoveFirst
               RsTxn_GodownRent.Find "ClaimID = " & Val(.TextMatrix(I, 1)) & ""
               If RsTxn_GodownRent.EOF = False Then
                  RsTxn_GodownRent!Flag = "S"
                  RsTxn_GodownRent!DisbuDate = CDate(.TextMatrix(I, 16))
                  RsTxn_GodownRent!DisbuRemark = Trim(.TextMatrix(I, 17))
                  RsTxn_GodownRent!DisbuAmount = Val(.TextMatrix(I, 18))
                  RsTxn_GodownRent!ModeOfPay = Trim(.TextMatrix(I, 19))
                  RsTxn_GodownRent!FAGBankID = Val(.TextMatrix(I, 22))
                  RsTxn_GodownRent!DeductionRemark = Trim(.TextMatrix(I, 21))
                  RsTxn_GodownRent!UpdatedBy = Gen_UserLoginID
                  RsTxn_GodownRent!UpdatedDate = Now
                  RsTxn_GodownRent.Update
               End If
            End If
NextGodownRow:
        Next
    End With
    ''-------------------------------
    
ElseIf OptFranchiseeFee.Value = True Then
    ''-- FranchiseeFee
    ''-------------------------------
    Call TableTxn_FranchiseeRent(" Where BatchID > 0 And Flag = 'A' ")
    If RsTxn_FranchiseeRent.RecordCount = 0 Then
       MsgBox "No Claim present for settelment!!!"
       Exit Sub
    End If
    CmdSave.Enabled = False
    With MSHFlexGrid1
        For I = 2 To .Rows - 2
            If .TextMatrix(I, 0) = strChecked Then
               If .TextMatrix(I, 20) = "" Then
                    mErrorMsg = mErrorMsg & "Blank Bank-A/C No not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                    GoTo NextFranchiseeRow
               End If
               If .TextMatrix(I, 17) = "" Then
                  mErrorMsg = mErrorMsg & "Blank  Disbursement Remark not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextFranchiseeRow
               End If
               If .TextMatrix(I, 19) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Mode of Payment not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextFranchiseeRow
               End If
               If LCase(.TextMatrix(I, 19)) = "bank transfer" Or LCase(.TextMatrix(I, 19)) = "neft" Or LCase(.TextMatrix(I, 19)) = "rtgs" Then
                  If .TextMatrix(I, 23) = "" Then
                     mErrorMsg = mErrorMsg & "Blank Account Details not allowed for 'Bank Transfer', Contact FAG to update Vendor Account details for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                     GoTo NextFranchiseeRow
                  End If
               End If
               If .TextMatrix(I, 18) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Disbursment Amount not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextFranchiseeRow
               End If
               If Val(.TextMatrix(I, 18)) > 0 Then
                  If Val(.TextMatrix(I, 18)) < Val(.TextMatrix(I, 6)) Then
                     If Trim(.TextMatrix(I, 21)) = "" Then
                        mErrorMsg = mErrorMsg & "As disbursment amount is less than claim amount please provide deduction remarks for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                        GoTo NextFranchiseeRow
                     End If
                  End If
               End If
            
               If RsTxn_FranchiseeRent.RecordCount > 0 Then RsTxn_FranchiseeRent.MoveFirst
               RsTxn_FranchiseeRent.Find "ClaimID = " & Val(.TextMatrix(I, 1)) & ""
               If RsTxn_FranchiseeRent.EOF = False Then
                  RsTxn_FranchiseeRent!Flag = "S"
                  RsTxn_FranchiseeRent!DisbuDate = CDate(.TextMatrix(I, 16))
                  RsTxn_FranchiseeRent!DisbuRemark = Trim(.TextMatrix(I, 17))
                  RsTxn_FranchiseeRent!DisbuAmount = Val(.TextMatrix(I, 18))
                  RsTxn_FranchiseeRent!ModeOfPay = Trim(.TextMatrix(I, 19))
                  RsTxn_FranchiseeRent!FAGBankID = Val(.TextMatrix(I, 22))
                  RsTxn_FranchiseeRent!DeductionRemark = Trim(.TextMatrix(I, 21))
                  RsTxn_FranchiseeRent!UpdatedBy = Gen_UserLoginID
                  RsTxn_FranchiseeRent!UpdatedDate = Now
                  RsTxn_FranchiseeRent.Update
               End If
            End If
NextFranchiseeRow:
        Next
    End With
    ''-------------------------------
    
ElseIf OptSecurityBill.Value = True Then
    ''-- SecurityBill
    ''-------------------------------
    Call TableTxn_SecurityBilling(" Where BatchID > 0 And Status = 'A' ")
    If RsTxn_SecurityBilling.RecordCount = 0 Then
       MsgBox "No Claim present for settelment!!!"
       Exit Sub
    End If
    CmdSave.Enabled = False
    With MSHFlexGrid1
        For I = 2 To .Rows - 2
            If .TextMatrix(I, 0) = strChecked Then
               If .TextMatrix(I, 20) = "" Then
                    mErrorMsg = mErrorMsg & "Blank Bank-A/C No not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                    GoTo NextSecurityRow
               End If
               If .TextMatrix(I, 17) = "" Then
                  mErrorMsg = mErrorMsg & "Blank  Disbursement Remark not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextSecurityRow
               End If
               If .TextMatrix(I, 19) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Mode of Payment not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextSecurityRow
               End If
               If LCase(.TextMatrix(I, 19)) = "bank transfer" Or LCase(.TextMatrix(I, 19)) = "neft" Or LCase(.TextMatrix(I, 19)) = "rtgs" Then
                  If .TextMatrix(I, 23) = "" Then
                     mErrorMsg = mErrorMsg & "Blank Account Details not allowed for 'Bank Transfer', Contact FAG to update Vendor Account details for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                     GoTo NextSecurityRow
                  End If
               End If
               If .TextMatrix(I, 18) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Disbursment Amount not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextSecurityRow
               End If
               If Val(.TextMatrix(I, 18)) > 0 Then
                  If Val(.TextMatrix(I, 18)) < Val(.TextMatrix(I, 6)) Then
                     If Trim(.TextMatrix(I, 21)) = "" Then
                        mErrorMsg = mErrorMsg & "As disbursment amount is less than claim amount please provide deduction remarks for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                        GoTo NextSecurityRow
                     End If
                  End If
               End If
            
               If RsTxn_SecurityBilling.RecordCount > 0 Then RsTxn_SecurityBilling.MoveFirst
               RsTxn_SecurityBilling.Find "BillTxnID = " & Val(.TextMatrix(I, 1)) & ""
               If RsTxn_SecurityBilling.EOF = False Then
                  RsTxn_SecurityBilling!Status = "S"
                  RsTxn_SecurityBilling!DisbuDate = CDate(.TextMatrix(I, 16))
                  RsTxn_SecurityBilling!DisbuRemark = Trim(.TextMatrix(I, 17))
                  RsTxn_SecurityBilling!DisbuAmount = Val(.TextMatrix(I, 18))
                  RsTxn_SecurityBilling!ModeOfPay = Trim(.TextMatrix(I, 19))
                  RsTxn_SecurityBilling!FAGBankID = Val(.TextMatrix(I, 22))
                  RsTxn_SecurityBilling!DeductionRemark = Trim(.TextMatrix(I, 21))
                  RsTxn_SecurityBilling!UpdatedBy = Gen_UserLoginID
                  RsTxn_SecurityBilling!UpdatedDate = Now
                  RsTxn_SecurityBilling.Update
               End If
            End If
NextSecurityRow:
        Next
    End With
    ''-------------------------------
    
ElseIf OptFumigationBill.Value = True Then
    ''-- FumigationBill
    ''-------------------------------
    Call TableTxn_FumigationApproval(" Where BatchID > 0 And Flag = 'A' ")
    If RsTxn_FumigationApproval.RecordCount = 0 Then
       MsgBox "No Claim present for settelment!!!"
       Exit Sub
    End If
    CmdSave.Enabled = False
    With MSHFlexGrid1
        For I = 2 To .Rows - 2
            If .TextMatrix(I, 0) = strChecked Then
               If .TextMatrix(I, 20) = "" Then
                    mErrorMsg = mErrorMsg & "Blank Bank-A/C No not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                    GoTo NextFumigationRow
               End If
               If .TextMatrix(I, 17) = "" Then
                  mErrorMsg = mErrorMsg & "Blank  Disbursement Remark not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextFumigationRow
               End If
               If .TextMatrix(I, 19) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Mode of Payment not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextFumigationRow
               End If
               If LCase(.TextMatrix(I, 19)) = "bank transfer" Or LCase(.TextMatrix(I, 19)) = "neft" Or LCase(.TextMatrix(I, 19)) = "rtgs" Then
                  If .TextMatrix(I, 23) = "" Then
                     mErrorMsg = mErrorMsg & "Blank Account Details not allowed for 'Bank Transfer', Contact FAG to update Vendor Account details for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                     GoTo NextFumigationRow
                  End If
               End If
               If .TextMatrix(I, 18) = "" Then
                  mErrorMsg = mErrorMsg & "Blank Disbursment Amount not allowed for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                  GoTo NextFumigationRow
               End If
               If Val(.TextMatrix(I, 18)) > 0 Then
                  If Val(.TextMatrix(I, 18)) < Val(.TextMatrix(I, 6)) Then
                     If Trim(.TextMatrix(I, 21)) = "" Then
                        mErrorMsg = mErrorMsg & "As disbursment amount is less than claim amount please provide deduction remarks for Claim ID " & .TextMatrix(I, 1) & "" & vbCrLf
                        GoTo NextFumigationRow
                     End If
                  End If
               End If
            
               If RsTxn_FumigationApproval.RecordCount > 0 Then RsTxn_FumigationApproval.MoveFirst
               RsTxn_FumigationApproval.Find "MappingID = " & Val(.TextMatrix(I, 1)) & ""
               If RsTxn_FumigationApproval.EOF = False Then
                  RsTxn_FumigationApproval!Flag = "S"
                  RsTxn_FumigationApproval!DisbuDate = CDate(.TextMatrix(I, 16))
                  RsTxn_FumigationApproval!DisbuRemark = Trim(.TextMatrix(I, 17))
                  RsTxn_FumigationApproval!DisbuAmount = Val(.TextMatrix(I, 18))
                  RsTxn_FumigationApproval!ModeOfPay = Trim(.TextMatrix(I, 19))
                  RsTxn_FumigationApproval!FAGBankID = Val(.TextMatrix(I, 22))
                  RsTxn_FumigationApproval!DeductionRemark = Trim(.TextMatrix(I, 21))
                  RsTxn_FumigationApproval!UpdatedBy = Gen_UserLoginID
                  RsTxn_FumigationApproval!UpdatedDate = Now
                  RsTxn_FumigationApproval.Update
               End If
            End If
NextFumigationRow:
        Next
    End With
    ''-------------------------------
    
End If

TxtErrMsg.Text = mErrorMsg
MsgBox "Done!!!"
Call Grid_Head1
Label22.Refresh
Label22.Caption = ""
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied!!!"
   Frame0.Enabled = False
   Exit Sub
End If
Frame0.Enabled = True
Call CmdClear_Click
TxtErrMsg = ""
End Sub

Private Sub Grid_Head1()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 24 '20
    .WordWrap = True

    .TextMatrix(1, 0) = "Select":                           .ColWidth(0) = 800
    .TextMatrix(1, 1) = "Claim ID":                         .ColWidth(1) = 1000
    .TextMatrix(1, 2) = "Claim Date":                       .ColWidth(2) = 1400
    .TextMatrix(1, 3) = "Expense From":                     .ColWidth(3) = 1400
    .TextMatrix(1, 4) = "Expense To":                       .ColWidth(4) = 1400
    .TextMatrix(1, 5) = "Vendor Name":                      .ColWidth(5) = 3000
    .TextMatrix(1, 6) = "Claim Amount":                     .ColWidth(6) = 1400
    .TextMatrix(1, 7) = "Remarks":                          .ColWidth(7) = 1200
    .TextMatrix(1, 8) = "Status":                           .ColWidth(8) = 0
    .TextMatrix(1, 9) = "G_UID":                            .ColWidth(9) = 0
    .TextMatrix(1, 10) = "FAG Remarks":                     .ColWidth(10) = 1500
    .TextMatrix(1, 11) = "Booked in NAV Date":              .ColWidth(11) = 1200
    .TextMatrix(1, 12) = "Claim Forwarded To":              .ColWidth(12) = 0
    .TextMatrix(1, 13) = "Invoice No":                      .ColWidth(13) = 0
    .TextMatrix(1, 14) = "Invoice Date":                    .ColWidth(14) = 0
    .TextMatrix(1, 15) = "Batch No":                        .ColWidth(15) = 0
    .TextMatrix(1, 16) = "Disbursment Date":                .ColWidth(16) = 1500
    .TextMatrix(1, 17) = "Disbursment Remarks":             .ColWidth(17) = 3000
    .TextMatrix(1, 18) = "Disbursment Amount":              .ColWidth(18) = 2000
    .TextMatrix(1, 19) = "Mode for Payment":                .ColWidth(19) = 1800
    .TextMatrix(1, 20) = "Disbursment through Bank":        .ColWidth(20) = 3000
    .TextMatrix(1, 21) = "Deduction Remark":                .ColWidth(21) = 3000
    .TextMatrix(1, 22) = "FAG Bank ID":                     .ColWidth(22) = 0
    .TextMatrix(1, 23) = "Vendor Bank Acc":                 .ColWidth(23) = 2500
    .RowHeight(1) = 600
End With
  
End Sub

Private Sub CmbSVendorID_GotFocus()
tmp = CmbSVendorID.Text
CmbSVendorID.Clear
If OptGodownRent.Value = True Then
    Call TableMst_Lessor
    If RsMst_Lessor.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To RsMst_Lessor.RecordCount
        CmbSVendorID.AddItem RsMst_Lessor!LessorName
        RsMst_Lessor.MoveNext
    Next
ElseIf OptFranchiseeFee.Value = True Or OptFumigationBill.Value = True Then
    Call TableMst_Vendor
    If RsMst_Vendor.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To RsMst_Vendor.RecordCount
        CmbSVendorID.AddItem RsMst_Vendor!VendorName
        RsMst_Vendor.MoveNext
    Next
ElseIf OptSecurityBill.Value = True Then
    Call TableMst_SecurityAgency(" Where Flag='Y'")
    If RsMst_SecurityAgency.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To RsMst_SecurityAgency.RecordCount
        CmbSVendorID.AddItem RsMst_SecurityAgency!AgencyName
        RsMst_SecurityAgency.MoveNext
    Next
End If

CmbSVendorID.Text = tmp
End Sub

Private Sub CmbSVendorID_LostFocus()
If CmbSVendorID.Text = "" Then
   Exit Sub
End If

If OptGodownRent.Value = True Then
    RsMst_Lessor.MoveFirst
    RsMst_Lessor.Find "LessorName = '" & CmbSVendorID.Text & "'"
    If RsMst_Lessor.EOF Then
       MsgBox "Invalid selection "
       CmbSVendorID.SetFocus
       Exit Sub
    End If
    mSVendorID = RsMst_Lessor!LessorID
ElseIf OptFranchiseeFee.Value = True Or OptFumigationBill.Value = True Then
    RsMst_Vendor.MoveFirst
    RsMst_Vendor.Find "VendorName = '" & CmbSVendorID.Text & "'"
    If RsMst_Vendor.EOF Then
       MsgBox "Invalid selection "
       CmbSVendorID.SetFocus
       Exit Sub
    End If
    mSVendorID = RsMst_Vendor!VendorID
ElseIf OptSecurityBill.Value = True Then
    RsMst_SecurityAgency.MoveFirst
    RsMst_SecurityAgency.Find "AgencyName = '" & CmbSVendorID.Text & "'"
    If RsMst_SecurityAgency.EOF Then
       MsgBox "Invalid selection "
       CmbSVendorID.SetFocus
       Exit Sub
    End If
    mSVendorID = RsMst_SecurityAgency!SecurityAgencyID
End If

End Sub

Private Sub CmbSForwardTo_GotFocus()
tmp = CmbSForwardTo.Text
Call TableMst_Employee(" Where Flag = 'Y' Order By EmployeeName ")
CmbSForwardTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSForwardTo.SetFocus
   'RsMst_Employee.MoveFirst
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSForwardTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSForwardTo.Text = tmp

End Sub

Private Sub CmbSForwardTo_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSForwardTo.Text = "" Then
   mSForwardToID = Null
   Exit Sub
End If

strSplitString = Split(CmbSForwardTo, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid selection!!!"
   CmbForwardTo.SetFocus
   Exit Sub
End If
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSForwardTo.SetFocus
   Exit Sub
End If
mSForwardToID = RsMst_Employee!EmployeeID
End Sub
Private Sub CmdGo_Click()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim wc As Variant

Call Grid_Head1
TxtErrMsg = ""
If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To date'!!!"
   SFrom.SetFocus
   Exit Sub
End If
wc = ""

If OptGodownRent.Value = True Then
    tmp = "SELECT TGR.ClaimID,TGR.ClaimDate,TGR.ExpenseFrom,TGR.ExpenseTo,TGR.LessorID,ML.LessorName[VendorName],TGR.Amount,TGR.Remarks,TGR.Flag,"
    tmp = tmp & " TGR.G_UID,TGR.ApprovalRemarks,TGR.ApprovedDate,EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as ForwardedTo,TGR.BatchID,TGR.DisbuDate,TGR.DisbuRemark,TGR.DisbuAmount,TGR.ModeOfPay,MFB.BankName,TGR.DeductionRemark,MFB.FAGBankID,ML.AccountNo as 'VendorACCNo' "
    tmp = tmp & " FROM Txn_GodownRent TGR"
    tmp = tmp & " Inner Join Mst_Lessor ML On TGR.LessorID = ML.LessorID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP1 On TGR.ForwardToID = EMP1.EmployeeID"
    tmp = tmp & " Left Join Mst_FAGBank MFB On TGR.FAGBankID = MFB.FAGBankID"
    If Gen_DBType = "MDB" Then
       tmp = tmp & " Where TGR.BatchID > 0 And TGR.Flag = 'A' And TGR.ClaimDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
    Else
       tmp = tmp & " Where TGR.BatchID > 0 And TGR.Flag = 'A' And TGR.ClaimDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
    End If
    If CmbSVendorID.Text <> "" Then
       tmp = tmp & " AND ML.LessorID = " & mSVendorID
    End If
    If CmbSForwardTo.Text <> "" Then
       strSplitString = Split(CmbSForwardTo, "~")
       EmpName_ = Trim(strSplitString(0))
       EmpNo_ = Trim(strSplitString(1))
       tmp = tmp & " AND EMP1.EmployeeNo = '" & EmpNo_ & "' "
    End If
    If TxtSClaimID <> "" Then
       tmp = tmp & " AND TGR.ClaimID Like " & TxtSClaimID & " "
    End If
    tmp = tmp & " Order By TGR.ClaimID"
ElseIf OptFranchiseeFee.Value = True Then
    tmp = "Select TFR.ClaimID,TFR.ClaimDate,TFR.ExpenseFrom,TFR.ExpenseTo,TFR.VendorID,EMP.VendorName,TFR.Amount,TFR.Remarks,TFR.Flag,"
    tmp = tmp & " TFR.G_UID,TFR.ApprovalRemarks,TFR.ApprovedDate,EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as ForwardedTo,TFR.BatchID,TFR.DisbuDate,TFR.DisbuRemark,TFR.DisbuAmount,TFR.ModeOfPay,MFB.BankName,TFR.DeductionRemark,MFB.FAGBankID,EMP.BankAcNo as 'VendorACCNo'"
    tmp = tmp & " From Txn_FranchiseeRent TFR"
    tmp = tmp & " Inner Join Mst_Vendor EMP On TFR.VendorID = EMP.VendorID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP1 On TFR.ForwardToID = EMP1.EmployeeID"
    tmp = tmp & " Left Join Mst_FAGBank MFB On TFR.FAGBankID = MFB.FAGBankID"
    If Gen_DBType = "MDB" Then
       tmp = tmp & " Where TFR.BatchID > 0 And TFR.Flag = 'A' And TFR.ClaimDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
    Else
       tmp = tmp & " Where TFR.BatchID > 0 And TFR.Flag = 'A' And TFR.ClaimDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
    End If
    If CmbSVendorID.Text <> "" Then
       tmp = tmp & " AND EMP.VendorID = " & mSVendorID
    End If
    If CmbSForwardTo.Text <> "" Then
       strSplitString = Split(CmbSForwardTo, "~")
       EmpName_ = Trim(strSplitString(0))
       EmpNo_ = Trim(strSplitString(1))
       tmp = tmp & " AND EMP1.EmployeeNo = '" & EmpNo_ & "' "
    End If
    If TxtSClaimID <> "" Then
       tmp = tmp & " AND TFR.ClaimID Like " & TxtSClaimID & " "
    End If
    tmp = tmp & " Order By TFR.ClaimID"
ElseIf OptSecurityBill.Value = True Then
    tmp = "Select TSB.BillTxnID ClaimID,TSB.BillTxnDate ClaimDate,TSB.MonthYear ExpenseFrom,TSB.MonthYear ExpenseTo,TSB.SecurityAgencyID,ME.AgencyName VendorName,TSB.TotalAmt Amount,TSB.Remarks,TSB.Status Flag,TSB.BillNo InvoiceNo,TSB.BillDate InvoiceDate,"
    tmp = tmp & " TSB.G_UID,TSB.FAGRemarks ApprovalRemarks,TSB.ApprovedDate,EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as ForwardedTo,TSB.BatchID,TSB.DisbuDate,TSB.DisbuRemark,TSB.DisbuAmount,TSB.ModeOfPay,MFB.BankName,TSB.DeductionRemark,TSB.FAGBankID,ME.BankAcNo as 'VendorACCNo'"
    tmp = tmp & " From Txn_SecurityBilling TSB"
    tmp = tmp & " Inner Join Mst_SecurityAgency ME On TSB.SecurityAgencyID=ME.SecurityAgencyID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP1 On TSB.ForwardTo = EMP1.EmployeeID"
    tmp = tmp & " Left Join Mst_FAGBank MFB On TSB.FAGBankID = MFB.FAGBankID"
    If Gen_DBType = "MDB" Then
       tmp = tmp & " Where TSB.BatchID > 0 And TSB.Status = 'A' And TSB.BillTxnDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
    Else
       tmp = tmp & " Where TSB.BatchID > 0 And TSB.Status = 'A' And TSB.BillTxnDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
    End If
    If CmbSVendorID.Text <> "" Then
       tmp = tmp & " AND TSB.SecurityAgencyID = " & mSVendorID
    End If
    If CmbSForwardTo.Text <> "" Then
       strSplitString = Split(CmbSForwardTo, "~")
       EmpName_ = Trim(strSplitString(0))
       EmpNo_ = Trim(strSplitString(1))
       tmp = tmp & " AND EMP1.EmployeeNo = '" & EmpNo_ & "' "
    End If
    If TxtSClaimID <> "" Then
       tmp = tmp & " AND TSB.BillTxnID Like " & TxtSClaimID & " "
    End If
    tmp = tmp & " Order By TSB.BillTxnID"
ElseIf OptFumigationBill.Value = True Then
    tmp = "Select TFA.MappingID ClaimID,TFA.MappingDate ClaimDate,TFA.MonthYear ExpenseFrom,TFA.MonthYear ExpenseTo,TFA.VendorID,MV.VendorName,TFA.ActualAmount Amount,TFA.Remark Remarks,TFA.Flag,TFA.BillNo InvoiceNo,TFA.BillDate InvoiceDate,"
    tmp = tmp & " TFA.G_UID,TFA.ApprovalRemark ApprovalRemarks,TFA.ApprovedDate,EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as ForwardedTo,TFA.BatchID,TFA.DisbuDate,TFA.DisbuRemark,TFA.DisbuAmount,TFA.ModeOfPay,MFB.BankName,TFA.DeductionRemark,MFB.FAGBankID,MV.BankAcNo as 'VendorACCNo'"
    tmp = tmp & " From Txn_FumigationApproval TFA"
    tmp = tmp & " Inner Join Mst_Vendor MV On TFA.VendorID = MV.VendorID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP1 On TFA.ForwardTo = EMP1.EmployeeID"
    tmp = tmp & " Left Join Mst_FAGBank MFB On TFA.FAGBankID = MFB.FAGBankID"
    If Gen_DBType = "MDB" Then
       tmp = tmp & " Where TFA.BatchID > 0 And TFA.Flag = 'A' And TFA.MappingDate Between #" & Format(SFrom, Gen_DatFormat) & "# And #" & Format(STo, Gen_DatFormat) & "#) "
    Else
       tmp = tmp & " Where TFA.BatchID > 0 And TFA.Flag = 'A' And TFA.MappingDate Between '" & Format(SFrom, Gen_DatFormat) & "' And '" & Format(STo, Gen_DatFormat) & "' "
    End If
    If CmbSVendorID.Text <> "" Then
       tmp = tmp & " AND TFA.VendorID = " & mSVendorID
    End If
    If CmbSForwardTo.Text <> "" Then
       strSplitString = Split(CmbSForwardTo, "~")
       EmpName_ = Trim(strSplitString(0))
       EmpNo_ = Trim(strSplitString(1))
       tmp = tmp & " AND EMP1.EmployeeNo = '" & EmpNo_ & "' "
    End If
    If TxtSClaimID <> "" Then
       tmp = tmp & " AND TFA.MappingID Like " & TxtSClaimID & " "
    End If
    tmp = tmp & " Order By TFA.MappingID"
End If

Call TmpTable
Call Grid_Head1

x = 2

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 3

For iu = 1 To TmpRs.RecordCount
With MSHFlexGrid1

    .row = x
    .Col = 0
    .CellFontName = "Wingdings"
    .CellFontSize = 14
    .CellAlignment = flexAlignCenterCenter
    .Text = strUnChecked

    .TextMatrix(x, 1) = IIf(IsNull(TmpRs!ClaimID), "", TmpRs!ClaimID) '"Claim ID"
    .TextMatrix(x, 2) = IIf(IsNull(TmpRs!ClaimDate), "", Format(TmpRs!ClaimDate, "DD-MMM-YYYY")) '"Claim Date"
    .TextMatrix(x, 3) = IIf(IsNull(TmpRs!ExpenseFrom), "", Format(TmpRs!ExpenseFrom, "DD-MMM-YYYY")) '"Expense From"
    .TextMatrix(x, 4) = IIf(IsNull(TmpRs!ExpenseTo), "", Format(TmpRs!ExpenseTo, "DD-MMM-YYYY")) '"Expense To"
    .TextMatrix(x, 5) = IIf(IsNull(TmpRs!VendorName), "", TmpRs!VendorName) '"VendorName"
    .TextMatrix(x, 6) = IIf(IsNull(TmpRs!Amount), "", TmpRs!Amount) '"Claim Amount"
    .TextMatrix(x, 7) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks) '"Remarks"
    .TextMatrix(x, 8) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag) '"Status"
    .TextMatrix(x, 9) = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID) '"G_UID"
    .TextMatrix(x, 10) = IIf(IsNull(TmpRs!ApprovalRemarks), "", TmpRs!ApprovalRemarks) '"FAG Remarks"
    .TextMatrix(x, 11) = IIf(IsNull(TmpRs!ApprovedDate), "", Format(TmpRs!ApprovedDate, "DD-MMM-YYYY")) '"Booked in NAV Date"
    .TextMatrix(x, 12) = IIf(IsNull(TmpRs!ForwardedTo), "", TmpRs!ForwardedTo) '"Claim Forwarded To"
    
    If OptFumigationBill.Value = True Or OptSecurityBill.Value = True Then
    .TextMatrix(x, 13) = IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo) '"Invoice No"
    .TextMatrix(x, 14) = IIf(IsNull(TmpRs!InvoiceDate), "", Format(TmpRs!InvoiceDate, "DD-MMM-YYYY")) '"Invoice Date"
    End If
    
    .TextMatrix(x, 15) = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID) '"Batch No"
    .TextMatrix(x, 16) = IIf(IsNull(TmpRs!DisbuDate), "", Format(TmpRs!DisbuDate, "DD-MMM-YYYY")) '"Disbursment Date"
    .TextMatrix(x, 17) = IIf(IsNull(TmpRs!DisbuRemark), "", TmpRs!DisbuRemark) '"Disbursment Remarks"
    .TextMatrix(x, 18) = IIf(IsNull(TmpRs!DisbuAmount), "", TmpRs!DisbuAmount) '"Disbursment Amount"
    .TextMatrix(x, 19) = IIf(IsNull(TmpRs!ModeOfPay), "", TmpRs!ModeOfPay) '"Mode for Payment"
    .TextMatrix(x, 20) = IIf(IsNull(TmpRs!BankName), "", TmpRs!BankName) '"Disbursment through Bank"
    .TextMatrix(x, 21) = IIf(IsNull(TmpRs!DeductionRemark), "", TmpRs!DeductionRemark) '"Deduction Remark"
    .TextMatrix(x, 22) = IIf(IsNull(TmpRs!FAGBankID), "", TmpRs!FAGBankID) '"Deduction Remark"
    .TextMatrix(x, 23) = IIf(IsNull(TmpRs!VendorACCNo), "", TmpRs!VendorACCNo) '"Deduction Remark"

    If LCase(.TextMatrix(x, 8)) = "a" Then .TextMatrix(x, 8) = "Approved"
    If LCase(.TextMatrix(x, 8)) = "p" Then .TextMatrix(x, 8) = "Pending"
    If LCase(.TextMatrix(x, 8)) = "r" Then .TextMatrix(x, 8) = "Rejected"
    If LCase(.TextMatrix(x, 8)) = "f" Then .TextMatrix(x, 8) = "Frozen"
    If LCase(.TextMatrix(x, 8)) = "s" Then .TextMatrix(x, 8) = "Settled"
    .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) + Val(.TextMatrix(x, 6))
    .Rows = .Rows + 1
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
    x = x + 1
End With
Next

ProgressBar1.Value = ProgressBar1.Max
CmdSave.Enabled = True
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp

If OptGodownRent.Value = True Then
    Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_GodownRent.xls")
ElseIf OptFranchiseeFee.Value = True Then
    Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_FranchiseeFee.xls")
ElseIf OptSecurityBill.Value = True Then
    Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_SecurityBill.xls")
ElseIf OptFumigationBill.Value = True Then
    Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_FumigationBill.xls")
End If
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

If OptGodownRent.Value = True Then
    Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_GodownRent.xls")
ElseIf OptFranchiseeFee.Value = True Then
    Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_FranchiseeFee.xls")
ElseIf OptSecurityBill.Value = True Then
    Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_SecurityBill.xls")
ElseIf OptFumigationBill.Value = True Then
    Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_VendorClaimSettleMultiple_FumigationBill.xls")
End If

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

Private Sub MSHFlexGrid1_Click()
Dim SelRow As Variant
SelRow = MSHFlexGrid1.RowSel
If MSHFlexGrid1.TextMatrix(SelRow, 5) = "" Then Exit Sub
With MSHFlexGrid1
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = strUnChecked Then
           .TextMatrix(.row, .Col) = strChecked
           '.TextMatrix(.Row, 0) = strUnChecked
           If Trim(.TextMatrix(.row, 10)) <> "" Then
              .TextMatrix(.row, 17) = .TextMatrix(.row, 10)
           End If
           If Trim(.TextMatrix(.row, 20)) = "" Then
              Call TableMst_FAGBank(" Where FAGBankID in (" & Gen_FAGBankIDForPayment & ") ")
              If RsMst_FAGBank.RecordCount > 0 Then
                 .TextMatrix(.row, 20) = RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
                 .TextMatrix(.row, 22) = RsMst_FAGBank!FAGBankID
              End If
           End If
        Else
           .TextMatrix(.row, .Col) = strUnChecked
           .TextMatrix(.row, 17) = ""
        End If
    End If
End With

If MSHFlexGrid1.TextMatrix(SelRow, 0) = strChecked Then 'Or MSHFlexGrid1.TextMatrix(SelRow, 2) = strChecked Then
   If MSHFlexGrid1.Col = 16 Then
      mSelRow = MSHFlexGrid1.row
      mSelCol = MSHFlexGrid1.Col
      TxtDisbuDate.Visible = True
      TxtDisbuAmount.Visible = False
      TxtDisbuRemarks.Visible = False
      TxtDeductionRemarks.Visible = False
      CmbModeOfPay.Visible = False
      CmbFAGBank.Visible = False
      With TxtDisbuDate
          'TxtDisbuDate.Value = Date
          .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
          .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
          .Width = MSHFlexGrid1.CellWidth - 10
          .Height = MSHFlexGrid1.CellHeight - 10
          If MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = "" Then
             .Value = Date 'IIf(MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = "", Date, CDate(MSHFlexGrid1.TextMatrix(mSelRow, mSelCol)))
          Else
             .Value = CDate(MSHFlexGrid1.TextMatrix(mSelRow, mSelCol))
          End If
          .ZOrder
          .SetFocus
      End With
   End If

   If MSHFlexGrid1.Col = 17 Then
      mSelRow = MSHFlexGrid1.row
      mSelCol = MSHFlexGrid1.Col
      TxtDisbuDate.Visible = False
      TxtDisbuAmount.Visible = False
      TxtDisbuRemarks.Visible = True
      TxtDeductionRemarks.Visible = False
      CmbModeOfPay.Visible = False
      CmbFAGBank.Visible = False
      With TxtDisbuRemarks
          .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
          .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
          .Width = MSHFlexGrid1.CellWidth - 10
          .Height = MSHFlexGrid1.CellHeight - 10
          .Text = MSHFlexGrid1.TextMatrix(mSelRow, mSelCol)
          .ZOrder
          .SetFocus
      End With
   End If


   If MSHFlexGrid1.Col = 18 Then
      mSelRow = MSHFlexGrid1.row
      mSelCol = MSHFlexGrid1.Col
      TxtDisbuDate.Visible = False
      TxtDisbuAmount.Visible = True
      TxtDisbuRemarks.Visible = False
      TxtDeductionRemarks.Visible = False
      CmbModeOfPay.Visible = False
      CmbFAGBank.Visible = False
      With TxtDisbuAmount
          .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
          .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
          .Width = MSHFlexGrid1.CellWidth - 10
          .Height = MSHFlexGrid1.CellHeight - 10
          .Text = MSHFlexGrid1.TextMatrix(mSelRow, mSelCol)
          If .Text = "" Then
             .Text = Val(MSHFlexGrid1.TextMatrix(mSelRow, 6))
          End If
          .ZOrder
          .SetFocus
      End With
   End If


   If MSHFlexGrid1.Col = 19 Then
      mSelRow = MSHFlexGrid1.row
      mSelCol = MSHFlexGrid1.Col
      TxtDisbuDate.Visible = False
      TxtDisbuAmount.Visible = False
      TxtDisbuRemarks.Visible = False
      TxtDeductionRemarks.Visible = False
      CmbModeOfPay.Visible = True
      CmbFAGBank.Visible = False
      With CmbModeOfPay
          .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
          .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
          .Width = MSHFlexGrid1.CellWidth - 10
          '.Height = MSHFlexGrid1.CellHeight - 10
          .Text = MSHFlexGrid1.TextMatrix(mSelRow, mSelCol)
          .ZOrder
          .SetFocus
      End With
   End If


   If MSHFlexGrid1.Col = 20 Then
      mSelRow = MSHFlexGrid1.row
      mSelCol = MSHFlexGrid1.Col
      TxtDisbuDate.Visible = False
      TxtDisbuAmount.Visible = False
      TxtDisbuRemarks.Visible = False
      TxtDeductionRemarks.Visible = False
      CmbModeOfPay.Visible = False
      CmbFAGBank.Visible = True
      With CmbFAGBank
          .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
          .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
          .Width = MSHFlexGrid1.CellWidth - 10
          '.Height = MSHFlexGrid1.CellHeight - 10
          .Text = MSHFlexGrid1.TextMatrix(mSelRow, mSelCol)
          .ZOrder
          .SetFocus
      End With
   End If

   If MSHFlexGrid1.Col = 21 Then
      mSelRow = MSHFlexGrid1.row
      mSelCol = MSHFlexGrid1.Col
      TxtDisbuDate.Visible = False
      TxtDisbuAmount.Visible = False
      TxtDisbuRemarks.Visible = False
      TxtDeductionRemarks.Visible = True
      CmbModeOfPay.Visible = False
      CmbFAGBank.Visible = False
      With TxtDeductionRemarks
          .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
          .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
          .Width = MSHFlexGrid1.CellWidth - 10
          .Height = MSHFlexGrid1.CellHeight - 10
          .Text = MSHFlexGrid1.TextMatrix(mSelRow, mSelCol)
          .ZOrder
          .SetFocus
      End With
   End If
   
   If MSHFlexGrid1.Col <> 21 And MSHFlexGrid1.Col <> 20 And MSHFlexGrid1.Col <> 19 And MSHFlexGrid1.Col <> 18 And MSHFlexGrid1.Col <> 17 And MSHFlexGrid1.Col <> 16 Then
      TxtDisbuDate.Visible = False
      TxtDisbuAmount.Visible = False
      TxtDisbuRemarks.Visible = False
      TxtDeductionRemarks.Visible = False
      CmbModeOfPay.Visible = False
      CmbFAGBank.Visible = False
   End If
   
End If

End Sub

Private Sub CmbModeOfPay_LostFocus()

If CmbModeOfPay.Text = "" Then
   MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = CmbModeOfPay.Text
   CmbModeOfPay.Visible = False
   Exit Sub
End If
If LCase(CmbModeOfPay.Text) <> "rtgs" And LCase(CmbModeOfPay.Text) <> "neft" And LCase(CmbModeOfPay.Text) <> "bank transfer" And LCase(CmbModeOfPay.Text) <> "cheque" Then
   MsgBox "Invalid Selection!!!"
   CmbModeOfPay.SetFocus
   Exit Sub
End If

MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = CmbModeOfPay.Text
CmbModeOfPay.Visible = False

End Sub

Private Sub CmbFAGBank_GotFocus()
tmp = CmbFAGBank.Text
Call TableMst_FAGBank("")
CmbFAGBank.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbFAGBank.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbFAGBank.Text = tmp
End Sub

Private Sub CmbFAGBank_LostFocus()
Dim strSplitString() As String
Dim AcNo_, BankName_ As String

If CmbFAGBank.Text = "" Then
   MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = CmbFAGBank.Text
   MSHFlexGrid1.TextMatrix(mSelRow, 22) = ""
   CmbFAGBank.Visible = False
   Exit Sub
End If
strSplitString = Split(CmbFAGBank, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbFAGBank.SetFocus
   Exit Sub
End If
BankName_ = Trim(strSplitString(0))
AcNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo='" & AcNo_ & "'"
If RsMst_FAGBank.EOF Then
   MsgBox "Invalid selection "
   CmbFAGBank.SetFocus
   Exit Sub
End If
mFAGBankID = RsMst_FAGBank!FAGBankID
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = CmbFAGBank.Text
MSHFlexGrid1.TextMatrix(mSelRow, 22) = mFAGBankID
CmbFAGBank.Visible = False
End Sub

Private Sub MSHFlexGrid1_Scroll()
TxtDisbuDate.Visible = False
TxtDisbuAmount.Visible = False
TxtDisbuRemarks.Visible = False
TxtDeductionRemarks.Visible = False
CmbModeOfPay.Visible = False
CmbFAGBank.Visible = False
End Sub

''--- Clear Grid On every Claim Type
Private Sub OptFranchiseeFee_Click()
    If OptFranchiseeFee.Value = True Then Call CmdClear_Click
End Sub
Private Sub OptFumigationBill_Click()
    If OptFumigationBill.Value = True Then Call CmdClear_Click
End Sub
Private Sub OptGodownRent_Click()
    If OptGodownRent.Value = True Then Call CmdClear_Click
End Sub
Private Sub OptSecurityBill_Click()
    If OptSecurityBill.Value = True Then Call CmdClear_Click
End Sub
''------

Private Sub TxtDeductionRemarks_LostFocus()
If TxtDeductionRemarks.Text = "" Then
   MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Trim(TxtDeductionRemarks.Text)
   TxtDeductionRemarks.Visible = False
   Exit Sub
End If
TxtDeductionRemarks = Replace(TxtDeductionRemarks.Text, vbNewLine, "")
TxtDeductionRemarks = Replace(TxtDeductionRemarks.Text, ",", "~")
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Trim(TxtDeductionRemarks.Text)
TxtDeductionRemarks.Visible = False
End Sub

Private Sub TxtDisbuAmount_LostFocus()
If TxtDisbuAmount = "" Then
   MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Val(TxtDisbuAmount)
   TxtDisbuAmount.Visible = False
   Exit Sub
End If
If IsNumeric(TxtDisbuAmount) = False Or Val(TxtDisbuAmount) < 0 Then
   MsgBox "Invalid format!!!"
   TxtDisbuAmount.SetFocus
   Exit Sub
End If
If Val(TxtDisbuAmount) > Val(MSHFlexGrid1.TextMatrix(mSelRow, 6)) Then
   MsgBox "Disbusment amount should be less than or equal to Claim Amount!!!"
   TxtDisbuAmount.SetFocus
   Exit Sub
End If
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Val(TxtDisbuAmount)
TxtDisbuAmount.Visible = False
End Sub

Private Sub TxtDisbuDate_LostFocus()
If TxtDisbuDate.Value > Date Then
   MsgBox "Future date not allowed!!"
   TxtDisbuDate.SetFocus
   Exit Sub
End If
If TxtDisbuDate.Value < CDate(MSHFlexGrid1.TextMatrix(mSelRow, 2)) Then
   MsgBox "Disbursement Date should not less than Claim Date!!"
   TxtDisbuDate.SetFocus
   Exit Sub
End If
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = TxtDisbuDate.Value
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Format(MSHFlexGrid1.TextMatrix(mSelRow, mSelCol), "dd-mmm-yyyy")
End Sub

Private Sub TxtDisbuRemarks_LostFocus()
If TxtDisbuRemarks.Text = "" Then
   MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Trim(TxtDisbuRemarks.Text)
   TxtDisbuRemarks.Visible = False
   Exit Sub
End If
TxtDisbuRemarks = Replace(TxtDisbuRemarks.Text, vbNewLine, "")
TxtDisbuRemarks = Replace(TxtDisbuRemarks.Text, ",", "~")
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Trim(TxtDisbuRemarks.Text)
TxtDisbuRemarks.Visible = False
End Sub

Private Function ChkGridSelection() As Boolean
ChkGridSelection = False
For I = 2 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       ChkGridSelection = True
       Exit For
    End If
Next
End Function
