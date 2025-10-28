VERSION 5.00
Begin VB.Form FrmMst_UnmarkPrint 
   Caption         =   "Unmark Print Flag"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   12285
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   12000
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
         Height          =   615
         Left            =   2175
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdUnmarkPrint 
         Caption         =   "Unmark Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "You can unmark print flag only for the date next to BOD date for the selected CMP."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4920
         TabIndex        =   29
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.TextBox TxtRemarks 
         Appearance      =   0  'Flat
         Height          =   630
         Left            =   1680
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3480
         Width           =   4410
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox TxtModuleDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1395
      End
      Begin VB.TextBox TxtCurrentDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1395
      End
      Begin VB.TextBox TxtStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1395
      End
      Begin VB.TextBox TxtExchangeType 
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3060
         Width           =   4410
      End
      Begin VB.TextBox TxtModuleNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         Top             =   3060
         Width           =   1275
      End
      Begin VB.ComboBox CmbModuleName 
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
         ItemData        =   "FrmMstUnmarkPrint.frx":0000
         Left            =   1680
         List            =   "FrmMstUnmarkPrint.frx":0013
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   2640
         Width           =   4440
      End
      Begin VB.TextBox TxtEHoPCMPID 
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
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4425
      End
      Begin VB.TextBox TxtCMPAddress 
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
         Height          =   1305
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5565
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4410
      End
      Begin VB.ComboBox CmbCMPID 
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
         Left            =   1665
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   4440
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   4395
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   10125
      End
      Begin VB.Label Label11 
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
         Left            =   195
         TabIndex        =   32
         Top             =   3668
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Hologram"
         Height          =   180
         Left            =   6720
         TabIndex        =   31
         Top             =   3570
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "Exchange"
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
         Left            =   6255
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   180
         Left            =   4200
         TabIndex        =   23
         Top             =   3090
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "No"
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
         Left            =   195
         TabIndex        =   22
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Module Name"
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
         Left            =   195
         TabIndex        =   21
         Top             =   2693
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "CMP Current Date"
         Height          =   180
         Left            =   3240
         TabIndex        =   20
         Top             =   2250
         Width           =   1560
      End
      Begin VB.Label LblStartDate 
         Caption         =   "CMP Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   19
         Top             =   2190
         Width           =   1440
      End
      Begin VB.Label Label8 
         Caption         =   "EHoP CMP ID"
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
         Left            =   195
         TabIndex        =   18
         Top             =   1770
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "CMP"
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
         Left            =   195
         TabIndex        =   16
         Top             =   1253
         Width           =   1695
      End
      Begin VB.Label Label7 
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
         Left            =   6240
         TabIndex        =   15
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "CMP Address"
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
         Left            =   7710
         TabIndex        =   14
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label Label31 
         Caption         =   "SM_Prefix"
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
         Left            =   195
         TabIndex        =   11
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   9
         Top             =   270
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FrmMst_UnmarkPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mWHRNo, mCommodityID, mExchangeTypeID As Variant
Dim mCurrentDate, mModuleDate As Variant

Private Sub CmbModuleName_LostFocus()
If CmbModuleName.Text = "" Then
   Label4.Caption = "No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   TxtModuleDate = ""
   TxtHologramNo = ""
   Exit Sub
End If

If CmbModuleName.Text = "CDTF" Then
   Label4.Caption = "CDTF No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   
ElseIf CmbModuleName.Text = "CITF" Then
   Label4.Caption = "CITF No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""

ElseIf CmbModuleName.Text = "WHR" Then
   Label4.Caption = "WHR No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""

'ElseIf CmbModuleName.Text = "Cash Receipt" Then
'   Label4.Caption = "Cash Receipt No"
'   TxtModuleNo = ""
'   TxtCommodity = ""
'   TxtExchangeType = ""
   
ElseIf CmbModuleName.Text = "CAR" Then
      
   Label4.Caption = "CAR"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   If LCase(Gen_UserRole) <> "power" Then
      MsgBox "Only 'POWER' user can re-move CAR!!!"
      CmbModuleName.SetFocus
      Exit Sub
   End If
Else
   MsgBox "Invalid Selection!!!"
   Label4.Caption = "No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   CmbModuleName.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub ClearData()
CmbLocationID.Text = ""
CmbCMPID.Text = ""
CmbModuleName = ""
Call CmbLocationID_LostFocus
Call CmbCMPID_LostFocus
Call CmbModuleName_LostFocus
TxtRemarks = ""
End Sub
Private Sub CmdUnmarkPrint_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" Then
   MsgBox "Blank CMP not allowed!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

If CmbModuleName.Text = "" Then
   MsgBox "Blank Module Name not allowed!!!"
   CmbModuleName.SetFocus
   Exit Sub
End If

If TxtModuleNo = "" Then
   MsgBox "Blank " & Label4.Caption & " not allowed!!!"
   TxtModuleNo.SetFocus
   Exit Sub
End If

If Trim(TxtRemarks) = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtRemarks.SetFocus
   Exit Sub
End If

If CmbModuleName.Text = "CDTF" Then
   Call CDTF_Unmark
ElseIf CmbModuleName.Text = "CITF" Then
   Call CITF_Unmark
ElseIf CmbModuleName.Text = "Cash Receipt" Then
'   Call CashReceipt_Unmark
ElseIf LCase(CmbModuleName.Text) = "car" Then
   Call CAR_Unmark 'Claim_Unmark
ElseIf CmbModuleName.Text = "WHR" Then
   Call WHR_Unmark
End If

End Sub
Private Sub CashReceipt_Unmark()

'If ChkGSL = False Then
'   TxtModuleNo.SetFocus
'   Exit Sub
'End If

If IsNull(RsTxn_NewCashReceipt!Flag) Or RsTxn_NewCashReceipt!Flag = "" Then
   Call ClearData
   Exit Sub
End If

RsTxn_NewCashReceipt!Flag = Null
RsTxn_NewCashReceipt!HologramNo = ""
RsTxn_NewCashReceipt.Update

Call SaveLog_UnmarkPrint
Call ClearHologramHistory
Call ClearData

End Sub

Private Sub Claim_Unmark()

'If ChkGSL = False Then
'   TxtModuleNo.SetFocus
'   Exit Sub
'End If

If RsTxn_ClaimSubmission!BatchID <> 0 Then
   MsgBox "NAV Data upload executed you can't unflag this Record !!!"
   Call ClearData
   Exit Sub
End If

If RsTxn_ClaimSubmission!Flag = "P" Then
   Call ClearData
   Exit Sub
End If

RsTxn_ClaimSubmission!Flag = "P"
RsTxn_ClaimSubmission!ApprovalRemarks = ""
RsTxn_ClaimSubmission!ApprovalDate = Null
RsTxn_ClaimSubmission.Update

Call SaveLog_UnmarkPrint

Call ClearData
End Sub

Private Sub CAR_Unmark()

RsTxn_WHRDetail!CARNo = Null
RsTxn_WHRDetail.Update

Call SaveLog_UnmarkPrint

Call ClearData
End Sub



Private Sub CDTF_Unmark()
'Dim ans As Variant

If ChkInvoiceFlag = False Then
   TxtModuleNo.SetFocus
   Exit Sub
End If

If ChkGSL = False Then
   'MsgBox "Can't Unmark print Flag!!!"
   TxtModuleNo.SetFocus
   Exit Sub
End If

If ChkNextCDTF = False Then
   'MsgBox "Can't Unmark print Flag!!!"
   TxtModuleNo.SetFocus
   Exit Sub
End If
   
If IsNull(RsTxn_CxTF_Details!Flag) Or RsTxn_CxTF_Details!Flag = "" Then
   Call ClearData
   Exit Sub
End If

RsTxn_CxTF_Details!Flag = Null
RsTxn_CxTF_Details!HologramNo = ""
RsTxn_CxTF_Details.Update
'ans = MsgBox("Do you want to clear history for respective hologram", vbYesNo)
'If ans = vbYes Then
'   Call TableMst_Hologram("Where HologramNo = " & TxtHologranNo & "")
'   If RsMst_Hologram.RecordCount > 0 Then
'      RsMst_Hologram.Delete
'      RsMst_Hologram.Update
'   End If
'End If
Call SaveLog_UnmarkPrint
Call ClearHologramHistory
Call ClearData
End Sub
Private Sub SaveLog_UnmarkPrint()


Call TableLog_UnmarkPrint_History

RsLog_UnmarkPrint_History.AddNew
RsLog_UnmarkPrint_History!SM_Prefix = TxtPreFix
RsLog_UnmarkPrint_History!DocType = Left(CmbModuleName.Text, 10)

If Left(CmbModuleName.Text, 10) = "WHR" Then
   Dim strSplitString() As String
   strSplitString = Split(TxtModuleNo.Text, "-")
   tmp = UBound(strSplitString)
   If tmp <> 1 Then
      MsgBox "Not a valid WHR Number!!!"
      TxtModuleNo.SetFocus
      Exit Sub
   End If
   tmp = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
   RsLog_UnmarkPrint_History!DocNo = tmp
Else
   RsLog_UnmarkPrint_History!DocNo = TxtModuleNo
End If
RsLog_UnmarkPrint_History!UserId = Gen_UserLoginID
RsLog_UnmarkPrint_History!Remarks = TxtRemarks
RsLog_UnmarkPrint_History!HologramNo = TxtHologramNo
RsLog_UnmarkPrint_History.Update


End Sub
Private Function ChkGSL() As Boolean
Dim Retval As Boolean
Retval = True
Call TableMst_GSL(" Where SM_Prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtModuleNo & " ")
'RsMst_GSL.Filter =
If RsMst_GSL.RecordCount > 0 Then
   
   For I = 0 To RsMst_GSL.RecordCount - 1
       If (RsMst_GSL!KRNo <> 0) Or (IsNull(RsMst_GSL!ISINID) = False) Or (RsMst_GSL!ISINID <> 0) Then ' Or IsNull(RsMst_GSL!ISINID) = False) Then
       'If (RsMst_GSL!ISINID <> "" Or RsMst_GSL!ISINID <> 0 Or IsNull(RsMst_GSL!ISINID) = False) Then '(IsNull(RsMst_GSL!KRNo) = False Or RsMst_GSL!KRNo <> 0) Then 'Or Then
          Retval = False
          ChkGSL = Retval
          MsgBox "ISIN/CDC/KR is mapped against one of the lots. You can't unmark !!! "
          Exit Function
       End If
   
       If RsMst_GSL!Status <> "O" Then
          Retval = False
          ChkGSL = Retval
          MsgBox "GSL Flag of One or more lots are not 'O' . You can't unmark !!! "
          Exit Function
       End If

       If RsMst_GSL!CDTFBags <> 0 Or RsMst_GSL!CDTFWeight <> 0 Then
          MsgBox "One of the lots could have been withdrawn. You can't unmark !!! "
          Retval = False
          Exit For
       End If
   
       If RsMst_GSL!WHRBags <> 0 Or RsMst_GSL!WHRWeight <> 0 Then
          MsgBox "WHR genrated for One or more lots. You can't unmark !!! "
          Retval = False
          Exit For
       End If
   
       If RsMst_GSL!PledgeBags <> 0 Or RsMst_GSL!PledgeWeight <> 0 Then
          MsgBox "Pledge created for One or more lots. You can't unmark !!! "
          Retval = False
          Exit For
       End If
   
       If RsMst_GSL!PledgeNo <> "" And IsNull(RsMst_GSL!PledgeNo) <> True Then
          MsgBox "Pledge was created for One or more lots. You can't unmark !!! "
          Retval = False
          Exit For
       End If
       
       If RsMst_GSL!Connected_Mst_GSL_ID <> 0 And IsNull(RsMst_GSL!Connected_Mst_GSL_ID) <> True Then
          MsgBox "Made up Lot is generated against any one lot. Can't Unmark print flag !!! "
          Retval = False
          Exit For
       End If
   Next
   
End If
   

'RsMst_GSL.Filter = ""
ChkGSL = Retval
End Function
Private Function ChkWHR() As Boolean
Dim Retval As Boolean
Retval = True

'tmp = "Select Count(WHRNo) from Txn_WHR_Detail Where SM_Prefix = '" & TxtPreFix & "' " & _
'      " and CommodityID = " & mCommodityID & " and CMPID = " & mCMPID & " And ('" & ModuleDate.Value & "' > StorageFrom and " & _
'      " '" & ModuleDate.Value & "' < StorageTo)"


If Gen_DBType = "MDB" Then
   tmp = "Select Count(WHRNo) from Txn_WHR_Detail Where SM_Prefix = '" & TxtPreFix & "' " & _
         " and CommodityID = " & mCommodityID & " and CMPID = " & mCMPID & " And (#" & Format(mModuleDate, Gen_DatFormat) & "# > StorageFrom and " & _
         " #" & Format(mModuleDate, Gen_DatFormat) & "# < StorageTo)"
Else
   tmp = "Select Count(WHRNo) from Txn_WHR_Detail Where SM_Prefix = '" & TxtPreFix & "' " & _
         " and CommodityID = " & mCommodityID & " and CMPID = " & mCMPID & " And ('" & Format(mModuleDate, Gen_DatFormat) & "' > StorageFrom and " & _
         " '" & Format(mModuleDate, Gen_DatFormat) & "' < StorageTo)"
End If

Call TmpTable

If Not IsNull(TmpRs.Fields(0)) Then
   If TmpRs.Fields(0) > 0 Then
      'MsgBox " WHR can't be generated since all related CDTFs are not printed !!!!! "
      Retval = False
   End If
End If


ChkWHR = Retval
End Function
Private Sub CITF_Unmark()
'Dim ans As Variant

If ChkInvoiceFlag = False Then
   TxtModuleNo.SetFocus
   Exit Sub
End If

RsTxn_CxTF_Details!Flag = ""
RsTxn_CxTF_Details!HologramNo = ""
RsTxn_CxTF_Details.Update

Call SaveLog_UnmarkPrint
Call ClearHologramHistory
Call ClearData
End Sub
Private Function ChkInvoiceFlag() As Boolean

Dim Retval As Boolean
Retval = True

If LCase(CmbModuleName.Text) = "cdtf" Then

   tmp = "Select * from Txn_CxTF_GSL Where SM_Prefix = '" & TxtPreFix & "' And "
   tmp = tmp & " CxTF_DetailsID = " & TxtModuleNo & "  And TxnType = 'D' And BillFlag is not null "
   
   Call TmpTable
Else
   tmp = "Select * from Txn_CxTF_GSL Where SM_Prefix = '" & TxtPreFix & "' And "
   tmp = tmp & " CxTF_DetailsID = " & TxtModuleNo & "  And TxnType = 'W' And BillFlag is not null "

   Call TmpTable

End If

If TmpRs.RecordCount > 0 Then
   MsgBox "Invoice Generated for this Transaction!!!"
   Retval = False
End If

ChkInvoiceFlag = Retval

End Function
Private Sub WHR_Unmark()
'Dim ans As Variant
Call TableTxn_PledgeCreation(" Where SM_Prefix = '" & TxtPreFix & "' And WHRNo = " & mWHRNo & " ")
Call TableTxn_PledgeCreation_GSL
If RsTxn_PledgeCreation.RecordCount > 0 Then
   For I = 0 To RsTxn_PledgeCreation.RecordCount - 1
       RsTxn_PledgeCreation_GSL.MoveFirst
       RsTxn_PledgeCreation_GSL.Find "PledgeID = " & RsTxn_PledgeCreation!PledgeId & ""
       If Not RsTxn_PledgeCreation_GSL.EOF Then
          If Val(RsTxn_PledgeCreation_GSL!PledgeNoOfBags) > 0 Or Val(RsTxn_PledgeCreation_GSL!PledgeQuantity) > 0 Then
             MsgBox "WHR already Pledged can't unmark print Flag!!!"
             TxtModuleNo.SetFocus
             Exit Sub
          End If
       End If
       RsTxn_PledgeCreation.MoveNext
   Next
End If


If Val(RsTxn_WHRDetail!DeliveredNoOfBags) > 0 Or Val(RsTxn_WHRDetail!DeliveredQuantity) > 0 Then
   MsgBox "WHR already Withdrawn can't unmark print flag!!!"
   TxtModuleNo.SetFocus
   Exit Sub
End If

RsTxn_WHRDetail!Flag = Null
RsTxn_WHRDetail!HologramNo = Null
RsTxn_WHRDetail.Update
'ans = MsgBox("Do you want to clear history for respective hologram", vbYesNo)
'If ans = vbYes Then
'   Call TableMst_Hologram("Where HologramNo = " & TxtHologranNo & "")
'   If RsMst_Hologram.RecordCount > 0 Then
'      RsMst_Hologram.Delete
'      RsMst_Hologram.Update
'   End If
'End If
Call SaveLog_UnmarkPrint
Call ClearHologramHistory
Call ClearData

End Sub
Private Sub ClearHologramHistory()
Dim ans As Variant
If TxtHologramNo = "" Then Exit Sub
'ans = MsgBox("Do you want to clear history for respective hologram", vbYesNo)
'If ans = vbYes Then
'   Call TableMst_Hologram("Where HologramNo = '" & TxtHologramNo & "'")
'   If RsMst_Hologram.RecordCount > 0 Then
'      RsMst_Hologram.Delete
'      RsMst_Hologram.Update
'   End If
'End If


MsgBox "Hologram No can't be reused!!!"

End Sub
Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

Call TableTxn_CxTF_Details
Call TableTxn_WHRDetail
Call ClearData

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
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
   TxtPreFix = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   CmbCMPID.Text = ""
   TxtPreFix = ""
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbCMPID_GotFocus()
    tmp = CmbCMPID.Text
    Call TableMst_CMP(" where CloseDate is null and LocationID = " & mLocationID & " ")
    CmbCMPID.Clear
    If RsMst_CMP.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    For I = 1 To RsMst_CMP.RecordCount
        CmbCMPID.AddItem RsMst_CMP!CMPName
        RsMst_CMP.MoveNext
    Next
    CmbCMPID.Text = tmp

End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   mCMPID = Null
   TxtCMPAddress = ""
   TxtCommodity = ""
   TxtEHoPCMPID = ""
   TxStartDate = ""
   TxtCurrentDate = ""
   CmbModuleName.Text = ""
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   TxtCommodity = ""
   TxtStartDate = ""
   TxtCurrentDate = ""
   CmbModuleName.Text = ""
   TxtEHoPCMPID = ""
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
End If


mCMPID = RsMst_CMP!CMPID
TxtCMPAddress = IIf(IsNull(RsMst_CMP!Address), "", RsMst_CMP!Address)
TxtStartDate = RsMst_CMP!StartDate 'Format(RsMst_CMP!StartDate, "dd/mm/yyyy")
TxtCurrentDate = RsMst_CMP!CurrentDate + 1 ' CDate(Format(RsMst_CMP!CurrentDate, "dd/mm/yyyy")) + 1
TxtEHoPCMPID = RsMst_CMP!EHOPCMPID


'TxtCxTFDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
'TxtCxTFDispalyDate = TxtCxTFDate.Value
'TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub
Private Sub TxtModuleNo_LostFocus()
Dim ans As Variant
If CmbLocationID.Text = "" Then Exit Sub
If CmbCMPID.Text = "" Then Exit Sub
If CmbModuleName.Text = "" Then Exit Sub
If TxtModuleNo = "" Then
   RsTxn_CxTF_Details.Filter = ""
   RsTxn_WHRDetail.Filter = ""
   Exit Sub
End If

RsTxn_CxTF_Details.MoveFirst
If RsTxn_WHRDetail.RecordCount > 0 Then
   RsTxn_WHRDetail.MoveFirst
End If
tmp = "SM_Prefix = '" & TxtPreFix & "' And CMPID = " & mCMPID & " And "
If CmbModuleName.Text = "CDTF" Then
   mCurrentDate = Format(TxtCurrentDate, Gen_DatFormat)
   'mCurrentDate = CDate(mCurrentDate) + 1
   
   'Previous code
   'tmp = tmp & "TxnType = 'D' And CxTFDate = '" & mCurrentDate & "' And CxTFNo = " & Val(TxtModuleNo) & ""
   tmp = tmp & "TxnType = 'D'  And CxTF_DetailsID = " & Val(TxtModuleNo) & " " 'And ExchangeTypeID <> 1 " 'And CxTFDate = '" & mCurrentDate & "'
   
   RsTxn_CxTF_Details.Filter = tmp
   
   If RsTxn_CxTF_Details.RecordCount > 0 Then
      TxtModuleDate = RsTxn_CxTF_Details!CxTFDate
      If Format(RsTxn_CxTF_Details!CxTFDate, Gen_DatFormat) <> mCurrentDate Then
         ans = MsgBox("BOD executed. Still you want to continue!!!", vbYesNo)
         If ans = vbNo Then
            TxtModuleNo.SetFocus
            Exit Sub
         End If
      End If
   End If
   
ElseIf CmbModuleName.Text = "CITF" Then
   'mCurrentDate = CDate(TxtCurrentDate) + 1 'CurrentDate.Value + 1
   mCurrentDate = Format(TxtCurrentDate, Gen_DatFormat)
   
   'Previous code
   'tmp = tmp & "TxnType = 'W' And CxTFDate = '" & mCurrentDate & "' And CxTFNo = " & Val(TxtModuleNo) & ""
   
   tmp = tmp & "TxnType = 'W'  And CxTF_DetailsID = " & Val(TxtModuleNo) & "" 'And CxTFDate = '" & mCurrentDate & "'
   
   
   
   RsTxn_CxTF_Details.Filter = tmp

   If RsTxn_CxTF_Details.RecordCount > 0 Then
      TxtModuleDate = RsTxn_CxTF_Details!CxTFDate
      If Format(RsTxn_CxTF_Details!CxTFDate, Gen_DatFormat) <> mCurrentDate Then
         ans = MsgBox("BOD executed. Still you want to continue!!!", vbYesNo)
         If ans = vbNo Then
            TxtModuleNo.SetFocus
            Exit Sub
         End If
      End If
   End If
   
ElseIf CmbModuleName.Text = "WHR" Then
   strSplitString = Split(TxtModuleNo.Text, "-")
   tmp1 = UBound(strSplitString)
   If tmp1 <> 1 Then
      MsgBox "Invalid WHR No Selection"
      TxtModuleNo.SetFocus
      Exit Sub
   End If
   mWHRNo = (Val(strSplitString(0)) * 1000000) + Val(strSplitString(1))
   RsTxn_WHRDetail.Find "WHRNo = " & mWHRNo & ""
   tmp = tmp & " WHRNo = " & mWHRNo & ""
   RsTxn_WHRDetail.MoveFirst
   
   RsTxn_WHRDetail.Filter = tmp
   
End If

If CmbModuleName.Text = "CDTF" Or CmbModuleName.Text = "CITF" Then
   If RsTxn_CxTF_Details.RecordCount > 0 Then
      mCommodityID = RsTxn_CxTF_Details!CommodityID
      TxtCommodity = GetCommodityName(RsTxn_CxTF_Details!CommodityID)
      mExchangeTypeID = RsTxn_CxTF_Details!ExchangeTypeID
      TxtExchangeType = GetExchangeName(RsTxn_CxTF_Details!ExchangeTypeID)
      mModuleDate = RsTxn_CxTF_Details!CxTFDate
      
      TxtHologramNo = IIf(IsNull(RsTxn_CxTF_Details!HologramNo), "", RsTxn_CxTF_Details!HologramNo)
   Else
      MsgBox "No Record found!!!"
      TxtModuleNo.SetFocus
      TxtCommodity = ""
      TxtExchangeType = ""
      RsTxn_CxTF_Details.Filter = ""
      TxtHologramNo = ""
      Exit Sub
   End If
ElseIf CmbModuleName.Text = "WHR" Then
   If RsTxn_WHRDetail.RecordCount > 0 Then
      mCommodityID = RsTxn_WHRDetail!CommodityID
      mExchangeTypeID = 1
      TxtCommodity = GetCommodityName(RsTxn_WHRDetail!CommodityID)
      TxtExchangeType = GetExchangeName(1)
      mModuleDate = RsTxn_WHRDetail!WHRDate
      TxtModuleDate = RsTxn_WHRDetail!WHRDate
      TxtHologramNo = IIf(IsNull(RsTxn_WHRDetail!HologramNo), "", RsTxn_WHRDetail!HologramNo)
   Else
      MsgBox "No Record found!!!"
      TxtModuleNo.SetFocus
      TxtCommodity = ""
      TxtExchangeType = ""
      TxtHologramNo = ""
      RsTxn_WHRDetail.Filter = ""
   End If

ElseIf CmbModuleName.Text = "Cash Receipt" Then

    Call TableTxn_NewCashReceipt(" where CRID = " & Val(TxtModuleNo) & " And  sm_prefix = '" & TxtPreFix & "'  Order by SM_Prefix,CRID")
    
    If RsTxn_NewCashReceipt.RecordCount > 0 Then
       TxtCommodity = RsTxn_NewCashReceipt!Remark
       TxtHologramNo = IIf(IsNull(RsTxn_NewCashReceipt!HologramNo), "", RsTxn_NewCashReceipt!HologramNo)
    Else
       MsgBox "No Record found!!!"
       TxtModuleNo.SetFocus
       TxtCommodity = ""
       TxtExchangeType = ""
       TxtHologramNo = ""
    End If

ElseIf CmbModuleName.Text = "CAR" Then
    'Call TableTxn_ClaimSubmission(" where ClaimRowID = " & Val(TxtModuleNo) & " Order by ClaimRowID ")
    If Val(TxtModuleNo) = 0 Then Exit Sub
    Call TableTxn_WHRDetail(" Where CARNo = " & Val(TxtModuleNo) & " ")
    If RsTxn_WHRDetail.RecordCount > 0 Then
       tmp = "Select * from Txn_CAR Where CARNO = " & Val(TxtModuleNo) & ""
       Call TmpTableCMPro
       If TmpRsCMPro.RecordCount > 0 Then
          MsgBox "CAR not deleted from CM Pro. Please check!!!"
       End If
    Else
       MsgBox "No Record found!!!"
       TxtModuleNo.SetFocus
       TxtCommodity = ""
      TxtExchangeType = ""
      TxtHologramNo = ""
    End If
End If

End Sub

Private Function ChkNextCDTF() As Boolean
Dim Retval As Boolean
Retval = True
tmp = "Select Distinct TCD.CxTFDate,TCD.CxTF_DetailsID,TCD.Flag  From Txn_CxTF_Details TCD"
tmp = tmp & " Inner Join Txn_CxTF_GSL TCG On TCD.CxTF_DetailsID = TCG.CxTF_DetailsID And TCD.SM_Prefix = TCG.SM_Prefix And TCD.TxnType = TCG.TxnType "
tmp = tmp & " Where TCD.TxnType = 'D' And TCD.SM_Prefix='" & TxtPreFix & "' And TCD.CxTF_DetailsID > " & Val(TxtModuleNo) & " And TCG.Mst_GSL_ID in "
tmp = tmp & " (Select Mst_GSL_ID From Txn_CxTF_GSL TCG Where TCG.TxnType = 'D' And TCG.SM_Prefix='" & TxtPreFix & "' And TCG.CxTF_DetailsID=" & Val(TxtModuleNo) & ") And TCD.Flag='P'"
tmp = tmp & " order by TCD.CxTF_DetailsID Desc"

Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = False
   ChkNextCDTF = Retval
   MsgBox "Further CDTF is Printed . You can't unmark !!! "
   Exit Function
End If
ChkNextCDTF = Retval
End Function

