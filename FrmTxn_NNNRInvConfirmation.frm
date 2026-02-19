VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_NNNRInvConfirmation 
   Caption         =   "NNNR / 7 Days Invoice Confirmation"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   9525
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   14970
      Begin VB.TextBox TxtRemark 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   720
         MaxLength       =   100
         TabIndex        =   24
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
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
         Left            =   13230
         TabIndex        =   23
         Top             =   1020
         Width           =   1125
      End
      Begin VB.CheckBox ChkMismatch 
         Appearance      =   0  'Flat
         Caption         =   "Mismatch"
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
         Top             =   960
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.ComboBox CmbInvoiceType 
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
         ItemData        =   "FrmTxn_NNNRInvConfirmation.frx":0000
         Left            =   3000
         List            =   "FrmTxn_NNNRInvConfirmation.frx":0010
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   1920
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   9120
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
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
         Left            =   1560
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtSInvoiceNo 
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
         Left            =   13335
         TabIndex        =   5
         Top             =   420
         Width           =   1305
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
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
         Left            =   10980
         TabIndex        =   7
         Top             =   1020
         Width           =   1125
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
         Left            =   12105
         TabIndex        =   8
         Top             =   1020
         Width           =   1125
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
         Left            =   4950
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   2520
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
         Left            =   7470
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   420
         Width           =   5745
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   90
         TabIndex        =   10
         Top             =   1845
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12726
         _Version        =   393216
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
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   132382721
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1575
         TabIndex        =   1
         Top             =   420
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   132382721
         CurrentDate     =   39884
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
         Left            =   120
         TabIndex        =   22
         Top             =   1425
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Invoices in red below can be approved through Transaction --> Invoice Transaction --> NNNR --> NNNR Invoice Confirmation screen"
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
         Left            =   765
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   13935
      End
      Begin VB.Label Label1 
         Caption         =   "Invoice Type"
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
         Left            =   3360
         TabIndex        =   20
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label14 
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
         Height          =   225
         Left            =   1800
         TabIndex        =   17
         Top             =   165
         Width           =   870
      End
      Begin VB.Label Label15 
         Caption         =   "Invoice No"
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
         Left            =   13680
         TabIndex        =   16
         Top             =   165
         Width           =   840
      End
      Begin VB.Label Label34 
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
         Height          =   225
         Left            =   5820
         TabIndex        =   15
         Top             =   165
         Width           =   810
      End
      Begin VB.Label Label16 
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
         Left            =   10020
         TabIndex        =   14
         Top             =   165
         Width           =   1245
      End
      Begin VB.Label Label23 
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
         Height          =   225
         Left            =   285
         TabIndex        =   13
         Top             =   165
         Width           =   990
      End
      Begin VB.Label LblTotalRecord 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5760
         TabIndex        =   12
         Top             =   975
         Width           =   75
      End
      Begin VB.Label Label24 
         Caption         =   "Total Record Found :-"
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
         Left            =   3030
         TabIndex        =   11
         Top             =   930
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmTxn_NNNRInvConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID As Double
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mMonth, mTotalDays, mYear As Variant
Dim mInvoiceNo As Variant
Dim mFlag As Variant
Dim mSelRow As Variant
Dim mSelCol As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mNoNavcode As Variant

Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

Private Sub ChkCheckAll_Click()
If MSHFlexGrid1.Rows <= 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If (MSHFlexGrid1.CellBackColor) <> vbRed Then
          MSHFlexGrid1.TextMatrix(I, 0) = strChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbInvoiceType_LostFocus()

Call Grid_Head

Label6.Visible = False
Label8.Visible = False
ChkMismatch.Value = 0
ChkMismatch.Visible = False
CmdDelete.Visible = False

If CmbInvoiceType.Text = "" Then Exit Sub

If LCase(CmbInvoiceType.Text) <> "nnnr" And LCase(CmbInvoiceType.Text) <> "7 days" And LCase(CmbInvoiceType.Text) <> "reservation" And LCase(CmbInvoiceType.Text) <> "adhoc" Then
   MsgBox "Invalid Selection!!!"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

If LCase(CmbInvoiceType.Text) = "nnnr" Then
   Label6.Visible = True
   Label8.Visible = True
   ChkMismatch.Visible = True
   If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "headcmg" Then
      CmdDelete.Visible = True
   End If
End If

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

End Sub

Private Sub CmdDelete_Click()

If ChkMismatch.Value = 0 Then Exit Sub
ProgressBar1.Value = 0

If CheckSelect = False Then
   MsgBox "Please select atlease one Invoice!!"
   Exit Sub
End If

If CheckRemarks = False Then
   MsgBox "Please provide remarks for all invoices which are to be deleted!!!"
   Exit Sub
End If

CmdDelete.Enabled = False

Call TableTxn_ProInvoice(" Where Flag='P'")
Call TableLog_InvoiceDeletetion(" Where DocumentType=''")
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If LCase(CmbInvoiceType.Text) = "nnnr" Then
          If RsTxn_ProInvoice.RecordCount > 0 Then RsTxn_ProInvoice.MoveFirst
          RsTxn_ProInvoice.Find "InvoiceNo= " & Val(Replace(MSHFlexGrid1.TextMatrix(I, 1), "NN", ""))
          If RsTxn_ProInvoice.EOF = False Then
             RsLog_InvoiceDeletetion.AddNew
             RsLog_InvoiceDeletetion!DocumentType = CmbInvoiceType.Text
             RsLog_InvoiceDeletetion!DcoumentID = RsTxn_ProInvoice!InvoiceNo
             RsLog_InvoiceDeletetion!SM_Prefix = RsTxn_ProInvoice!SM_Prefix
             RsLog_InvoiceDeletetion!Client = RsTxn_ProInvoice!ClientID
             RsLog_InvoiceDeletetion!Amount = RsTxn_ProInvoice!Amount
             RsLog_InvoiceDeletetion!DeletetionRemark = MSHFlexGrid1.TextMatrix(I, 15)
             RsLog_InvoiceDeletetion!DeletedBy = Gen_UserLoginID
             RsLog_InvoiceDeletetion!DeletedOn = Now
             RsLog_InvoiceDeletetion.Update
             
             Call DeleteChildTableDetail(Val(Replace(MSHFlexGrid1.TextMatrix(I, 1), "NN", "")))
             
             RsTxn_ProInvoice.Delete
             RsTxn_ProInvoice.Update
          End If
       End If
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
MsgBox "Deleted Succesfully!!"
ProgressBar1.Value = ProgressBar1.Max
Call CmdSearch_Click

End Sub
Private Sub DeleteChildTableDetail(x_ As Variant)

Dim mSMonth, mSYear, mSFromDate As Variant

mSMonth = Month(CDate(TxtInvoiceDate))
mSYear = Year(CDate(TxtInvoiceDate))

mSFromDate = Format(mSYear & "-" & mSMonth & "-" & "01", "yyyy-mm-dd")

tmp = "Select * From Txn_InvMonGSLDetail Where PINo = " & x_ & " "
Call TmpTable

For j = 1 To TmpRs.RecordCount
    TmpRs!PINo = Null
    TmpRs!InvoiceDate = Null
    TmpRs.Update
    TmpRs.MoveNext
Next

tmp = "Delete From Txn_MonthwiseInvoiceData Where PINo = " & x_ & " "
Call TmpTable


End Sub

Private Sub CmdSave_Click()

If ChkMismatch.Value = 1 Then Exit Sub
ProgressBar1.Value = 0

If CheckSelect = False Then
   MsgBox "Please select atlease one Invoice!!"
   Exit Sub
End If

CmdSave.Enabled = False

If LCase(CmbInvoiceType.Text) = "nnnr" Then
   Call TableTxn_ProInvoice(" where Flag='P'")
ElseIf LCase(CmbInvoiceType.Text) = "reservation" Then
   Call TableTxn_ReservationInvoice(" where Flag='P'")
ElseIf LCase(CmbInvoiceType.Text) = "adhoc" Then
   Call TableTxn_TempInvoice(" where Flag='P'")

Else
   Call TableTxn_CDTFBOE(" where Flag='P'")
End If

ProgressBar1.Max = MSHFlexGrid1.Rows + 1
mNoNavcode = 0
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If LCase(CmbInvoiceType.Text) = "nnnr" Then
          If RsTxn_ProInvoice.RecordCount > 0 Then RsTxn_ProInvoice.MoveFirst
          RsTxn_ProInvoice.Find "InvoiceNo= " & Val(Replace(MSHFlexGrid1.TextMatrix(I, 1), "NN", ""))
        
          If RsTxn_ProInvoice.EOF = False Then
             If GetStateWiseNAVCode(RsTxn_ProInvoice!ClientID, RsTxn_ProInvoice!LocationID, "N") <> "" Then
                RsTxn_ProInvoice!Flag = "C"
                RsTxn_ProInvoice!UpdatedDate = Now
                RsTxn_ProInvoice.Update
              Else
                For C = 0 To MSHFlexGrid1.Cols - 1
                    MSHFlexGrid1.Col = C
                    MSHFlexGrid1.row = I
                    MSHFlexGrid1.CellBackColor = vbRed
                Next
                mNoNavcode = mNoNavcode + 1
              End If
          End If
       ElseIf LCase(CmbInvoiceType.Text) = "reservation" Then
          If RsTxn_ReservationInvoice.RecordCount > 0 Then RsTxn_ReservationInvoice.MoveFirst
          RsTxn_ReservationInvoice.Find "InvoiceNo= " & Val(Replace(MSHFlexGrid1.TextMatrix(I, 1), "RI", ""))
          If RsTxn_ReservationInvoice.EOF = False Then
             If GetStateWiseNAVCode(GetClientfromContract(RsTxn_ReservationInvoice!ContractID), RsTxn_ReservationInvoice!LocationID, "R") <> "" Then
                RsTxn_ReservationInvoice!Flag = "A"
                RsTxn_ReservationInvoice!UpdatedDate = Now
                RsTxn_ReservationInvoice.Update
              Else
                For C = 0 To MSHFlexGrid1.Cols - 1
                    MSHFlexGrid1.Col = C
                    MSHFlexGrid1.row = I
                    MSHFlexGrid1.CellBackColor = vbRed
                Next
                mNoNavcode = mNoNavcode + 1
              End If
          End If
       ElseIf LCase(CmbInvoiceType.Text) = "7 days" Then
          If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
          RsTxn_CDTFBOE.Find "BOEID= " & Val(Replace(MSHFlexGrid1.TextMatrix(I, 1), "SD", ""))
          If RsTxn_CDTFBOE.EOF = False Then
             If GetStateWiseNAVCode(RsTxn_CDTFBOE!ClientID, RsTxn_CDTFBOE!LocationID, "N") <> "" Then
                RsTxn_CDTFBOE!Flag = "C"
                RsTxn_CDTFBOE!UpdatedDate = Now
                RsTxn_CDTFBOE.Update
              Else
                For C = 0 To MSHFlexGrid1.Cols - 1
                    MSHFlexGrid1.Col = C
                    MSHFlexGrid1.row = I
                    MSHFlexGrid1.CellBackColor = vbRed
                Next
                mNoNavcode = mNoNavcode + 1
              End If
          End If
       ElseIf LCase(CmbInvoiceType.Text) = "adhoc" Then
          If RsTxn_TempInvoice.RecordCount > 0 Then RsTxn_TempInvoice.MoveFirst
          RsTxn_TempInvoice.Find "PINo= " & Val(Replace(MSHFlexGrid1.TextMatrix(I, 1), "AH", ""))
          If RsTxn_TempInvoice.EOF = False Then
              
             If IsNull(RsTxn_TempInvoice!ClientID) = False Then
                If RsTxn_TempInvoice!IETypeID = 4 Then
                    If GetStateWiseNAVCode(RsTxn_TempInvoice!ClientID, RsTxn_TempInvoice!LocationID, "R") <> "" Then
                       RsTxn_TempInvoice!Flag = "A"
                       RsTxn_TempInvoice!FAGRemarks = RsTxn_TempInvoice!Remarks
                       RsTxn_TempInvoice!FAGPostingDate = Date
                       RsTxn_TempInvoice!UpdatedDate = Now
                       RsTxn_TempInvoice.Update
                     Else
                       For C = 0 To MSHFlexGrid1.Cols - 1
                           MSHFlexGrid1.Col = C
                           MSHFlexGrid1.row = I
                           MSHFlexGrid1.CellBackColor = vbRed
                       Next
                       mNoNavcode = mNoNavcode + 1
                     End If
                 Else
                    If GetStateWiseNAVCode(RsTxn_TempInvoice!ClientID, RsTxn_TempInvoice!LocationID, "N") <> "" Then
                       RsTxn_TempInvoice!Flag = "A"
                       RsTxn_TempInvoice!FAGRemarks = RsTxn_TempInvoice!Remarks
                       RsTxn_TempInvoice!FAGPostingDate = Date
                       RsTxn_TempInvoice!UpdatedDate = Now
                       RsTxn_TempInvoice.Update
                     Else
                       For C = 0 To MSHFlexGrid1.Cols - 1
                           MSHFlexGrid1.Col = C
                           MSHFlexGrid1.row = I
                           MSHFlexGrid1.CellBackColor = vbRed
                       Next
                       mNoNavcode = mNoNavcode + 1
                     End If
                 
                 End If
              Else
                 RsTxn_TempInvoice!Flag = "A"
                 RsTxn_TempInvoice!FAGRemarks = RsTxn_TempInvoice!Remarks
                 RsTxn_TempInvoice!FAGPostingDate = Date
                 RsTxn_TempInvoice!UpdatedDate = Now
                 RsTxn_TempInvoice.Update
              End If
          End If
       End If
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
If mNoNavcode = 0 Then
   MsgBox "Save Succesfully!!"
Else
   MsgBox " " & mNoNavcode & " number of Invoices are marked in red whose Client Navcode mapping in pending. Please check, other then same are saved succesfully!!"
End If
ProgressBar1.Value = ProgressBar1.Max
Call CmdSearch_Click
End Sub

Private Sub CmdSearch_Click()

ProgressBar1.Value = 0
Call Grid_Head
mNoNavcode = 0
If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date Should Be Less Than To Date !!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFromDate.Value, "yyyy-mm") Then
   MsgBox "From Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtToDate.Value, "yyyy-mm") Then
   MsgBox "To Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtToDate.SetFocus
   Exit Sub
End If

If LCase(CmbInvoiceType.Text) = "nnnr" Then
   
   
   tmp = " Select TPI.InvoiceNo,TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,ME.ExchangeType,MC.ClientName, " & _
         " TPI.Amount , TPI.Discount, TPI.Premium, TPI.BatchID, TPI.ServiceTaxNAVCode,TPI.G_UID,TPI.Remarks,TPI.Flag,TPI.LocationID,TPI.ClientID " & _
         " From Txn_ProInvoice TPI " & _
         " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID " & _
         " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID " & _
         " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID "
    
   If Gen_DBType = "MDB" Then
      wc = " Where TPI.InvoiceDate >= #" & TxtFromDate.Value & "# And TPI.InvoiceDate <= #" & TxtToDate.Value & "# "
   Else
      wc = " Where TPI.InvoiceDate >= '" & TxtFromDate.Value & "' And TPI.InvoiceDate <= '" & TxtToDate.Value & "' "
   End If
               
   If CmbPreFix.Text <> "" Then
      wc = wc & " And  ML.LocationName = '" & CmbPreFix.Text & "'"
   Else
      If Gen_WcLocation <> "" Then
         wc = wc & " And ML." & Gen_WcLocation
      End If
   End If
          
   If CmbNonNcdexClient.Text <> "" Then
      wc = wc & " And  MC.ClientName = '" & CmbNonNcdexClient.Text & "'"
   End If
   
   wc = wc & " And  TPI.Flag = 'P'"
   
   If TxtSInvoiceNo.Text <> "" Then
      wc = wc & " And  TPI.InvoiceNo = '" & Replace(TxtSInvoiceNo.Text, "NN", "") & "'"
   End If
   
   wc = wc & " And  TPI.BatchID = 0 " '
'   Wc = Wc & " And  TPI.InvoiceNo not in (Select Distinct PINo from  Txn_MonthwiseInvoiceData)  " '
                
   tmp1 = " Order By ML.LocationName,TPI.InvoiceNo "
    
   tmp = tmp & " " & wc & " " & tmp1

ElseIf LCase(CmbInvoiceType.Text) = "reservation" Then

   tmp = "Select TPI.InvoiceNo,TPI.YearMonth 'InvoiceDate',TPI.BillingDate 'BillingMonth',ML.LocationName,1 'ExchangeType',MC.ContractName 'ClientName',"
   'tmp = tmp & " TPI.ActualAmount 'Amount' , 0 'Discount', 0 'Premium', TPI.BatchID, TPI.ServiceTaxNAVCode,TPI.G_UID,TPI.Remarks,TPI.Flag,TPI.LocationID,MCl.ClientName 'Client'"
   tmp = tmp & " TPI.ActualAmount 'Amount' , 0 'Discount', 0 'Premium', TPI.BatchID, TPI.ServiceTaxNAVCode,TPI.G_UID,TPI.Remarks,TPI.Flag,TPI.ContractID , TPI.LocationID, MCl.ClientName 'Client',MC.ClientID" 'Query changed - Condition for reservation added on 31-01-2025
   tmp = tmp & " From Txn_ReservationInvoice TPI"
   tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_StorageContract MC On TPI.ContractID = MC.ContractID"
   tmp = tmp & " Inner Join Mst_Client MCl On MC.ClientID = MCl.ClientID"

   If Gen_DBType = "MDB" Then
      wc = " Where TPI.BillingDate >= #" & TxtFromDate.Value & "# And TPI.BillingDate <= #" & TxtToDate.Value & "# "
   Else
      wc = " Where TPI.BillingDate >= '" & TxtFromDate.Value & "' And TPI.BillingDate <= '" & TxtToDate.Value & "' "
   End If
               
   If CmbPreFix.Text <> "" Then
      wc = wc & " And  ML.LocationName = '" & CmbPreFix.Text & "'"
   Else
      If Gen_WcLocation <> "" Then
         wc = wc & " And ML." & Gen_WcLocation
      End If
   End If
         
   If CmbNonNcdexClient.Text <> "" Then
      wc = wc & " And  MCl.ClientName = '" & CmbNonNcdexClient.Text & "'"
   End If
   
   wc = wc & " And  TPI.Flag = 'P'"
   
   If TxtSInvoiceNo.Text <> "" Then
      wc = wc & " And  TPI.InvoiceNo = '" & Replace(TxtSInvoiceNo.Text, "RI", "") & "'"
   End If
   
   wc = wc & " And  TPI.BatchID = 0 " '
               
   tmp1 = " Order By ML.LocationName,TPI.InvoiceNo "
   
   tmp = tmp & " " & wc & " " & tmp1

ElseIf LCase(CmbInvoiceType.Text) = "adhoc" Then


'   tmp = "Select TPI.InvoiceNo,TPI.YearMonth 'InvoiceDate',TPI.BillingDate 'BillingMonth',ML.LocationName,1 'ExchangeType',MC.ContractName 'ClientName',"
'   tmp = tmp & " TPI.ActualAmount 'Amount' , 0 'Discount', 0 'Premium', TPI.BatchID, TPI.ServiceTaxNAVCode,TPI.G_UID,TPI.Remarks,TPI.Flag,TPI.LocationID,MCl.ClientName 'Client'"
'   tmp = tmp & " From Txn_TempInvoice TPI"
'   tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
'   'tmp = tmp & " Inner Join Mst_StorageContract MC On TPI.ContractID = MC.ContractID"
'   tmp = tmp & " Left Join Mst_Client MCl On MC.ClientID = MCl.ClientID"
'   tmp = tmp & " Left Join Mst_Depository MD On TPI.DepositoryID = MD.DepositoryID"
   
   
    tmp = " Select TPI.PINo,TPI.PIDate 'InvoiceDate',TPI.InvoiceDate 'BillingMonth',ML.LocationName,ET.ExchangeType 'ExchangeType',MCl.ClientName 'ClientName',"
    tmp = tmp & " TPI.Amount 'Amount' , 0 'Discount', 0 'Premium', TPI.BatchID, ST.NavCode,TPI.G_UID,TPI.Remarks,TPI.Flag,TPI.LocationID,MCl.ClientName 'Client'"
    tmp = tmp & " From Txn_TempInvoice TPI"
    tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
    tmp = tmp & " Left Join Mst_Client MCl On TPI.ClientID = MCl.ClientID"
    tmp = tmp & " Left Join Mst_Depository MD On TPI.DPID = MD.DepositoryID"
    tmp = tmp & " Inner Join Mst_ServiceTax ST on TPI.ServiceTaxID = ST.ServiceTaxID"
    tmp = tmp & " Inner Join Mst_ExchangeType ET on TPI.ExchangeTypeID = ET.ExchangeTypeID"
    
   
   
   If Gen_DBType = "MDB" Then
      wc = " Where TPI.PIDate >= #" & TxtFromDate.Value & "# And TPI.PIDate <= #" & TxtToDate.Value & "# "
   Else
      wc = " Where TPI.PIDate >= '" & TxtFromDate.Value & "' And TPI.PIDate <= '" & TxtToDate.Value & "' "
   End If
               
   If CmbPreFix.Text <> "" Then
      wc = wc & " And  ML.LocationName = '" & CmbPreFix.Text & "'"
   Else
      If Gen_WcLocation <> "" Then
         wc = wc & " And ML." & Gen_WcLocation
      End If
   End If
         
   If CmbNonNcdexClient.Text <> "" Then
      wc = wc & " And  MCl.ClientName = '" & CmbNonNcdexClient.Text & "'"
   End If
   
   wc = wc & " And  TPI.Flag = 'P'"
   
   If TxtSInvoiceNo.Text <> "" Then
      wc = wc & " And  TPI.PINo = '" & Replace(TxtSInvoiceNo.Text, "AH", "") & "'"
   End If
   
   wc = wc & " And  TPI.BatchID = 0 " '
               
   tmp1 = " Order By ML.LocationName,TPI.PINo "
   
   tmp = tmp & " " & wc & " " & tmp1



Else
   
   tmp = " Select TPI.BOEID,TPI.InvoiceDate,TPI.BillingMonth,ML.LocationName,ME.ExchangeType,MC.ClientName, " & _
         " TPI.Amount , TPI.Discount, TPI.Premium, TPI.BatchID, MS.NAVCode,TPI.G_UID,TPI.Remarks,TPI.Flag " & _
         " From Txn_CDTFBOE TPI " & _
         " Inner Join Txn_CxTF_Details TC On TPI.CDTFNo = TC.CxTFNo And TPI.SM_Prefix=TC.SM_Prefix and TC.TxnType='D' " & _
         " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID " & _
         " Inner Join Mst_Commodity MCom On TC.CommodityID = MCom.CommodityID " & _
         " inner join Mst_ServiceTax MS On MCom.ServiceTaxID = MS.ServiceTaxID " & _
         " Inner Join Mst_ExchangeType ME On TC.ExchangeTypeID = ME.ExchangeTypeID " & _
         " Inner Join Mst_Client MC On TPI.ClientID = MC.ClientID "
   
   If Gen_DBType = "MDB" Then
      wc = " Where TPI.InvoiceDate >= #" & TxtFromDate.Value & "# And TPI.InvoiceDate <= #" & TxtToDate.Value & "# "
   Else
      wc = " Where TPI.InvoiceDate >= '" & TxtFromDate.Value & "' And TPI.InvoiceDate <= '" & TxtToDate.Value & "' "
   End If
               
   If CmbPreFix.Text <> "" Then
      wc = wc & " And  ML.LocationName = '" & CmbPreFix.Text & "'"
   Else
      If Gen_WcLocation <> "" Then
         wc = wc & " And ML." & Gen_WcLocation
      End If
   End If
         
   If CmbNonNcdexClient.Text <> "" Then
      wc = wc & " And  MC.ClientName = '" & CmbNonNcdexClient.Text & "'"
   End If
   
   wc = wc & " And  TPI.Flag = 'P'"
   
   If TxtSInvoiceNo.Text <> "" Then
      wc = wc & " And  TPI.InvoiceNo = '" & Replace(TxtSInvoiceNo.Text, "SD", "") & "'"
   End If
   
   wc = wc & " And  TPI.BatchID = 0 " '
               
   tmp1 = " Order By ML.LocationName,TPI.BOEID "
   
   tmp = tmp & " " & wc & " " & tmp1
End If

Call TmpTable

If LCase(CmbInvoiceType.Text) = "nnnr" Then
   
   tmp = " Select PINo, sum(InvoiceAmount) GSLAmount From Txn_InvMonGSLDetail Where ExchangeTypeID = 1 And PINo is not null"
   tmp = tmp & " and yearmonth between '" & Format(DateSerial(TxtFromDate.Year, TxtFromDate.Month, 1), "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
   tmp = tmp & " group by PINo "
   Call Tmp7Table
   
   tmp = "Select Distinct PINo from  Txn_MonthwiseInvoiceData  Where TransferFlag is null And yearmonth between '" & Format(DateSerial(TxtFromDate.Year, TxtFromDate.Month, 1), "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "' "
   Call Tmp6Table
   
   xCol = 3
ElseIf LCase(CmbInvoiceType.Text) = "reservation" Then
   xCol = 3
ElseIf LCase(CmbInvoiceType.Text) = "adhoc" Then
   xCol = 3

Else
   xCol = 1
End If

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 2
'   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For x = 1 To TmpRs.RecordCount
       I = MSHFlexGrid1.Rows - 1
       
       If LCase(CmbInvoiceType.Text) = "nnnr" Then
          If TmpRs6.RecordCount > 0 Then
             TmpRs6.MoveFirst
             TmpRs6.Find "PINo = " & TmpRs!InvoiceNo
             If TmpRs6.EOF = True Then
                If TmpRs7.RecordCount > 0 Then
                   TmpRs7.MoveFirst
                   TmpRs7.Find "PINo = " & TmpRs!InvoiceNo
    
                 'If Format(TmpRs!Amount, "#######") <> Format(TmpRs7!GSLAmount, "#######") Then GoTo nextrow
                   If ChkMismatch.Value = 0 Then
                      If TmpRs7.EOF = True Then GoTo NextRow
                      If Abs(TmpRs!Amount - TmpRs7!GSLAmount) > 1 Then GoTo NextRow
                   Else
                      If TmpRs7.EOF = False Then
                         If Abs(TmpRs!Amount - TmpRs7!GSLAmount) < 1 Then GoTo NextRow
                      End If
                   End If
                End If
             Else
                If ChkMismatch.Value = 1 Then
                   GoTo NextRow
                End If
             End If
          Else
             If TmpRs7.RecordCount > 0 Then
                TmpRs7.MoveFirst
                TmpRs7.Find "PINo = " & TmpRs!InvoiceNo
    
                 'If Format(TmpRs!Amount, "#######") <> Format(TmpRs7!GSLAmount, "#######") Then GoTo nextrow
                If ChkMismatch.Value = 0 Then
                   If TmpRs7.EOF = True Then GoTo NextRow
                   If Abs(TmpRs!Amount - TmpRs7!GSLAmount) > 1 Then GoTo NextRow
                Else
                   If TmpRs7.EOF = False Then
                      If Abs(TmpRs!Amount - TmpRs7!GSLAmount) < 1 Then GoTo NextRow
                   End If
                End If
             End If
          End If
       End If
       
       For C = 0 To TmpRs.Fields.Count - xCol
           MSHFlexGrid1.TextMatrix(I, C + 1) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       If LCase(CmbInvoiceType.Text) = "nnnr" Then
          MSHFlexGrid1.TextMatrix(I, 1) = "NN" & MSHFlexGrid1.TextMatrix(I, 1)
       ElseIf LCase(CmbInvoiceType.Text) = "reservation" Then
          MSHFlexGrid1.TextMatrix(I, 1) = "RI" & MSHFlexGrid1.TextMatrix(I, 1)
       ElseIf LCase(CmbInvoiceType.Text) = "adhoc" Then
          MSHFlexGrid1.TextMatrix(I, 1) = "AH" & MSHFlexGrid1.TextMatrix(I, 1)
       Else
          MSHFlexGrid1.TextMatrix(I, 1) = "SD" & MSHFlexGrid1.TextMatrix(I, 1)
       End If
       MSHFlexGrid1.TextMatrix(I, 2) = Format(MSHFlexGrid1.TextMatrix(I, 2), "dd-mmm-yyyy")
      
       MSHFlexGrid1.TextMatrix(I, 3) = Format(MSHFlexGrid1.TextMatrix(I, 3), "mmm-yyyy")
       If MSHFlexGrid1.TextMatrix(I, 14) = "P" Then
          MSHFlexGrid1.TextMatrix(I, 14) = "Pending"
       ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "C" Then
          MSHFlexGrid1.TextMatrix(I, 14) = "Confirmed"
       ElseIf MSHFlexGrid1.TextMatrix(I, 14) = "A" Then
          MSHFlexGrid1.TextMatrix(I, 14) = "Approved"
       End If
       
       If LCase(CmbInvoiceType.Text) = "nnnr" And ChkMismatch.Value = 0 Then
          
          tmp = "Select Count(*) 'TotInv'  from Txn_InvMonGSLDetail Where (OldRate is not null Or OldBillingCycleID is not null Or OldBillingUnit is not null ) "
          tmp = tmp & "  And PINo = " & TmpRs!InvoiceNo & " And YearMonth = '" & TmpRs!BillingMonth & "' "
       
          Call Tmp1Table
          
          If IsNull(TmpRs1!TotInv) = False Then
             If TmpRs1!TotInv > 0 Then
                For C = 0 To MSHFlexGrid1.Cols - 1
                    MSHFlexGrid1.Col = C
                    MSHFlexGrid1.row = I
                    MSHFlexGrid1.CellBackColor = vbRed
                Next
             End If
          End If
          
          MSHFlexGrid1.TextMatrix(I, 16) = GetADhocInvoiceAmount(TmpRs!LocationID, TmpRs!ClientID, I)
            If IsDiscountAppliedInDC(TmpRs!InvoiceNo) > 0 Then
                MSHFlexGrid1.TextMatrix(I, 8) = MSHFlexGrid1.TextMatrix(I, 7)
            End If
          
          If Val(MSHFlexGrid1.TextMatrix(I, 16)) > 0 Then
             For C = 0 To MSHFlexGrid1.Cols - 1
                 MSHFlexGrid1.Col = C
                 MSHFlexGrid1.row = I
                 MSHFlexGrid1.CellBackColor = vbRed
             Next
          End If
       End If
       
       
       If LCase(CmbInvoiceType.Text) = "reservation" And ChkMismatch.Value = 0 Then 'Condition for reservation added on 31-01-2025
          
          MSHFlexGrid1.TextMatrix(I, 16) = GetADhocInvoiceAmount_Reservation(TmpRs!LocationID, TmpRs!ClientID, TmpRs!ContractID, I)
          
          If Val(MSHFlexGrid1.TextMatrix(I, 16)) > 0 Then
             For C = 0 To MSHFlexGrid1.Cols - 1
                 MSHFlexGrid1.Col = C
                 MSHFlexGrid1.row = I
                 MSHFlexGrid1.CellBackColor = vbRed
             Next
          End If
       End If
       
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       
NextRow:       TmpRs.MoveNext
       
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   CmdSave.Enabled = True
End If
LblTotalRecord.Caption = MSHFlexGrid1.Rows - 2

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame0.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Frame0.Enabled = True
Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
CmbPreFix.Text = ""
CmbNonNcdexClient.Text = ""
TxtSInvoiceNo.Text = ""
LblTotalRecord.Caption = "-"
Label6.Visible = False
Label8.Visible = False
ChkMismatch.Visible = False
ChkMismatch.Value = 0
CmdDelete.Visible = False
mNoNavcode = 0

End Sub

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
    If .TextMatrix(.row, 2) = "" Then Exit Sub
    If .Col = 0 Then
        If (.CellBackColor) <> vbRed Then
           If .TextMatrix(.row, 0) = strUnChecked Then
              .TextMatrix(.row, 0) = strChecked
              '.TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.Row, 8))
           Else
              .TextMatrix(.row, 0) = strUnChecked
              .TextMatrix(.row, 15) = ""
              '.TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.Row, 8))
           End If
        End If
    End If
    mSelRow = MSHFlexGrid1.row
    mSelCol = MSHFlexGrid1.Col
    
    If mSelCol <> 15 Then Exit Sub
    If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = strUnChecked Then Exit Sub
    
    TxtRemark.Visible = True
    TxtRemark.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 15)
    
    With TxtRemark
        .Left = (MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft) - 30
        .Top = (MSHFlexGrid1.Top + MSHFlexGrid1.CellTop) - 30
        .Width = MSHFlexGrid1.CellWidth + 35
        .Height = MSHFlexGrid1.CellHeight
        .ZOrder
        .SetFocus
    End With
End With

End Sub


Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
End Sub


Private Sub TxtRemark_LostFocus()
If TxtRemark = "" Then Exit Sub
If IsNull(TxtRemark) Then
   MsgBox "Blank Remark is not Allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If
'mSelRow = MSHFlexGrid1.Row
'mSelCol = MSHFlexGrid1.Col
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = TxtRemark
TxtRemark.Visible = False
MSHFlexGrid1.SetFocus

End Sub





Private Sub TxtToDate_LostFocus()
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End Sub

Private Sub CmbNonNcdexClient_GotFocus()
If CmbInvoiceType.Text = "" Then
   MsgBox "Please select Invoice Type"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
tmp = CmbNonNcdexClient.Text
If LCase(CmbInvoiceType.Text) = "nnnr" Then
   Call TableMst_Client("where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
Else
   Call TableMst_Client("where ExchangeTypeId = 2 ")
End If
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
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 0

If LCase(CmbInvoiceType.Text) <> "nnnr" And LCase(CmbInvoiceType.Text) <> "reservation" Then 'Condition for reservation added on 31-01-2025
   MSHFlexGrid1.Cols = 15
Else
   MSHFlexGrid1.Cols = 17
End If
MSHFlexGrid1.TextMatrix(0, 0) = "Select"
MSHFlexGrid1.TextMatrix(0, 1) = "Invoice No"
MSHFlexGrid1.TextMatrix(0, 2) = "Invoice Date"
MSHFlexGrid1.TextMatrix(0, 3) = "Billing Month"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "Exchange Type"
MSHFlexGrid1.TextMatrix(0, 6) = "Client"
MSHFlexGrid1.TextMatrix(0, 7) = "Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "Discount"
MSHFlexGrid1.TextMatrix(0, 9) = "Premium"
MSHFlexGrid1.TextMatrix(0, 10) = "Batch ID"
MSHFlexGrid1.TextMatrix(0, 11) = "GST Group Code" '"Service Tax NAVCode"
MSHFlexGrid1.TextMatrix(0, 12) = "G_UID"
MSHFlexGrid1.TextMatrix(0, 13) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 14) = "Status"
'MSHFlexGrid1.TextMatrix(0, 15) = "Remark for Deletion"
If LCase(CmbInvoiceType.Text) = "nnnr" Then
   MSHFlexGrid1.TextMatrix(0, 15) = "Remark for Deletion"
   MSHFlexGrid1.TextMatrix(0, 16) = "ADHOC Amount"
End If

If LCase(CmbInvoiceType.Text) = "reservation" Then 'Condition for reservation added on 31-01-2025
   MSHFlexGrid1.TextMatrix(0, 15) = "Contract ID"
   MSHFlexGrid1.TextMatrix(0, 16) = "ADHOC Amount"
End If

MSHFlexGrid1.RowHeight(0) = 900
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 3200
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 0
MSHFlexGrid1.ColWidth(13) = 3000
MSHFlexGrid1.ColWidth(14) = 1600
MSHFlexGrid1.ColWidth(15) = 3000
End Sub


Private Function CheckSelect() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckSelect = Retval
End Function

Private Function GetADhocInvoiceAmount(mLocID_ As Double, mClientID_ As Double, i_ As Variant)
Dim Retval As Variant
Retval = 0

tmp = "Select LocationID,ClientID,Sum(Amount) As 'Amount' From Txn_TempInvoice "
tmp = tmp & " Where IETypeID = 7 And LocationID = " & mLocID_ & " And ClientID = " & mClientID_ & " And Month(InvoiceDate) = " & Month(CDate(MSHFlexGrid1.TextMatrix(i_, 3))) & "  And Year(InvoiceDate) = " & Year(CDate(MSHFlexGrid1.TextMatrix(i_, 3))) & " And Flag <> 'R' "
tmp = tmp & " Group By LocationID,ClientID "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1!Amount
End If

GetADhocInvoiceAmount = Retval

End Function
Private Function GetADhocInvoiceAmount_Reservation(mLocID_ As Double, mClientID_ As Double, mContractID_ As Double, i_ As Variant)
Dim Retval As Variant
Retval = 0

tmp = "Select LocationID,ClientID,ContractID,Sum(Amount) As 'Amount' From Txn_TempInvoice "
tmp = tmp & " Where IETypeID = 4 And ContractID = " & mContractID_ & " And LocationID = " & mLocID_ & " And ClientID = " & mClientID_ & " And Month(InvoiceDate) = " & Month(CDate(MSHFlexGrid1.TextMatrix(i_, 3))) & "  And Year(InvoiceDate) = " & Year(CDate(MSHFlexGrid1.TextMatrix(i_, 3))) & " And Flag <> 'R' "
tmp = tmp & " Group By LocationID,ClientID,ContractID "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1!Amount
End If

GetADhocInvoiceAmount_Reservation = Retval

End Function

Private Function CheckRemarks() As Boolean
Dim Retval As Boolean
Retval = True
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If MSHFlexGrid1.TextMatrix(I, 15) = "" Then
          Retval = False
          Exit For
       End If
    End If
Next
CheckRemarks = Retval
End Function

Private Function IsDiscountAppliedInDC(mInvoiceNo As Double)
Dim Retval As Variant
Retval = 0

tmp = "select distinct b.PINo, e.DepConID from Txn_ProInvoice a"
tmp = tmp & " inner join Txn_InvMonGSLDetail b on a.InvoiceNo = b.PINo"
tmp = tmp & " inner join Mst_GSL c on b.Mst_GSL_ID = c.Mst_GSL_ID"
tmp = tmp & " inner join Txn_DepositContractGodownDetail d on c.DepConDID = d.DepConDID"
tmp = tmp & " inner join Txn_DepositContract e on d.DepConID = e.DepConID"
tmp = tmp & " Where a.InvoiceNo = " & mInvoiceNo & " And IsNull(E.DiscountApplied, 0) = 1"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Retval = TmpRs1.RecordCount
End If

IsDiscountAppliedInDC = Retval

End Function
