VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_ImportReservationInvoice 
   Caption         =   "Import Reservation Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   20
      Top             =   3645
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   16
      Top             =   50
      Width           =   15015
      Begin VB.CheckBox ChkRRNInvoice 
         Caption         =   "RRN Invoice"
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
         TabIndex        =   7
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CheckBox ChkRejectedInvoice 
         Caption         =   "Rejected Invoice"
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
         TabIndex        =   6
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CheckBox ChkPWRInvoice 
         Caption         =   "PWR Invoice"
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
         TabIndex        =   5
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox ChkNNNRYearlyInvoice 
         Caption         =   "NNNR Yearly Invoice"
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
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox Chk7DaysInvoice 
         Caption         =   "7 Days Invoice"
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
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox ChkTDSAdjEntryNo 
         Caption         =   "TDS Adj Entry No"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox ChkCashReceipt 
         Caption         =   "Cash Receipt"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox ChkAdHocInvoice 
         Caption         =   "Ad-Hoc Invoice"
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
         TabIndex        =   8
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CheckBox ChkNNNRInvoice 
         Caption         =   "NNNR Invoice"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox ChkCommtrackInvoice 
         Caption         =   "Commtrack Invoice"
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
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox ChkReservationInvoice 
         Caption         =   "Reservation Invoice"
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
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
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
         Left            =   6720
         TabIndex        =   18
         Top             =   270
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
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
         Height          =   350
         Left            =   9960
         TabIndex        =   11
         Top             =   270
         Width           =   1620
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Left            =   11580
         TabIndex        =   12
         Top             =   270
         Width           =   1620
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
         Height          =   350
         Left            =   13200
         TabIndex        =   13
         Top             =   270
         Width           =   1620
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
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
         Height          =   350
         Left            =   8340
         TabIndex        =   17
         Top             =   270
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   720
         Width           =   660
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   5760
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   10160
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_ImportReservationInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim WrkSheet As String
Dim RowCount As Double
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim objFScp As Scripting.FileSystemObject
Dim objFilecp As File
Dim objTScp As TextStream
Dim sFileLine As String
Dim sFileLineNo As Long
Dim temp As Variant
Dim arecordset As String
Dim strFilename As String
Dim strsql As String
Dim strdata As String
Dim t_line As String
Dim strfile As String
Dim strqry As String
Dim sArray(0 To 10) As String
Dim mInvoiceNo As Variant

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGetFile_Click()

On Error GoTo errortrap

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "CSV files(*.csv)|*.csv" '|Text files(*.txt)|*.txt"
     .MaxFileSize = &H7FFF ' 32KB filename buffer
    'same as .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNLongNames or cdlOFNExplorer --Or cdlOFNAllowMultiselect
     .Flags = &H4 Or &H800 Or &H40000 Or &H80000 Or &H200
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
mFilenames = CommonDialog1.FileName
If LblFile.Caption <> Path Then
 '  Call Clear
End If
LblFile.Caption = Path
If LblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"

End Sub

Private Sub CmdImport_Click()


Call ReadFiles(LblFile.Caption)
CmdUpdate.Enabled = True

End Sub

Private Sub CmdUpdate_Click()
Dim mR, mC As Double
Dim mInvoice, mstr, mCRID, mSMPrefix As Variant

'Commented on 24 Aug 2017 For CSV Upload of Reservation Invoice
If ChkReservationInvoice.Value = vbUnchecked And ChkNNNRYearlyInvoice.Value = vbUnchecked And ChkCommtrackInvoice.Value = vbUnchecked And ChkNNNRInvoice.Value = vbUnchecked And ChkAdHocInvoice.Value = vbUnchecked And ChkCashReceipt.Value = vbUnchecked And ChkPWRInvoice.Value = vbUnchecked And ChkRejectedInvoice.Value = vbUnchecked And ChkRRNInvoice.Value = vbUnchecked And Chk7DaysInvoice.Value = vbUnchecked And ChkTDSAdjEntryNo.Value = vbUnchecked Then
    MsgBox "Please, Select at-least one of the Check Box to process."
    Exit Sub
Else
    MsgBox "Please Wait till Next Message."
    
''-- Process to Update the Invoice / CR / .. Data

'Commented on 24 Aug 2017 For CSV Upload of Reservation Invoice
    Call TableTxn_TIDetails
    ''-- Reservation Invoice
    If ChkReservationInvoice.Value = vbChecked Then
       Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'RESERVATION' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> '' ")
       If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
        Call TableTxn_ReservationInvoice
           ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
           For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount
            If RsTxn_ReservationInvoice.RecordCount > 0 Then
               RsTxn_ReservationInvoice.MoveFirst
            End If
            mInvoice = Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "RI", "")
            RsTxn_ReservationInvoice.Find "InvoiceNo= " & Val(mInvoice) & " "
            If Not RsTxn_ReservationInvoice.EOF Then
               If RsTxn_TIDetails.RecordCount > 0 Then
                  RsTxn_TIDetails.MoveFirst
                  RsTxn_TIDetails.Find "EntryNo= " & Val(RsTxn_InvoicesForNAVISION!EntryNo)
                  If RsTxn_TIDetails.EOF = True Then
                     RsTxn_TIDetails.AddNew
'                     RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                     RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                     RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                     RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                     RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                     RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                     RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                     RsTxn_TIDetails!PIID = RsTxn_ReservationInvoice!InvoiceNo
                     RsTxn_TIDetails!LocationID = RsTxn_ReservationInvoice!LocationID

                     tmp = " Select TRI.ContractID,MSC.ClientID from Txn_ReservationInvoice TRI"
                     tmp = tmp & " INNER join MST_StorageContract MSC ON TRI.ContractID=MSc.ContractID"
                     tmp = tmp & " Where InvoiceNo = " & RsTxn_ReservationInvoice!InvoiceNo & ""
                     Call TmpTable

                     RsTxn_TIDetails!ClientID = TmpRs!ClientID
                     RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                     RsTxn_TIDetails!ExchangeTypeID = 1 'Reservation Income is Non-Ncdex
                     RsTxn_TIDetails!RecAmount = 0
                     RsTxn_TIDetails!RecTDSAmount = 0
                     RsTxn_TIDetails!PIFlag = "Reservation"
                     RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                     'RsTxn_TIDetails!EntryFlag = "System"
                     RsTxn_TIDetails!G_UID = GetGUID
                     RsTxn_TIDetails!CreatedDate = Now
                     RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                     RsTxn_TIDetails.Update
                  End If
                  RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
                  RsTxn_InvoicesForNAVISION.Update
               End If
            End If
           RsTxn_InvoicesForNAVISION.MoveNext
           ProgressBar1.Value = mR
           Next mR
       End If
    End If

    ''-- Commtrack Invoice
    If ChkCommtrackInvoice.Value = vbChecked Then
       Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'COMMTRACK' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> '' ")
       If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
        Call TableTxn_MonthlyCommtrackInvoice
           ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
           For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount

            If RsTxn_MonthlyCommtrackInvoice.RecordCount > 0 Then
               RsTxn_MonthlyCommtrackInvoice.MoveFirst
            End If
            mInvoice = Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "COMMTRACK", "")
            RsTxn_MonthlyCommtrackInvoice.Find "PINo= " & Val(mInvoice) & " "
            If Not RsTxn_MonthlyCommtrackInvoice.EOF Then
               If RsTxn_TIDetails.RecordCount > 0 Then
                  RsTxn_TIDetails.MoveFirst
                  RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
                  If RsTxn_TIDetails.EOF = True Then
                     RsTxn_TIDetails.AddNew
'                     RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                     RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                     RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                     RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                     RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                     RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                     RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                     RsTxn_TIDetails!PIID = RsTxn_MonthlyCommtrackInvoice!InvoiceID
                     RsTxn_TIDetails!LocationID = RsTxn_MonthlyCommtrackInvoice!LocationID
                     RsTxn_TIDetails!DepositoryID = GetDepositoryIDFromNavCode(RsTxn_MonthlyCommtrackInvoice!CPNAVCode)
                     RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                     RsTxn_TIDetails!ExchangeTypeID = 2
                     RsTxn_TIDetails!PIFlag = "Commtrack"
                     RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                     RsTxn_TIDetails!RecAmount = 0
                     RsTxn_TIDetails!RecTDSAmount = 0
                     RsTxn_TIDetails!RecOtherDed = 0
                     'RsTxn_TIDetails!EntryFlag = "System"
                     RsTxn_TIDetails!G_UID = GetGUID
                     RsTxn_TIDetails!CreatedDate = Now
                     RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                     RsTxn_TIDetails.Update

                     RsTxn_MonthlyCommtrackInvoice!InvoiceDate = RsTxn_TIDetails!TIDate 'Format(MSHFlexGrid3.TextMatrix(I, 21), Gen_DatFormat)
                     RsTxn_MonthlyCommtrackInvoice!InvoiceNo = Trim(RsTxn_InvoicesForNAVISION!TINo) 'Trim(MSHFlexGrid3.TextMatrix(I, 20))
                     RsTxn_MonthlyCommtrackInvoice!InvSTAmt = Val(RsTxn_InvoicesForNAVISION!STAmt) 'Val(MSHFlexGrid3.TextMatrix(I, 22))
                     RsTxn_MonthlyCommtrackInvoice!InvoiceAmt = Val(RsTxn_InvoicesForNAVISION!TIAmt) 'Val(MSHFlexGrid3.TextMatrix(I, 23))
                     RsTxn_MonthlyCommtrackInvoice!UpdatedBy = Gen_UserLoginID
                     RsTxn_MonthlyCommtrackInvoice!UpdatedDate = Now
                     RsTxn_MonthlyCommtrackInvoice.Update
                  End If
                  RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
                  RsTxn_InvoicesForNAVISION.Update
              End If
           End If
           RsTxn_InvoicesForNAVISION.MoveNext
           ProgressBar1.Value = mR
           Next mR
       End If
    End If
    ''-- 7-Days Invoice
    If Chk7DaysInvoice.Value = vbChecked Then
        Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = '7Days' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> '' ")
        If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
        Call TableTxn_CDTFBOE
        ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
        For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount
            If RsTxn_CDTFBOE.RecordCount > 0 Then
               RsTxn_CDTFBOE.MoveFirst
            End If
            RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "SD", "")) & ""
            If Not RsTxn_CDTFBOE.EOF Then
               If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
               RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
               If RsTxn_TIDetails.EOF = True Then
                  RsTxn_TIDetails.AddNew
'                  RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                  RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                  RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                  RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                  RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                  RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                  RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                  RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
                  RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
                  RsTxn_TIDetails!ClientID = RsTxn_CDTFBOE!ClientID
                  RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                  RsTxn_TIDetails!ExchangeTypeID = 2
                  RsTxn_TIDetails!RecAmount = 0
                  RsTxn_TIDetails!RecTDSAmount = 0
                  RsTxn_TIDetails!RecOtherDed = 0
                  RsTxn_TIDetails!PIFlag = "SD"
                  RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                  'RsTxn_TIDetails!EntryFlag = "System"
                  RsTxn_TIDetails!G_UID = GetGUID
                  RsTxn_TIDetails!CreatedDate = Now
                  RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                  RsTxn_TIDetails.Update
               End If
            RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
            RsTxn_InvoicesForNAVISION.Update
            End If
        RsTxn_InvoicesForNAVISION.MoveNext
        ProgressBar1.Value = mR
        Next
        End If
    End If
    ''-- NNNR Invoice
    If ChkNNNRInvoice.Value = vbChecked Then
       Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'NNNR' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> '' ")
       If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
            Call TableTxn_ProInvoice
            ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
            For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount
            If RsTxn_ProInvoice.RecordCount > 0 Then
               RsTxn_ProInvoice.MoveFirst
            End If
            RsTxn_ProInvoice.Find "InvoiceNo=" & Val(Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "NN", "")) & ""
            If Not RsTxn_ProInvoice.EOF Then
               If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
               RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
               If RsTxn_TIDetails.EOF = True Then
                  RsTxn_TIDetails.AddNew
'                  RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                  RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                  RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                  RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                  RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                  RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                  RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                  RsTxn_TIDetails!PIID = RsTxn_ProInvoice!InvoiceNo
                  RsTxn_TIDetails!LocationID = RsTxn_ProInvoice!LocationID
                  RsTxn_TIDetails!ClientID = RsTxn_ProInvoice!ClientID
                  RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                  RsTxn_TIDetails!ExchangeTypeID = RsTxn_ProInvoice!ExchangeTypeID
                  RsTxn_TIDetails!RecAmount = 0
                  RsTxn_TIDetails!RecTDSAmount = 0
                  RsTxn_TIDetails!RecOtherDed = 0
                  RsTxn_TIDetails!PIFlag = "NNNR"
                  RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                  'RsTxn_TIDetails!EntryFlag = "System"
                  RsTxn_TIDetails!G_UID = GetGUID
                  RsTxn_TIDetails!CreatedDate = Now
                  RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                  RsTxn_TIDetails.Update
               End If
            RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
            RsTxn_InvoicesForNAVISION.Update
            End If
            RsTxn_InvoicesForNAVISION.MoveNext
            ProgressBar1.Value = mR
            Next mR
       End If
    End If
    ''-- NNNR Yearly Invoice
    If ChkNNNRYearlyInvoice.Value = vbChecked Then
        Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'NNNRY' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> '' ")
        If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
            ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
            Call TableTxn_CDTFBOE
            For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount
            If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
            RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "NNNRY", "")) & ""
            If Not RsTxn_CDTFBOE.EOF Then
               If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
               RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
               If RsTxn_TIDetails.EOF = True Then
                  RsTxn_TIDetails.AddNew
'                  RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                  RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                  RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                  RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                  RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                  RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                  RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                  RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
                  RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
                  RsTxn_TIDetails!ClientID = RsTxn_CDTFBOE!ClientID
                  RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                  RsTxn_TIDetails!ExchangeTypeID = 1
                  RsTxn_TIDetails!RecAmount = 0
                  RsTxn_TIDetails!RecTDSAmount = 0
                  RsTxn_TIDetails!RecOtherDed = 0
                  RsTxn_TIDetails!PIFlag = "NNNRY"
                  RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                  'RsTxn_TIDetails!EntryFlag = "System"
                  RsTxn_TIDetails!G_UID = GetGUID
                  RsTxn_TIDetails!CreatedDate = Now
                  RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                  RsTxn_TIDetails.Update
               End If
            RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
            RsTxn_InvoicesForNAVISION.Update
            End If
            RsTxn_InvoicesForNAVISION.MoveNext
            ProgressBar1.Value = mR
            Next mR
        End If
    End If
    ''-- PWR Invoice
    If ChkPWRInvoice.Value = vbChecked Then
        Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'PWR' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> ''  ")
        If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
           ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
           Call TableTxn_CDTFBOE
           For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount

            If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
            RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "PWR", "")) & ""

            If Not RsTxn_CDTFBOE.EOF Then
               If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
               RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
               If RsTxn_TIDetails.EOF = True Then
                  RsTxn_TIDetails.AddNew
'                  RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                  RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                  RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                  RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                  RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                  RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                  RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                  RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
                  RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
                  RsTxn_TIDetails!ClientID = RsTxn_CDTFBOE!ClientID
                  RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                  RsTxn_TIDetails!ExchangeTypeID = 1
                  RsTxn_TIDetails!RecAmount = 0
                  RsTxn_TIDetails!RecTDSAmount = 0
                  RsTxn_TIDetails!RecOtherDed = 0
                  RsTxn_TIDetails!PIFlag = "PWR"
                  RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                  'RsTxn_TIDetails!EntryFlag = "System"
                  RsTxn_TIDetails!G_UID = GetGUID
                  RsTxn_TIDetails!CreatedDate = Now
                  RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                  RsTxn_TIDetails.Update
               End If
           RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
           RsTxn_InvoicesForNAVISION.Update
           End If
           RsTxn_InvoicesForNAVISION.MoveNext
           ProgressBar1.Value = mR
           Next mR
        End If
    End If
    ''-- Rejected Invoice
    If ChkRejectedInvoice.Value = vbChecked Then
       Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'Rejected' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> '' ")
       If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
        ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
        Call TableTxn_CDTFBOE

            For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount
                If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
                RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "RJ", "")) & ""
                If Not RsTxn_CDTFBOE.EOF Then
                   If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
                   RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
                   If RsTxn_TIDetails.EOF = True Then
                      RsTxn_TIDetails.AddNew
'                      RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                      RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                      RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                      RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                      RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                      RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                      RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                      RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
                      RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
                      RsTxn_TIDetails!ClientID = RsTxn_CDTFBOE!ClientID
                      RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                      RsTxn_TIDetails!ExchangeTypeID = 2
                      RsTxn_TIDetails!RecAmount = 0
                      RsTxn_TIDetails!RecTDSAmount = 0
                      RsTxn_TIDetails!RecOtherDed = 0
                      RsTxn_TIDetails!PIFlag = "RJ"
                      RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                      'RsTxn_TIDetails!EntryFlag = "System"
                      RsTxn_TIDetails!G_UID = GetGUID
                      RsTxn_TIDetails!CreatedDate = Now
                      RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                      RsTxn_TIDetails.Update
                   End If
                RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
                RsTxn_InvoicesForNAVISION.Update
                End If
            RsTxn_InvoicesForNAVISION.MoveNext
            ProgressBar1.Value = mR
            Next mR

       End If
    End If
    ''-- RRN Invoice
    If ChkRRNInvoice.Value = vbChecked Then
        Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'RRN' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> ''  ")
        If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
            ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
            Call TableTxn_CDTFBOE
            For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount
                If RsTxn_CDTFBOE.RecordCount > 0 Then RsTxn_CDTFBOE.MoveFirst
                RsTxn_CDTFBOE.Find "BOEID=" & Val(Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "RRN", "")) & ""
                If Not RsTxn_CDTFBOE.EOF Then
                   If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
                   RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
                   If RsTxn_TIDetails.EOF = True Then
                      RsTxn_TIDetails.AddNew
'                      RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                      RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                      RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                      RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                      RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                      If Val(RsTxn_InvoicesForNAVISION!STAmt) > 0 And RsTxn_TIDetails!TISTAmount = 0 Then
                         a = a
                      End If
                      RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                      RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                      RsTxn_TIDetails!PIID = RsTxn_CDTFBOE!BOEID
                      RsTxn_TIDetails!LocationID = RsTxn_CDTFBOE!LocationID
                      RsTxn_TIDetails!ClientID = RsTxn_CDTFBOE!ClientID
                      RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                      RsTxn_TIDetails!ExchangeTypeID = 2
                      RsTxn_TIDetails!RecAmount = 0
                      RsTxn_TIDetails!RecTDSAmount = 0
                      RsTxn_TIDetails!RecOtherDed = 0
                      RsTxn_TIDetails!PIFlag = "RRN"
                      RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                      'RsTxn_TIDetails!EntryFlag = "System"
                      RsTxn_TIDetails!G_UID = GetGUID
                      RsTxn_TIDetails!CreatedDate = Now
                      RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                      RsTxn_TIDetails.Update
                   End If
                   RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
                   RsTxn_InvoicesForNAVISION.Update
                End If
                RsTxn_InvoicesForNAVISION.MoveNext
                ProgressBar1.Value = mR
            Next mR
        End If
    End If
    ''-- Ad-Hoc Invoice
    If ChkAdHocInvoice.Value = vbChecked Then
       Call TableTxn_InvoicesForNAVISION(" Where InvoiceType = 'ADHOC' And Invoice = 'Invoice' And ISNULL(ReverseInvoice,'') = '' And ISNULL(DataTransfer,'') = 'T' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(EntryNo,0) <> 0) And ISNULL(PINo,'') <> '' ")
       If RsTxn_InvoicesForNAVISION.RecordCount > 0 Then
        ProgressBar1.Max = RsTxn_InvoicesForNAVISION.RecordCount
        Call TableTxn_TempInvoice
        For mR = 1 To RsTxn_InvoicesForNAVISION.RecordCount
            If RsTxn_TempInvoice.RecordCount > 0 Then RsTxn_TempInvoice.MoveFirst
            mInvoice = Replace(UCase(RsTxn_InvoicesForNAVISION!PINo), "AH", "")
            RsTxn_TempInvoice.Find "PINo=" & Val(mInvoice) & ""

            If Not RsTxn_TempInvoice.EOF Then
               If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
               RsTxn_TIDetails.Find "EntryNo=" & Val(RsTxn_InvoicesForNAVISION!EntryNo)
               If RsTxn_TIDetails.EOF = True Then
                  RsTxn_TIDetails.AddNew
'                  RsTxn_TIDetails!TIID = GetMaxID("TIID", "Txn_TIDetails")
                  RsTxn_TIDetails!TINo = Trim(RsTxn_InvoicesForNAVISION!TINo)
                  RsTxn_TIDetails!TIDate = Format(RsTxn_InvoicesForNAVISION!TIDate, Gen_DatFormat)
                  RsTxn_TIDetails!TIAmount = Val(RsTxn_InvoicesForNAVISION!TIAmt)
                  RsTxn_TIDetails!TISTAmount = Val(RsTxn_InvoicesForNAVISION!STAmt)
                  If Val(RsTxn_InvoicesForNAVISION!STAmt) > 0 And RsTxn_TIDetails!TISTAmount = 0 Then
                         a = a
                  End If
                  RsTxn_TIDetails!PostInvoiceNo = Trim(RsTxn_InvoicesForNAVISION!PostInvoiceNo)
                  RsTxn_TIDetails!EntryNo = Val(RsTxn_InvoicesForNAVISION!EntryNo)
                  RsTxn_TIDetails!PIID = RsTxn_TempInvoice!PINo
                  RsTxn_TIDetails!LocationID = RsTxn_TempInvoice!LocationID
                  RsTxn_TIDetails!ClientID = RsTxn_TempInvoice!ClientID
                  RsTxn_TIDetails!NavCode = Trim(RsTxn_InvoicesForNAVISION!NavCode)
                  RsTxn_TIDetails!ExchangeTypeID = RsTxn_TempInvoice!ExchangeTypeID
                  RsTxn_TIDetails!RecAmount = 0
                  RsTxn_TIDetails!RecTDSAmount = 0
                  RsTxn_TIDetails!RecOtherDed = 0
                  RsTxn_TIDetails!PIFlag = "ADHOC"
                  RsTxn_TIDetails!SMID = RsTxn_InvoicesForNAVISION!PINo
                  'RsTxn_TIDetails!EntryFlag = "System"
                  RsTxn_TIDetails!G_UID = GetGUID
                  RsTxn_TIDetails!CreatedDate = Now
                  RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
                  RsTxn_TIDetails.Update
               End If
               RsTxn_InvoicesForNAVISION!EntryNoTransfer = "T"
               RsTxn_InvoicesForNAVISION.Update
            End If
            RsTxn_InvoicesForNAVISION.MoveNext
            ProgressBar1.Value = mR
        Next mR
       End If
    End If
    ''-- Cash Receipt Entry No
    
    If ChkCashReceipt.Value = vbChecked Then
       Call TableTxn_CashRepeiptsForNAVISION(" Where Concurrency <= '2020-01-25' And InvoiceType = 'CR' And AccountType = 'Customer' And ISNULL(DataTransfer,'') = 'T' And ISNULL(ReverseInvoice,'') = ''  And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(CRADJEntryNo,0) <> 0) And ISNULL(PINo,'') <> '' And ISNULL(EntryNoTransfer,'') = '' ")
       If RsTxn_CashRepeiptsForNAVISION.RecordCount > 0 Then
          ProgressBar1.Max = RsTxn_CashRepeiptsForNAVISION.RecordCount
          For mR = 1 To RsTxn_CashRepeiptsForNAVISION.RecordCount
                
            mstr = Split(RsTxn_CashRepeiptsForNAVISION!PINo, "-")
            If UBound(mstr) = 1 Then
               mCRID = mstr(0)
               mCRID = Replace(UCase(mCRID), "CR", "")
               mSMPrefix = mstr(1)
               
               tmp = "Update Txn_NewCashReceiptTIDetail Set DocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',EntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & ""
               tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_NewCashReceiptTIDetail Where CRID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') And Isnull(CRAAID,0)=0 And CRID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "' And Isnull(DocNo,'') ='' "
               Call TmpTable
              
               '--for Advace
               tmp = "Update Txn_NewCashReceiptClientDetail set DocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',EntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & " Where CRID=" & Val(mCRID) & " And SM_Prefix='" & Trim(mSMPrefix) & "' And Amount > 0  And "
               'tmp = tmp & " (ClientIDRow in (Select ClientIDRow from Mst_Client Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' or RNAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') or DPID in (Select DepositoryID from Mst_Depository Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "'))"
               tmp = tmp & " (ClientID in (Select ClientID From (Select ClientID from Mst_Client Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' or RNAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' Union All Select a.ClientID    from Mst_ClientStateWiseGST a Where  (a.GSTNavCode is not null and a.GSTNavCode <> '' And a.GSTNavCode = '" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') Or (a.RGSTNavCode is not null and a.RGSTNavCode <> '' And a.RGSTNavCode = '" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' )) A) or DPID in (Select Distinct DepositoryID From(Select DepositoryID from Mst_Depository Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' Union All Select DepositoryID from Mst_DepositoryStateWiseGST Where GSTNavCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') A))"
               tmp = tmp & " And AdvNAVCode  = '" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' And CRID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "' " '--And Isnull(DocNo,'') ='' "
               Call TmpTable
               
               RsTxn_CashRepeiptsForNAVISION!EntryNoTransfer = "T"
               RsTxn_CashRepeiptsForNAVISION.Update
               
            End If
                
            RsTxn_CashRepeiptsForNAVISION.MoveNext
            ProgressBar1.Value = mR
          Next mR
       End If
    
    
       Call TableTxn_CashRepeiptsForNAVISION(" Where Concurrency >= '2020-01-25' And InvoiceType = 'CR' And AccountType = 'Customer' And ISNULL(DataTransfer,'') = 'T' And ISNULL(ReverseInvoice,'') = ''  And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(CRADJEntryNo,0) <> 0) And ISNULL(PINo,'') <> '' And ISNULL(EntryNoTransfer,'') = '' ")
       If RsTxn_CashRepeiptsForNAVISION.RecordCount > 0 Then
          ProgressBar1.Max = RsTxn_CashRepeiptsForNAVISION.RecordCount
          For mR = 1 To RsTxn_CashRepeiptsForNAVISION.RecordCount
                
            mstr = Split(RsTxn_CashRepeiptsForNAVISION!PINo, "-")
            If UBound(mstr) = 1 Then
               mCRID = mstr(0)
               mCRID = Replace(UCase(mCRID), "CR", "")
               mSMPrefix = mstr(1)
               
               tmp = "Update Txn_NewCashReceiptTIDetail Set DocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',EntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & ""
               tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_NewCashReceiptTIDetail Where NewCRID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') And Isnull(CRAAID,0)=0 And NewCRID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "' And Isnull(DocNo,'') ='' "
               Call TmpTable
              
               '--for Advace
               tmp = "Update Txn_NewCashReceiptClientDetail set DocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',EntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & " Where NewCRID=" & Val(mCRID) & " And SM_Prefix='" & Trim(mSMPrefix) & "' And Amount > 0  And "
               'tmp = tmp & " (ClientIDRow in (Select ClientIDRow from Mst_Client Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' or RNAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') or DPID in (Select DepositoryID from Mst_Depository Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "'))"
               tmp = tmp & " (ClientID in (Select ClientID From (Select ClientID from Mst_Client Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' or RNAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' Union All Select a.ClientID    from Mst_ClientStateWiseGST a Where  (a.GSTNavCode is not null and a.GSTNavCode <> '' And a.GSTNavCode = '" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') Or (a.RGSTNavCode is not null and a.RGSTNavCode <> '' And a.RGSTNavCode = '" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' )) A) or DPID in (Select Distinct DepositoryID From(Select DepositoryID from Mst_Depository Where NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' Union All Select DepositoryID from Mst_DepositoryStateWiseGST Where GSTNavCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') A))"
               tmp = tmp & " And AdvNAVCode  = '" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "' And NewCRID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "' " '--And Isnull(DocNo,'') ='' "
               Call TmpTable
               
               RsTxn_CashRepeiptsForNAVISION!EntryNoTransfer = "T"
               RsTxn_CashRepeiptsForNAVISION.Update
               
            End If
                
            RsTxn_CashRepeiptsForNAVISION.MoveNext
            ProgressBar1.Value = mR
          Next mR
       End If
    
    
    
    End If
    ''-- TDS Adj Entry No
    If ChkTDSAdjEntryNo.Value = vbChecked Then
       Call TableTxn_CashRepeiptsForNAVISION(" Where Concurrency <= '2020-01-25' And InvoiceType = 'CRADJ' And AccountType = 'Customer' And ISNULL(DataTransfer,'') = 'T' And ISNULL(ReverseInvoice,'') = '' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(CRADJEntryNo,0) <> 0) And ISNULL(PINo,'') <> ''  ")
       If RsTxn_CashRepeiptsForNAVISION.RecordCount > 0 Then
           ProgressBar1.Max = RsTxn_CashRepeiptsForNAVISION.RecordCount
           For mR = 1 To RsTxn_CashRepeiptsForNAVISION.RecordCount
                
            mstr = Split(RsTxn_CashRepeiptsForNAVISION!PINo, "-")
            If UBound(mstr) = 1 Then
               mCRID = mstr(0)
               mCRID = Replace(UCase(mCRID), "TDS", "")
               mSMPrefix = mstr(1)
               
               If LCase(RsTxn_CashRepeiptsForNAVISION!DocumentType) = "payment" Then
                  tmp = "Update Txn_CRTDSAdjustmentDetail Set TDSDocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',TDSEntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & ""
                  tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_CRTDSAdjustmentDetail Where TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') And RecTDS>0 And TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "'"
                  Call TmpTable
                  
               ElseIf LCase(RsTxn_CashRepeiptsForNAVISION!DocumentType) = "credit memo" Then
                  If InStr(1, LCase(RsTxn_CashRepeiptsForNAVISION!Narration), "other deduction") > 0 Then
                     tmp = "Update Txn_CRTDSAdjustmentDetail Set OtherDedDocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',OtherDedEntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & ""
                     tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_CRTDSAdjustmentDetail Where TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') And RecOtherDed>0 And TDSAID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "'"
                     Call TmpTable
                  End If
               End If
            
               RsTxn_CashRepeiptsForNAVISION!EntryNoTransfer = "T"
               RsTxn_CashRepeiptsForNAVISION.Update
               
            End If
                
            RsTxn_CashRepeiptsForNAVISION.MoveNext
            ProgressBar1.Value = mR
           Next mR
       End If
    
    
       Call TableTxn_CashRepeiptsForNAVISION(" Where Concurrency >= '2020-01-25' And InvoiceType = 'CRADJ' And AccountType = 'Customer' And ISNULL(DataTransfer,'') = 'T' And ISNULL(ReverseInvoice,'') = '' And ISNULL(EntryNoTransfer,'') = '' And (ISNULL(PostInvoiceNo,'') <> '' And Isnull(CRADJEntryNo,0) <> 0) And ISNULL(PINo,'') <> ''  ")
       If RsTxn_CashRepeiptsForNAVISION.RecordCount > 0 Then
           ProgressBar1.Max = RsTxn_CashRepeiptsForNAVISION.RecordCount
           For mR = 1 To RsTxn_CashRepeiptsForNAVISION.RecordCount
                
            mstr = Split(RsTxn_CashRepeiptsForNAVISION!PINo, "-")
            If UBound(mstr) = 1 Then
               mCRID = mstr(0)
               mCRID = Replace(UCase(mCRID), "TDS", "")
               mSMPrefix = mstr(1)
               
               If LCase(RsTxn_CashRepeiptsForNAVISION!DocumentType) = "payment" Then
                  tmp = "Update Txn_CRTDSAdjustmentDetail Set TDSDocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',TDSEntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & ""
                  tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_CRTDSAdjustmentDetail Where CRTDSAdjustmentID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') And RecTDS>0 And CRTDSAdjustmentID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "'"
                  Call TmpTable
                  
               ElseIf LCase(RsTxn_CashRepeiptsForNAVISION!DocumentType) = "credit memo" Then
                  If InStr(1, LCase(RsTxn_CashRepeiptsForNAVISION!Narration), "other deduction") > 0 Then
                     tmp = "Update Txn_CRTDSAdjustmentDetail Set OtherDedDocNo='" & Trim(RsTxn_CashRepeiptsForNAVISION!PostInvoiceNo) & "',OtherDedEntryNo=" & Val(RsTxn_CashRepeiptsForNAVISION!CRADJEntryNo) & ""
                     tmp = tmp & " Where TIID in( Select TIID from  Txn_TIDetails Where TIID in (select TIID from Txn_CRTDSAdjustmentDetail Where CRTDSAdjustmentID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "') and NAVCode='" & Trim(RsTxn_CashRepeiptsForNAVISION!AccountNo) & "') And RecOtherDed>0 And CRTDSAdjustmentID=" & Val(mCRID) & " and SM_Prefix='" & Trim(mSMPrefix) & "'"
                     Call TmpTable
                  End If
               End If
            
               RsTxn_CashRepeiptsForNAVISION!EntryNoTransfer = "T"
               RsTxn_CashRepeiptsForNAVISION.Update
               
            End If
                
            RsTxn_CashRepeiptsForNAVISION.MoveNext
            ProgressBar1.Value = mR
           Next mR
       End If
    
    
    
    End If
    
    MsgBox "Done."
    Call CmdReset_Click
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Call CmdReset_Click
End Sub

Private Sub CmdReset_Click()
    CmdUpdate.Enabled = True
    ChkReservationInvoice.Value = vbUnchecked
    ChkCommtrackInvoice.Value = vbUnchecked
    Chk7DaysInvoice.Value = vbUnchecked
    ChkNNNRInvoice.Value = vbUnchecked
    ChkNNNRYearlyInvoice.Value = vbUnchecked
    ChkPWRInvoice.Value = vbUnchecked
    ChkRejectedInvoice.Value = vbUnchecked
    ChkRRNInvoice.Value = vbUnchecked
    ChkAdHocInvoice.Value = vbUnchecked
    ChkCashReceipt.Value = vbUnchecked
    ChkTDSAdjEntryNo.Value = vbUnchecked
    ProgressBar1.Value = 0
    Call Grid_Head
    LblFile.Caption = ""
'    CmdUpdate.Enabled = False
    CmdImport.Enabled = False
    CmdGetFile.Enabled = True
    
    
End Sub

Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 31 '27
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Date"
    .TextMatrix(1, 1) = "Invoice"
    .TextMatrix(1, 2) = "Customer"
    .TextMatrix(1, 3) = "NavCode"
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "G/L Account"
    .TextMatrix(1, 6) = "GL A/C NAVCode "
    .TextMatrix(1, 12) = "Narration"
    .TextMatrix(1, 13) = "Location NAV Code"
    .TextMatrix(1, 14) = "CMP"
    .TextMatrix(1, 18) = "PI No"
    .TextMatrix(1, 19) = "ST Code"
    .TextMatrix(1, 20) = "Agree/Non Agree"
    .TextMatrix(1, 21) = "GST Code"
    .TextMatrix(1, 22) = "HSN No"
    .TextMatrix(1, 23) = "Invoice Type"
    .TextMatrix(1, 24) = "Number Series Flag"
    .TextMatrix(1, 25) = "Invoice No"
    .TextMatrix(1, 26) = "Invoice Date"
    .TextMatrix(1, 27) = "ST Amt"
    .TextMatrix(1, 28) = "Invoice Amt"
    .TextMatrix(1, 29) = "PostInvoice No"
    .TextMatrix(1, 30) = "Entry No"

    .ColWidth(0) = 1200
    .ColWidth(1) = 900
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 2000
    .ColWidth(13) = 1600
    .ColWidth(14) = 1400
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 1400
    .ColWidth(19) = 0
    .ColWidth(20) = 1200
    .ColWidth(21) = 900
    .ColWidth(22) = 1200
    .ColWidth(23) = 1200
    .ColWidth(24) = 1200
    .ColWidth(25) = 1200
    .ColWidth(26) = 1200
    .ColWidth(21) = 1200
    .ColWidth(22) = 1200
    .ColWidth(23) = 1200
    .ColWidth(24) = 1200
    .ColWidth(25) = 1200
    .ColWidth(26) = 1200
    .ColWidth(27) = 1200
    .ColWidth(28) = 1200
    .ColWidth(29) = 1200
    .ColWidth(30) = 1200


End With
End Sub
''
''Private Sub CmdReset_Click()
''Call Grid_Head
''LblFile.Caption = ""
''CmdUpdate.Enabled = False
''CmdImport.Enabled = False
''CmdGetFile.Enabled = True
''End Sub
''
''Private Sub CmdUpdate_Click()
''
''MsgBox "Please Wait till Next Message."
''Call TableTxn_ReservationInvoice
''Call TableTxn_TIDetails '(" Where PIFlag='Reservation'")
''
''For I = 2 To MSHFlexGrid3.Rows - 1
''    If MSHFlexGrid3.TextMatrix(I, 20) = "" Or MSHFlexGrid3.TextMatrix(I, 21) = "" Then Exit For
''    If RsTxn_ReservationInvoice.RecordCount > 0 Then
''       RsTxn_ReservationInvoice.MoveFirst
''    End If
''    mInvoice = Replace(MSHFlexGrid3.TextMatrix(I, 18), "RI", "")
''    RsTxn_ReservationInvoice.Find "InvoiceNo= " & Val(mInvoice) & " "  '" & Val(MSHFlexGrid3.TextMatrix(I, 18)) & ""
''    If Not RsTxn_ReservationInvoice.EOF Then
''       If RsTxn_TIDetails.RecordCount > 0 Then
''          RsTxn_TIDetails.MoveFirst
''          RsTxn_TIDetails.Find "EntryNo= " & Val(MSHFlexGrid3.TextMatrix(I, 26)) 'RsTxn_TIDetails.Find "PIID=" & RsTxn_ReservationInvoice!InvoiceNo
''          If RsTxn_TIDetails.EOF = True Then
''             RsTxn_TIDetails.AddNew
''             RsTxn_TIDetails!TIID = getTIDetailMax
''             RsTxn_TIDetails!TINo = Trim(MSHFlexGrid3.TextMatrix(I, 21))
''             RsTxn_TIDetails!TIDate = Format(MSHFlexGrid3.TextMatrix(I, 22), Gen_DatFormat)
''             RsTxn_TIDetails!TIAmount = Val(MSHFlexGrid3.TextMatrix(I, 24))
''             RsTxn_TIDetails!TISTAmount = Val(MSHFlexGrid3.TextMatrix(I, 23))
''             RsTxn_TIDetails!PostInvoiceNo = Trim(MSHFlexGrid3.TextMatrix(I, 25))
''             RsTxn_TIDetails!EntryNo = Val(MSHFlexGrid3.TextMatrix(I, 26))
''             RsTxn_TIDetails!PIID = RsTxn_ReservationInvoice!InvoiceNo
''             RsTxn_TIDetails!LocationID = RsTxn_ReservationInvoice!LocationID
''
''             tmp = " Select TRI.ContractID,MSC.ClientIDRow from Txn_ReservationInvoice TRI"
''             tmp = tmp & " INNER join MST_StorageContract MSC ON TRI.ContractID=MSc.ContractID"
''             tmp = tmp & " Where InvoiceNo = " & RsTxn_ReservationInvoice!InvoiceNo & ""
''
''             Call TmpTable
''
''             RsTxn_TIDetails!ClientIDRow = TmpRs!ClientIDRow
''             RsTxn_TIDetails!NAVCode = Trim(MSHFlexGrid3.TextMatrix(I, 3))
''             RsTxn_TIDetails!ExchangeTypeID = 1 ' reservation income is  non ncdex
''             RsTxn_TIDetails!RecAmount = 0
''             RsTxn_TIDetails!RecTDSAmount = 0
''             RsTxn_TIDetails!PIFlag = "Reservation"
''             RsTxn_TIDetails!SMID = MSHFlexGrid3.TextMatrix(I, 18)
'''             RsTxn_TIDetails!EntryFlag = "System"
''             RsTxn_TIDetails!G_UID = GetGUID
''             RsTxn_TIDetails!CreatedDate = Now
''             RsTxn_TIDetails!CreatedBy = Gen_UserLoginID
''             RsTxn_TIDetails.Update
''
''          End If
''       End If
''    End If
''Next
''MsgBox "Done !!"
''End Sub
''
''Private Sub Form_Load()
''Dim FrmLoadAccess() As Boolean
''FrmLoadAccess = GetFrmLoadAccess(Me.Name)
''If FrmLoadAccess(0) = False Then
''   'Call GButtonLockAD(Me)
''   MsgBox "Access Denied !!!!!!!!!"
''   Frame1.Enabled = False
''   Exit Sub
''End If
''Frame1.Enabled = True
''Call CmdReset_Click
''End Sub
''
Public Sub ReadFiles(file_name As String)
Dim lngFreefile As Long
Dim mRow As Variant
MsgBox "Wait till next message"

Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(file_name)
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
sFileLineNo = 1
mRow = 2

Do While Not objTS.AtEndOfStream
    sFileLine = objTS.ReadLine
    sFileLineNo = sFileLineNo + 1
    ps = Split(sFileLine, ",")
    ' Here total fields are 25 so commas between them are 24 in .CSV file
    ' so we have compared ps with 24(File ReservationInvoice in navision folder)
    If UBound(ps) <> 30 Then 'If UBound(ps) <> 26 Then
       MsgBox "Invalid file format."
       Exit Sub
    End If
    For I = 0 To UBound(ps)
         MSHFlexGrid3.TextMatrix(mRow, I) = Trim(ps(I))
    Next I
    mRow = mRow + 1
    MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
Loop
MsgBox "Import Successfully"
End Sub
Function getTIDetailMax() As Double
Dim Retval As Double
Retval = 0
tmp = "select max(TIID) from Txn_TIDetails"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
End If
Retval = Retval + 1
TmpRs.Close
getTIDetailMax = Retval
End Function
''


