VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_NavisionInvoiceCrossCheck 
   Caption         =   "Navision Invoice Cross Checking"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   200
      Left            =   120
      TabIndex        =   2
      Top             =   9600
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   8535
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8190
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   14446
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         Appearance      =   0
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
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.CheckBox ChkRed 
         Caption         =   "Invoice Not Available"
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
         Left            =   4680
         TabIndex        =   12
         Top             =   293
         Width           =   1575
      End
      Begin VB.CheckBox ChkGreen 
         Caption         =   "Invoice Available"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   293
         Width           =   1335
      End
      Begin VB.CommandButton CmdProcess 
         Caption         =   "Verify Invoice"
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
         Left            =   8565
         TabIndex        =   10
         Top             =   280
         Width           =   1545
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
         Height          =   400
         Left            =   10125
         TabIndex        =   5
         Top             =   280
         Width           =   1245
      End
      Begin VB.CommandButton CmdGetData 
         Caption         =   "Search Invoice"
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
         Left            =   6720
         TabIndex        =   4
         Top             =   280
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   3932163
         CurrentDate     =   42489
      End
      Begin VB.Label Label1 
         Caption         =   "From Month"
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
         Top             =   340
         Width           =   1095
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
         Left            =   11475
         TabIndex        =   7
         Top             =   420
         Width           =   2355
      End
      Begin VB.Label Label22 
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
         Left            =   13875
         TabIndex        =   6
         Top             =   420
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmTxn_NavisionInvoiceCrossCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFolderSubPath As Variant
Dim mInvoicePDFFolderpath As Variant
Dim mFilePath As Variant
Dim mMonthInterVal As Integer

Dim oFs As New FileSystemObject
Dim oFs1 As New FileSystemObject
Dim oFolder As Folder
Dim oFile As File


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_NavisionInvoiceCrossCheck.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow, mCol As Double
Dim mR, mC As Double
Dim Pat As String

MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_NavisionInvoiceCrossCheck.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
oWS.ClearArrows
mRow = 1
mR = 0
With MSHFlexGrid1
''-- Header
For mC = 0 To .Cols - 1
    oWS.Cells(mRow, mC + 1) = .TextMatrix(mR, mC)
Next mC
''-- Body
mRow = 1
mR = 0
For mR = 1 To .Rows - 1
    If (ChkGreen.Value = vbChecked And ChkRed.Value = vbChecked) Or (ChkGreen.Value = vbUnchecked And ChkRed.Value = vbUnchecked) Then
        mRow = mRow + 1
        For mC = 0 To .Cols - 1
            If InStr(.TextMatrix(mR, mC), "/") > 0 Then
               If IsDate(.TextMatrix(mR, mC)) Then
                  oWS.Cells(mRow, mC + 1) = Format(.TextMatrix(mR, mC), "mm/dd/yyyy")
               Else
                  oWS.Cells(mRow, mC + 1) = .TextMatrix(mR, mC)
               End If
            Else
               oWS.Cells(mRow, mC + 1) = .TextMatrix(mR, mC)
            End If
        Next mC
        .row = mR
        .Col = mC - 1
        If .CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, mC), oWS.Cells(mRow, mC)).Interior.Color = vbRed
        ElseIf .CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, mC), oWS.Cells(mRow, mC)).Interior.Color = vbGreen
        ElseIf .CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, mC), oWS.Cells(mRow, mC)).Interior.Color = vbYellow
        End If
    ElseIf (ChkGreen.Value = vbChecked And ChkRed.Value = vbUnchecked) Then
        If LCase(.TextMatrix(mR, 8)) = "invoice available" Then
            mRow = mRow + 1
            For mC = 0 To .Cols - 1
                If InStr(.TextMatrix(mR, mC), "/") > 0 Then
                   If IsDate(.TextMatrix(mR, mC)) Then
                      oWS.Cells(mRow, mC + 1) = Format(.TextMatrix(mR, mC), "mm/dd/yyyy")
                   Else
                      oWS.Cells(mRow, mC + 1) = .TextMatrix(mR, mC)
                   End If
                Else
                   oWS.Cells(mRow, mC + 1) = .TextMatrix(mR, mC)
                End If
            Next mC
            .row = mR
            .Col = mC - 1
            If .CellBackColor = vbGreen Then
               oWS.Range(oWS.Cells(mRow, mC), oWS.Cells(mRow, mC)).Interior.Color = vbGreen
            End If
        End If
    ElseIf (ChkGreen.Value = vbUnchecked And ChkRed.Value = vbChecked) Then
        If LCase(.TextMatrix(mR, 8)) = "invoice not available" Then
            mRow = mRow + 1
            For mC = 0 To .Cols - 1
                If InStr(.TextMatrix(mR, mC), "/") > 0 Then
                   If IsDate(.TextMatrix(mR, mC)) Then
                      oWS.Cells(mRow, mC + 1) = Format(.TextMatrix(mR, mC), "mm/dd/yyyy")
                   Else
                      oWS.Cells(mRow, mC + 1) = .TextMatrix(mR, mC)
                   End If
                Else
                   oWS.Cells(mRow, mC + 1) = .TextMatrix(mR, mC)
                End If
            Next mC
            .row = mR
            .Col = mC - 1
            If .CellBackColor = vbRed Then
               oWS.Range(oWS.Cells(mRow, mC), oWS.Cells(mRow, mC)).Interior.Color = vbRed
            End If
        End If
    End If
Next mR
End With
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetData_Click()
Dim mDate As Variant
Dim mFromDate, mToDate As Variant
Dim mR As Variant
CmdProcess.Enabled = False
mDate = CDate(Year(TxtFromDate) & "-" & Month(TxtFromDate) & "-01")
If mDate > Date Then
    MsgBox "Future Month is allowed."
    Exit Sub
End If
mFromDate = Format(mDate, "YYYY-MM-DD")
mToDate = Format(DateAdd("m", mMonthInterVal, mDate), "YYYY-MM-DD")
Call Grid_Head

tmp = " SELECT TID.TIID,TID.TINO,TID.TIDATE,TID.TIAmount,TID.PIID,TID.EntryNo,TID.Concurrency,Case Isnull(PIFlag,'') "
tmp = tmp & " When 'ADHOC' Then (Select CONVERT(Varchar,TI.BatchID) From Txn_TempInvoice TI Where TI.PINo = TID.PIID) "
tmp = tmp & " When 'Commtrack' Then (Select CONVERT(Varchar,MCI.BatchID) From Txn_MonthlyCommtrackInvoice MCI Where MCI.PINo = TID.PIID) "
tmp = tmp & " When 'DP' Then (Select CONVERT(Varchar,DPI.BatchID) From Mst_DPInvoiceUpload DPI Where DPI.PINo = TID.PIID) "
tmp = tmp & " When 'NNNR' Then (Select CONVERT(Varchar,PI.BatchID) From Txn_ProInvoice PI Where PI.InvoiceNo = TID.PIID) "
tmp = tmp & " When 'NNNRY' Then (Select CONVERT(Varchar,CDTF.BatchID) From Txn_CDTFBOE CDTF Where CDTF.BOEID = TID.PIID) "
tmp = tmp & " When 'Reservation' Then (Select CONVERT(Varchar,RI.BatchID) From Txn_ReservationInvoice RI Where RI.InvoiceNo = TID.PIID) "
tmp = tmp & " When 'RJ' Then (Select CONVERT(Varchar,CDTF.BatchID) From Txn_CDTFBOE CDTF Where CDTF.BOEID = TID.PIID) "
tmp = tmp & " When 'RRN' Then (Select CONVERT(Varchar,CDTF.BatchID) From Txn_CDTFBOE CDTF Where CDTF.BOEID = TID.PIID) "
tmp = tmp & " When 'SD' Then (Select CONVERT(Varchar,CDTF.BatchID) From Txn_CDTFBOE CDTF Where CDTF.BOEID = TID.PIID) "
tmp = tmp & " Else '' End [Batch],Case Isnull(PIFlag,'') When 'ADHOC' Then 'ADHOC Invoice' When 'Commtrack' Then 'Comm Track Invoice' "
tmp = tmp & " When 'DP' Then 'Depository Invoice' When 'NNNR' Then 'NNNR Invoice' When 'NNNRY' Then 'NNNR Yearly Invoice' "
tmp = tmp & " When 'Reservation' Then 'Reservation Invoice' When 'RJ' Then 'Rejected Invoice' When 'RRN' Then 'RRN Invoice' When 'SD' Then 'Seven Day Invoice' "
tmp = tmp & " Else '' End [Type] FROM Txn_TIDetails AS TID "
tmp = tmp & " Where TID.Concurrency BETWEEN '" & mFromDate & "' AND '" & mToDate & "' Order By TID.Concurrency "
Call TmpTable
If TmpRs.RecordCount > 0 Then
    With MSHFlexGrid1
    ProgressBar1.Max = TmpRs.RecordCount
    For mR = 1 To TmpRs.RecordCount
        .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!TIID), "", TmpRs!TIID)
        .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!TINo), "", TmpRs!TINo)
        .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!TIDate), "", Format(TmpRs!TIDate, "DD-MMM-yyyy"))
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!TIAmount), "", TmpRs!TIAmount)
        .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
        .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!Batch), "", TmpRs!Batch)
        .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!PIID), "", TmpRs!PIID)
        .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!Type), "", TmpRs!Type)
        .TextMatrix(mR, 8) = ""
        ProgressBar1.Value = mR
        .Rows = .Rows + 1
        TmpRs.MoveNext
    Next
    End With
    CmdProcess.Enabled = True
End If
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdProcess_Click()
Dim mR As Variant
CmdGetData.Enabled = False
CmdProcess.Enabled = False
CmdExcel.Enabled = False
If oFs.FolderExists(mInvoicePDFFolderpath) = False Then
    MsgBox "Given Invoice Folder Path is not reachable." & vbNewLine & "Contact to you Team."
    Exit Sub
End If
With MSHFlexGrid1
    If .Rows <= 2 Then Exit Sub
    ProgressBar1.Max = .Rows - 2
    For mR = 1 To .Rows - 2
        mFilePath = mInvoicePDFFolderpath & "\" & .TextMatrix(mR, 4) & ".pdf"
        If oFs.FileExists(mFilePath) = True Then
           .TextMatrix(mR, 8) = "Invoice Available"
           Call FlexCellColor(MSHFlexGrid1, CDbl(mR), 8, vbGreen)
        Else
           .TextMatrix(mR, 8) = "Invoice Not Available"
           Call FlexCellColor(MSHFlexGrid1, CDbl(mR), 8, vbRed)
        End If
        ProgressBar1.Value = mR
        .Refresh
    Next mR
    CmdGetData.Enabled = True
    CmdExcel.Enabled = True
End With
End Sub

Private Sub Form_Load()
mFolderSubPath = Gen_ReadPathConfigINI("NAVPATH")
mMonthInterVal = Val(Gen_ReadPathConfigINI("NAVTIME"))
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
'mMonthInterVal = 1
mInvoicePDFFolderpath = App.Path & "\" & mFolderSubPath
Call Grid_Head
TxtFromDate.Value = Date
Label22.Caption = "-"
CmdGetData.Enabled = True
CmdProcess.Enabled = False
End Sub


Private Sub Grid_Head()
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 9
   .Font.Size = 10
   .RowHeight(0) = 600
   .TextMatrix(0, 0) = "TI ID":             .ColWidth(0) = 1000
   .TextMatrix(0, 1) = "TI No":             .ColWidth(1) = 2500
   .TextMatrix(0, 2) = "TI Date":           .ColWidth(2) = 1500
   .TextMatrix(0, 3) = "TI Amount":         .ColWidth(3) = 1500
   .TextMatrix(0, 4) = "Entry No":          .ColWidth(4) = 1200
   .TextMatrix(0, 5) = "Batch ID":          .ColWidth(5) = 1000
   .TextMatrix(0, 6) = "PI No":             .ColWidth(6) = 1000
   .TextMatrix(0, 7) = "Invoice Type":      .ColWidth(7) = 2000
   .TextMatrix(0, 8) = "Status":            .ColWidth(8) = 2000
End With
End Sub

Private Sub FlexCellColor(Grid As MSHFlexGrid, row As Double, Col As Double, ColorCode As Variant)
    With Grid
        .row = row
        .Col = Col
        .CellBackColor = ColorCode
        .Refresh
        Me.Refresh
    End With
End Sub

Private Sub MSHFlexGrid1_Click()
Dim mR As Variant
With MSHFlexGrid1
    If .Rows > 2 Then
        mR = .RowSel
        If LCase(.TextMatrix(mR, 8)) = "invoice available" Then
            mFilePath = mInvoicePDFFolderpath & "\" & .TextMatrix(mR, 4) & ".pdf"
            Shell "Explorer.exe " & mFilePath, vbMaximizedFocus
        Else
            MsgBox "Sorry, Invoice Not Available."
        End If
    End If
End With
End Sub

