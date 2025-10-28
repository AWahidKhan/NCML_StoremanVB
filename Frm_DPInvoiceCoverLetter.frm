VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_DPInvoiceCoverLetter 
   Caption         =   "DP Invoice Cover Letter"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
      Height          =   8715
      Left            =   120
      TabIndex        =   4
      Top             =   750
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
         TabIndex        =   6
         Top             =   315
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7920
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13970
         _Version        =   393216
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
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   6780
         TabIndex        =   9
         Top             =   240
         Width           =   1260
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
         Left            =   5520
         TabIndex        =   8
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
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
         Left            =   4260
         TabIndex        =   7
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1260
      End
      Begin VB.ComboBox CmbBatch 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Batch No"
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
         Left            =   360
         TabIndex        =   3
         Top             =   270
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Frm_DPInvoiceCoverLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExpFlag, MaxID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mMonth As Variant
Dim mBatchID As Variant

Private Sub ChkCheckAll_Click()
If MSHFlexGrid3.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       MSHFlexGrid3.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid3.Rows - 2
       MSHFlexGrid3.row = I
       MSHFlexGrid3.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbBatch_GotFocus()
tmp = "select distinct BatchID As BatchID from Mst_DPInvoiceUpload Where BatchID > 0  order by BatchID"
Call Tmp1Table
CmbBatch.Clear
If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       CmbBatch.AddItem TmpRs1.Fields(0)
       MaxID = TmpRs1.Fields(0) + 1
       TmpRs1.MoveNext
   Next
Else
   MaxID = 1
End If
End Sub
Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 7
    .FixedCols = 0
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "Date"
    .TextMatrix(0, 2) = "DPNavCode"
    .TextMatrix(0, 3) = "DP Name"
    .TextMatrix(0, 4) = "Amount"
    .TextMatrix(0, 5) = "ST Amount"
    .TextMatrix(0, 6) = "Total Amount"

    .ColWidth(1) = 1400
    .ColWidth(2) = 1500
    .ColWidth(3) = 4800
    .ColWidth(4) = 1700
    .ColWidth(5) = 1700
    .ColWidth(6) = 1700
    
    
End With

End Sub
Private Sub CmbBatch_LostFocus()
If CmbBatch.Text = "" Then
   Exit Sub
End If
TmpRs1.MoveFirst
TmpRs1.Find "BatchID = " & Val(CmbBatch.Text) & ""
If TmpRs1.EOF Then
   MsgBox "Invalid selection "
   CmbBatch.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()
CmbBatch.Text = ""
Call Grid_Head
mBatchID = ""
ChkCheckAll.Value = vbUnchecked
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetReport_Click()

If CmbBatch.Text = "" Then
   MsgBox "Select Batch ID !!!! "
   Exit Sub
End If

Call Grid_Head

tmp = "Select MD.MonthYear,MD.DPNavCode,Sum(MD.Amt) as Amt,"
tmp = tmp & " Sum(MD.STAmt)as STAmt , Sum(MD.TotalAmt) As TotalAmt,MD.BatchID,MD.Flag"
tmp = tmp & " From Mst_DPInvoiceUpload MD"
tmp = tmp & " Left join Mst_Location ML on MD.LocationID=ML.LocationID"
tmp = tmp & " Where MD.BatchID = " & Val(CmbBatch.Text) & ""
tmp = tmp & " Group By MD.MonthYear,MD.BatchID,MD.DPNavCode,MD.Flag"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   If TmpRs!Flag = "T" Then
      Frame5.Caption = "With Service Tax"
   Else
      Frame5.Caption = "Without Service Tax"
   End If
   Frame5.Refresh
   mMonth = MonthName(Month(TmpRs!MonthYear)) & " " & Year(TmpRs!MonthYear)
   mBatchID = CmbBatch.Text
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid3.TextMatrix(I, 1) = TmpRs!MonthYear
       MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
       MSHFlexGrid3.TextMatrix(I, 3) = GetDepositoryNameFromNavCode(TmpRs!DPNAVCode)  'TmpRs!MonthYear
       MSHFlexGrid3.TextMatrix(I, 4) = TmpRs!Amt
       MSHFlexGrid3.TextMatrix(I, 5) = TmpRs!STAmt
       MSHFlexGrid3.TextMatrix(I, 6) = TmpRs!TotalAmt
       MSHFlexGrid3.row = I
       MSHFlexGrid3.Col = 0
       MSHFlexGrid3.CellFontName = "Wingdings"
       MSHFlexGrid3.CellFontSize = 12
       MSHFlexGrid3.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid3.Text = strUnChecked
       MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
       TmpRs.MoveNext
   Next
End If
End Sub

Private Sub CmdPrint_Click()
If CheckDPSelect = False Then
   MsgBox "Please select Atleast one DP"
   Exit Sub
End If
MsgBox ("Wait till next message")
For I = 1 To MSHFlexGrid3.Rows - 1
    If MSHFlexGrid3.TextMatrix(I, 0) = strChecked Then
       Gen_mTimeStamp = GetTimeStamp
       Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "DPInvoiceCoverLetter_" & MSHFlexGrid3.TextMatrix(I, 2) & "_" & CmbBatch.Text & ".xls")
       Call Xls_Detail_Rpt(MSHFlexGrid3.TextMatrix(I, 3), MSHFlexGrid3.TextMatrix(I, 2))
    End If
Next
MsgBox ("Excel files created in 'Excel' folder")
End Sub
Public Sub Xls_Detail_Rpt(DP_ As String, NavCode_ As String)


Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow, mRow1 As Integer
Dim mCol As Integer
Dim c As Variant
Set oXL = New Excel.Application
Dim Pat As String
Dim mAmt, mSTAmt, mTotAmt As Double
Pat = App.Path

Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "DPInvoiceCoverLetter_" & NavCode_ & "_" & CmbBatch.Text & ".xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows

oWS.Shapes.AddPicture(App.Path & "\logo.gif", msoTrue, msoTrue, 10, 10, 40, 40).Select

tmp = "B" & 1 & ":f" & 1
oWS.Range(tmp).Merge
oWS.Cells(1, 2) = "National Collateral Management Services Limited"
oWS.Cells(1, 2).HorizontalAlignment = -4108
oWS.Range(tmp).Font.Bold = True

tmp = "B" & 2 & ":f" & 2
oWS.Range(tmp).Merge
oWS.Cells(2, 2) = "Lodha Supremus, Unit No 505 to 509 , 5th Floor,"
oWS.Cells(2, 2).HorizontalAlignment = -4108

tmp = "B" & 3 & ":f" & 3
oWS.Range(tmp).Merge
oWS.Cells(3, 2) = "Kanjurmarg (East), Mumbai-400042"
oWS.Cells(3, 2).HorizontalAlignment = -4108

tmp = "B" & 4 & ":f" & 4
oWS.Range(tmp).Merge
oWS.Cells(4, 2) = "Tel.No-022 40419191, Fax. No-022 40419193"
oWS.Cells(4, 2).HorizontalAlignment = -4108

tmp = "a" & 7 & ":f" & 7
oWS.Range(tmp).Merge
oWS.Cells(7, 1) = "No :"
tmp = "a" & 8 & ":f" & 8
oWS.Range(tmp).Merge
oWS.Cells(8, 1) = "Date : " & Format(Date, "dd/MMM/yyyy")
oWS.Cells(9, 1) = "To :"

tmp = "a" & 10 & ":f" & 10
oWS.Range(tmp).Merge
oWS.Cells(10, 1) = "   " & DP_
oWS.Cells(10, 1).Font.Bold = True
tmp = "a" & 11 & ":f" & 11
oWS.Range(tmp).Merge
oWS.Cells(11, 1) = ""  'Address
oWS.Cells(12, 1) = "Dear Sir,"
tmp = "a" & 14 & ":f" & 16
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(14, 1) = "Please find enclosed Tax Invoices for the month of " & mMonth & " towards stock held in" & _
                  " Demat format in your esteemed organization by various clients. The details of the clients " & _
                  " and the amount outstanding are already sent to you by NCDEX directly."

oWS.Cells(18, 1) = "Invoice No"
oWS.Cells(18, 2) = "Invoice Date"
oWS.Cells(18, 3) = "Location"
oWS.Cells(18, 4) = "Amount"
oWS.Cells(18, 5) = "Service Tax"
oWS.Cells(18, 6) = "Amount (Incl.of Service Tax)"
tmp = "a" & 18 & ":f" & 18
oWS.Range(tmp).Font.Bold = True

tmp = "Select InvoiceNo,InvoiceDate,LocationName,Amt,STAmt,TotalAmt "
tmp = tmp & " From Mst_DPInvoiceUpload inner Join Mst_Location on Mst_Location.LocationID=Mst_DPInvoiceUpload.LocationID "
tmp = tmp & " Where DPNavCode='" & NavCode_ & " ' And BatchID=" & Val(mBatchID) & ""

Call TmpTable

mRow = 19

mAmt = 0
mSTAmt = 0
mTotAmt = 0
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       c = 1
       For Col = 0 To TmpRs.Fields.Count - 1
           If InStr(TmpRs.Fields(Col), "/") > 0 Then
              If IsDate(TmpRs.Fields(Col)) Then
                 oWS.Cells(mRow, c) = Format(TmpRs.Fields(Col), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
              End If
           Else
              If Col = 3 Then
                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
                 mAmt = mAmt + Val(TmpRs.Fields(Col))
              ElseIf Col = 4 Then
                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
                 mSTAmt = mSTAmt + Val(TmpRs.Fields(Col))
              ElseIf Col = 5 Then
                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
                 mTotAmt = mTotAmt + Val(TmpRs.Fields(Col))
              Else
                 oWS.Cells(mRow, c) = TmpRs.Fields(Col)
              End If
           End If
           c = c + 1
       Next
       TmpRs.MoveNext
       mRow = mRow + 1
   Next
End If
tmp = "a" & 18 & ":f" & mRow
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

tmp = "a" & mRow & ":c" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Font.Bold = True
oWS.Cells(mRow, 1) = "Total"
oWS.Cells(mRow, 4) = Format(mAmt, "#####0.00")
oWS.Cells(mRow, 5) = Format(mSTAmt, "#####0.00")
oWS.Cells(mRow, 6) = Format(mTotAmt, "#####0.00")

tmp = "d" & mRow & ":f" & mRow
oWS.Range(tmp).Font.Bold = True

tmp = "a" & mRow + 2 & ":f" & mRow + 3
oWS.Range(tmp).Merge
oWS.Range(tmp).WrapText = True
oWS.Cells(mRow + 2, 1) = "We request you to kindly arrange for remittance of the Invoice amount mentioned" & _
                         " above within 15 days of receipt of this invoice."

tmp = "a" & mRow + 7 & ":f" & mRow + 7
oWS.Range(tmp).Merge
oWS.Cells(mRow + 7, 1) = "For National Collateral Management Services Ltd."

tmp = "a" & mRow + 11 & ":f" & mRow + 11
oWS.Range(tmp).Merge
oWS.Cells(mRow + 11, 1) = "(Authorised Signatory)"

If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If

oWB.Save
oXL.Quit

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame5.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame5.Enabled = True
Call Grid_Head
End Sub
Private Function CheckDPSelect() As Boolean
Dim Retval As Boolean
Retval = False
For I = 0 To MSHFlexGrid3.Rows - 2
    If MSHFlexGrid3.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckDPSelect = Retval
End Function
Private Sub MSHFlexGrid3_Click()
With MSHFlexGrid3
    If .TextMatrix(.row, 1) = "" Then Exit Sub
    If .Col = 0 Then
        If .TextMatrix(.row, 0) = strUnChecked Then
           .TextMatrix(.row, 0) = strChecked
        Else
           .TextMatrix(.row, 0) = strUnChecked
        End If
    End If
End With
End Sub
