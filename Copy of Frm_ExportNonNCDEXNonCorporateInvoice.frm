VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ExportNonNCDEXNonCorporateInvoice 
   Caption         =   "Export Non NCDEX/Non Corporate Invoice"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15015
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
         ItemData        =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":0000
         Left            =   8160
         List            =   "Frm_ExportNonNCDEXNonCorporateInvoice.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox CmbBatch 
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
         Height          =   360
         Left            =   8160
         TabIndex        =   9
         Top             =   675
         Width           =   2040
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Old"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   255
         TabIndex        =   8
         Top             =   675
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export"
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
         Left            =   11460
         TabIndex        =   6
         Top             =   240
         Width           =   1140
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
         Left            =   12600
         TabIndex        =   5
         Top             =   240
         Width           =   1140
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
         Left            =   13740
         TabIndex        =   4
         Top             =   240
         Width           =   1140
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
         Left            =   10320
         TabIndex        =   3
         Top             =   240
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   2550
         TabIndex        =   13
         Top             =   240
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
         Format          =   72155137
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   4965
         TabIndex        =   14
         Top             =   240
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
         Format          =   72155137
         CurrentDate     =   39884
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
         Left            =   1605
         TabIndex        =   16
         Top             =   315
         Width           =   990
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
         Left            =   4230
         TabIndex        =   15
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   270
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
         Left            =   6840
         TabIndex        =   11
         Top             =   728
         Width           =   1095
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
      Height          =   8025
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   7560
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13335
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "Frm_ExportNonNCDEXNonCorporateInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mGLAccNAVCode, mSTCode As Variant
Dim ExpFlag, MaxID As Variant
Dim mSMonth, mSYear, mSTotalDays As Variant

Private Sub CmbBatch_GotFocus()
CmbInvoiceType.Text = ""
tmp = "select distinct BatchID As BatchID from Txn_ProInvoice Where BatchID > 0  order by BatchID"
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

Private Sub CmbInvoiceType_LostFocus()
If CmbInvoiceType.Text = "" Then Exit Sub
If CmbInvoiceType.Text <> "Non Taxable" And CmbInvoiceType.Text <> "Taxable" Then
   MsgBox "Invalid Selection"
   CmbInvoiceType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdExport_Click()
Dim Path, strFilName, strFilLocation As Variant
Path = App.Path
MaxID = GetMaxBatchID
If Option1.Value = True Then
   strFilName = "\Navision\NavUpload_NonNCDEXNonCorporate" & "_" & CmbInvoiceType.Text & "_" & MaxID & " .csv"
Else
   strFilName = "\Navision\NavUpload_NonNCDEXNonCorporate" & "_" & CmbInvoiceType.Text & "_" & CmbBatch.Text & ".csv"
End If
strFilLocation = Path + strFilName
If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
Open strFilLocation For Output As #1

For I = 2 To MSHFlexGrid3.Rows - 1
    tmp = ""
    For c = 0 To MSHFlexGrid3.Cols - 1
        If tmp = "" Then
           tmp = tmp & MSHFlexGrid3.TextMatrix(I, c)
        Else
           tmp = tmp & "," & MSHFlexGrid3.TextMatrix(I, c)
        End If
    Next
    Print #1, tmp
Next
If Option1.Value = True Then
   
   If CmbInvoiceType.Text = "Non Taxable" Then
      tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0  And ServiceTaxNAVCode = '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
   Else
      tmp = "Update Txn_ProInvoice set BatchID = " & MaxID & ",FAGPostingDate = '" & Date & "' Where BatchID = 0 And ServiceTaxNAVCode <> '' And InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And InvoiceDate <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
   End If
   Call TmpTable
End If

MsgBox "Done !!!"
Close #1

End Sub
Private Function GetMaxBatchID()
tmp = "select Max(BatchID) from Txn_ProInvoice"
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   GetMaxBatchID = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0)) + 1
Else
   GetMaxBatchID = 1
End If
End Function
Private Sub CmdGetReport_Click()
Dim wc As Variant
ExpFlag = ""
If Option1.Value = True Then
   
   If TxtFromDate.Value > TxtToDate.Value Then
      MsgBox "'To Date' must be greater than 'From Date' always!!!"
      TxtToDate.SetFocus
      Exit Sub
   End If
   If CmbInvoiceType.Text = "" Then
      MsgBox "Select Invoice Type."
      CmbInvoiceType.SetFocus
      Exit Sub
   End If
   
   If CmbInvoiceType.Text = "Non Taxable" Then
      wc = " Where  TPI.ServiceTaxNAVCode = '' And TPI.BatchID = 0"
   ElseIf CmbInvoiceType.Text = "Taxable" Then
      wc = " Where  TPI.ServiceTaxNAVCode <> '' And TPI.BatchID = 0"
   End If
   
   If Gen_DBType = "MDB" Then
      wc = wc & "And TPI.InvoiceDate >= #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And TPI.InvoiceDate <= #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# "
   Else
      wc = wc & "And TPI.InvoiceDate >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And TPI.InvoiceDate <= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' "
   End If
Else
   If CmbBatch.Text = "" Then
      MsgBox "Select Batch ID !!!! "
      Exit Sub
   End If
    wc = " Where  TPI.BatchID=" & Val(CmbBatch.Text) & ""
End If

Call Grid_Head


mGLAccNAVCode = GetGLAccNAVCode("NNNC")
'mSTCode = GetGLAccNAVCode("DPSTCode")

'tmp = "Select MD.DPInvoiceID,MD.MonthYear,MD.LocationID,MD.DPNavCode,MD.Amt,"
'tmp = tmp & " MD.STAmt , MD.TotalAmt, MD.PINo, MD.BatchID, ML.NAVCode, MD.Flag,ML.LocationName"
'tmp = tmp & " From Txn_ProInvoice TPI"
'tmp = tmp & " Left join Mst_Location ML on MD.LocationID=ML.LocationID "
'tmp = tmp & Wc

tmp = "Select TPI.InvoiceNo,TPI.SM_Prefix,ML.LocationName,ML.NAVCode 'LocatioNAVCode',MC.ClientName,MC.NAVCode 'ClientNAVCode',TPI.Amount,TPI.BatchID,TPI.ServiceTaxNAVCode,"
tmp = tmp & " TPI.InvoiceDate , TPI.Discount, TPI.Premium, TPI.FAGPostingDate"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " Inner join Mst_Location ML on TPI.LocationID=ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MC On TPI.ClientIDRow = MC.ClientIDRow"
tmp = tmp & wc & " Order By TPI.InvoiceDate"


Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
   If TmpRs!ServiceTaxNAVCode <> "" Then
      CmbInvoiceType.Text = "Taxable"
   Else
      CmbInvoiceType.Text = "Non Taxable"
   End If
   
   Frame5.Caption = CmbInvoiceType.Text
   
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid3.TextMatrix(I + 1, 0) = Format(TmpRs!InvoiceDate, "dd/mm/yyyy")
       MSHFlexGrid3.TextMatrix(I + 1, 1) = "Invoice"
       MSHFlexGrid3.TextMatrix(I + 1, 2) = "Customer"
       MSHFlexGrid3.TextMatrix(I + 1, 3) = IIf(IsNull(TmpRs!ClientNAVCode), "", TmpRs!ClientNAVCode)
       
       'If TmpRs!Discount <> 0 And IsNull(TmpRs!Discount) = Flase Then
       MSHFlexGrid3.TextMatrix(I + 1, 4) = TmpRs!Amount + TmpRs!Premium - TmpRs!Discount
       'ElseIf TmpRs!Premium <> 0 And IsNull(TmpRs!Premium) = Flase Then
       '   MSHFlexGrid3.TextMatrix(i + 1, 4) = TmpRs!Amount + TmpRs!Premium
       'Else
       '   MSHFlexGrid3.TextMatrix(i + 1, 4) = TmpRs!Amount
       'End If
       MSHFlexGrid3.TextMatrix(I + 1, 5) = "G/L Account"
       MSHFlexGrid3.TextMatrix(I + 1, 6) = mGLAccNAVCode
       
       If CmbInvoiceType.Text = "Non Taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (Without ST) having PI No :- NN" & TmpRs!InvoiceNo & " "
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 19) = TmpRs!ServiceTaxNAVCode
          MSHFlexGrid3.TextMatrix(I + 1, 12) = "Being Invoice for " & TmpRs!LocationName & " for the month of " & MonthName(Month(TmpRs!InvoiceDate)) & " " & Year(TmpRs!InvoiceDate) & " (With ST) having PI No :- NN" & TmpRs!InvoiceNo & " "
       End If
       MSHFlexGrid3.TextMatrix(I + 1, 18) = IIf(IsNull(TmpRs!InvoiceNo), "", "NN" & TmpRs!InvoiceNo)
       
       MSHFlexGrid3.TextMatrix(I + 1, 13) = TmpRs!LocatioNAVCode
       MSHFlexGrid3.TextMatrix(I + 1, 14) = "CMP"
       If LCase(CmbInvoiceType) = "non taxable" Then
          MSHFlexGrid3.TextMatrix(I + 1, 20) = "AGRI"
       Else
          MSHFlexGrid3.TextMatrix(I + 1, 20) = "NAGRI"
       End If
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(I + 1, 4))
       
       If IsNull(TmpRs!ClientNAVCode) = True Or TmpRs!ClientNAVCode = "" Then
          ExpFlag = "N"
       End If
       If IsNull(TmpRs!LocatioNAVCode) = True Or TmpRs!LocatioNAVCode = "" Then
          ExpFlag = "N"
       End If
'       If mGLAccNAVCode = "" Then
'          ExpFlag = "N"
'       End If
       TmpRs.MoveNext
   Next
   If ExpFlag = "N" Then
      CmdExport.Enabled = False
   Else
      CmdExport.Enabled = True
   End If
Else
   CmdExport.Enabled = False
End If

End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmbBatch.Text = ""
CmbInvoiceType.Text = ""
CmdExport.Enabled = False
Option1.Value = True
Call Option1_Click
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
Call CmdReset_Click

End Sub
Public Sub Grid_Head()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 21
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Date"
    .TextMatrix(1, 1) = "Invoice"
    .TextMatrix(1, 2) = "Customer"
    .TextMatrix(1, 3) = "DPNavCode"
    .TextMatrix(1, 4) = "Amount"
    .TextMatrix(1, 5) = "G/L Account"
    .TextMatrix(1, 6) = "GL A/C NAVCode "
    .TextMatrix(1, 12) = "Narration"
    .TextMatrix(1, 13) = "Location NAV Code"
    .TextMatrix(1, 14) = "CMP"
    .TextMatrix(1, 18) = "PI No"
    .TextMatrix(1, 19) = "ST Code"
    .TextMatrix(1, 20) = "Agri/Non Agri"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
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
    .ColWidth(19) = 1000
    .ColWidth(20) = 1400
End With

End Sub
Private Sub Option1_Click()
Call Option2_Click
End Sub

Private Sub Option2_Click()
CmdExport.Enabled = False
CmbInvoiceType.Text = ""
If Option2.Value = True Then
   CmbBatch.Enabled = True
   CmbInvoiceType.Enabled = False
   TxtFromDate.Enabled = False
   TxtToDate.Enabled = False
Else
   CmbBatch.Text = ""
   CmbBatch.Enabled = False
   CmbInvoiceType.Enabled = True
   TxtFromDate.Enabled = True
   TxtToDate.Enabled = True
   TxtFromDate.Value = Date
   TxtToDate.Value = Date
End If
End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
End Sub

Private Sub TxtToDate_LostFocus()
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)

mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function


