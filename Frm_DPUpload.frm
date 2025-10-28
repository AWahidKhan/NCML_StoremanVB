VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frm_DPUpload 
   Caption         =   "Import DP File"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6015
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10610
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   14775
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   14535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   14775
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
         Left            =   2820
         TabIndex        =   8
         Top             =   240
         Width           =   2220
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   9420
         TabIndex        =   7
         Top             =   240
         Width           =   2220
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
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   2220
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
         Left            =   7200
         TabIndex        =   5
         Top             =   240
         Width           =   2220
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
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   2220
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   14775
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   375
      Left            =   12480
      TabIndex        =   10
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "Frm_DPUpload"
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
Dim FName() As Variant, vTemp As Variant
Dim Namef As String

Private Sub CmdExit_Click()
On Error Resume Next
oXL.Quit
Unload Me
End Sub

Private Sub CmdReset_Click()
ListView1.ListItems.Clear
MSHFlexGrid1.Clear
CmdGetFile.Enabled = True
CmdExit.Enabled = True
CmdImport.Enabled = False
CmdSave.Enabled = False
TxtErrorMsg.Text = ""
ProgressBar1.Value = 0
End Sub
Private Sub CmdGetFile_Click()
ListView1.ListItems.Clear
TxtErrorMsg.Text = ""
On Error GoTo errortrap
With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls)|*.xls"
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
LblFile.Caption = Path
Call UploadFile
MsgBox "Click Import To import File!!!"
CmdImport.SetFocus
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub UploadFile()

Status = 0
TxtErrorMsg.Text = ""
If LblFile.Caption <> "" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   
   ' Code For Ceating List of Sheet in Excel Sheet
   Dim i As Integer
       For i = 1 To oWB.Worksheets.Count
           Set oWS = oWB.Worksheets(i)
           If i = 1 Then
              ListView1.ListItems.Add i, , oWS.Name
           End If
       Next
   ListView1.ListItems.Item(1).Checked = True
   If Status = 1 Then
      ListView1.ListItems.Clear
      MsgBox ErrMsg, vbExclamation
   End If
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   CmdExit.Enabled = True
   CmdSave.Enabled = False
   oXL.Quit
Else
   MsgBox "Select File", vbInformation, "Information"
End If
End Sub
Private Sub CmdImport_Click()
Dim mulCheck As Integer

Dim isSelected As Boolean

TxtErrorMsg.Text = ""
isSelected = False
mulCheck = 0
MSHFlexGrid1.Clear
For j = 1 To ListView1.ListItems.Count
    ErrMsg = ListView1.ListItems.Item(j).Text
    If ListView1.ListItems.Item(j).Checked Then
       isSelected = True
       mulCheck = mulCheck + 1
    End If
Next j
If mulCheck > 1 Then
   MsgBox "You Can Select Only One WorkSheet At A Time!!!"
   Exit Sub
End If
If isSelected = True Then
    ErrMsg = ""
    Status = 0
    RowCount = 0
    Set oXL = New Excel.Application
    Set oWB = oXL.Workbooks.Open(LblFile.Caption)
    Call Validation
    If Status = 1 Then
       MsgBox ErrMsg, vbExclamation, "Message"
    Else
       Call Grid_Head
       Call FillData
       vTemp = Split(LblFile.Caption, "\")
       ReDim FName(UBound(vTemp))
       Namef = vTemp(UBound(vTemp))
       CmdGetFile.Enabled = False
       CmdImport.Enabled = False
       CmdSave.Enabled = True
       CmdReset.Enabled = True
    End If
    CmdGetFile.Enabled = True
    CmdSave.SetFocus
Else
    MsgBox "Please Select WorkSheet To Import !...", vbInformation
End If
End Sub
Private Sub Validation()
For j = 1 To ListView1.ListItems.Count
' Check Column Name In Excel Sheet
    If ListView1.ListItems.Item(j).Checked Then
       If LCase(Trim(ListView1.ListItems.Item(j).Text)) = "sheet1" Then
          Set oWS = oWB.Worksheets(ListView1.ListItems.Item(j).Text)
          WrkSheet = "Sheet1"
          If LCase(Trim(oWS.Cells(1, 1))) <> "cycle" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 2))) <> "isin" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of B " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 3))) <> "dpid" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of C " & vbCrLf
             Status = 1
          End If
       
          If LCase(Trim(oWS.Cells(1, 4))) <> "dp name" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of D " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 5))) <> "clid" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of E " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 6))) <> "client name" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of F " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 7))) <> "qty" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of G " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 8))) <> "start date" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of H " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 9))) <> "end date" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of I " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 10))) <> "amount" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of J " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 11))) <> "amt. incl. st" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of K " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 12))) <> "warehouse name" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of L " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 13))) <> "commodity" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of M " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 14))) <> "status" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of N " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 15))) <> "warehouse code" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of O " & vbCrLf
             Status = 1
          End If
          
       Else
          ErrMsg = ErrMsg + "Please Select Sheet1 "
          Status = 1
       End If
    End If
Next
End Sub
Private Sub FillData()
Dim intNumberRows As Double
For j = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(j).Checked Then
       Set oWS = oWB.Worksheets(ListView1.ListItems.Item(j).Text)
       oWS.ClearArrows
       mRow = 1
       RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
       'RowCount = oWS.Rows.Count
       ProgressBar1.Value = 0
       ProgressBar1.Max = RowCount
       'For L = 1 To RowCount
       Do
         If oWS.Cells(mRow, 1) = "" Then
            Exit Do
         End If
                    
         If oWS.Cells(mRow, 1) <> "" Then
            For c = 0 To MSHFlexGrid1.Cols - 1
                MSHFlexGrid1.TextMatrix(mRow, c) = Trim(oWS.Cells(mRow + 1, c + 1).Value)
            Next
            MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
         End If
         ProgressBar1.Value = ProgressBar1.Value + 1
         mRow = mRow + 1
       Loop Until False
       oXL.Quit
    End If
Next
ProgressBar1.Value = 0
End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
    .Rows = 2
    .Cols = 15
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Cycle"
    .TextMatrix(0, 1) = "ISIN"
    .TextMatrix(0, 2) = "DPID"
    .TextMatrix(0, 3) = "DP Name"
    .TextMatrix(0, 4) = "CL ID"
    .TextMatrix(0, 5) = "Client Name"
    .TextMatrix(0, 6) = "Quantity"
    .TextMatrix(0, 7) = "Start Date"
    .TextMatrix(0, 8) = "End Date"
    .TextMatrix(0, 9) = "Amount"
    .TextMatrix(0, 10) = "Amt.Incl.ST"
    .TextMatrix(0, 11) = "Warehouse Name"
    .TextMatrix(0, 12) = "Commodity"
    .TextMatrix(0, 13) = "Status"
    .TextMatrix(0, 14) = "Warehouse Code"
    .ColWidth(0) = 1000
    .ColWidth(1) = 1800
    .ColWidth(2) = 1500
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = 0
    .ColWidth(6) = 900
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 0
    .ColWidth(12) = 2500
    .ColWidth(13) = 0
    .ColWidth(14) = 4200
End With
End Sub

Private Sub CmdSave_Click()

Dim Location_ID As Variant, Commodity_ID As Variant, ISIN_ID As Variant
Dim DP_ID As Variant, Qty_xls As Variant, Cycle_ID As Variant, CMP_ID As Variant 'used for Warehouse code
Dim AmtInclST As Variant, Amt As Variant
Dim CancelUpdate As Boolean

Call TableTxn_DPIncomeDetails
Call TableTmp_DPIncomeDetails
Call TableMst_CMP
Call TableMst_ISIN
ErrMsg = ""
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows - 1
mRow = 1
If LCase(WrkSheet) = "sheet1" Then
   Do
      If MSHFlexGrid1.TextMatrix(mRow, 1) = "" Then
         Exit Do
      End If
      Status = 0
      wc = ""
      
      wc = " Filename='" & Namef & "'"
      
      If MSHFlexGrid1.TextMatrix(mRow, 1) <> "" Then
         RsMst_ISIN.MoveFirst
         RsMst_ISIN.Find " ISINID='" & MSHFlexGrid1.TextMatrix(mRow, 1) & "'"
         If RsMst_ISIN.EOF Then
            MSHFlexGrid1.Row = mRow
            MSHFlexGrid1.Col = 1
            MSHFlexGrid1.CellBackColor = vbRed
            ErrMsg = ErrMsg + "ISIN  " + MSHFlexGrid1.TextMatrix(mRow, 1) + " Not Found of Row " & mRow + 1 & vbCrLf
            Status = 1
         Else
            ISIN_ID = MSHFlexGrid1.TextMatrix(mRow, 1)
            Location_ID = RsMst_ISIN!LocationID
            Commodity_ID = RsMst_ISIN!CommodityID
            CMP_ID = RsMst_ISIN!CMPID
            wc = wc + " and ISINID='" & ISIN_ID & "' and LocationID=" & Location_ID & " and CommodityID=" & Commodity_ID & " and CMPID=" & CMP_ID & ""
         End If
      Else
         ErrMsg = ErrMsg + "ISIN  ID is blank at row " & mRow + 1 & vbCrLf
         Status = 1
      End If

      If MSHFlexGrid1.TextMatrix(mRow, 0) <> "" Then
         Cycle_ID = MSHFlexGrid1.TextMatrix(mRow, 0)
         wc = wc + " and Cycle=" & Cycle_ID & ""
      Else
         ErrMsg = ErrMsg + "Cycle  ID is blank at row " & mRow + 1 & vbCrLf
         Status = 1
      End If
                                
      If MSHFlexGrid1.TextMatrix(mRow, 2) <> "" Then
         DP_ID = MSHFlexGrid1.TextMatrix(mRow, 2)
         'Wc = Wc + " and DPID='" & RsTxn_DPIncomeDetails!DPID & "'"
      Else
         ErrMsg = ErrMsg + "DP ID is blank at row " & mRow + 1 & vbCrLf
         Status = 1
      End If
                        
      If MSHFlexGrid1.TextMatrix(mRow, 6) <> "" Then
         Qty_xls = MSHFlexGrid1.TextMatrix(mRow, 6)
         'Wc = Wc + " and Qty=" & RsTxn_DPIncomeDetails!Qty & ""
      Else
         ErrMsg = ErrMsg + "Quantity is blank at row " & mRow + 1 & vbCrLf
         Status = 1
      End If
                        
      If MSHFlexGrid1.TextMatrix(mRow, 9) <> "" Then
         Amt = MSHFlexGrid1.TextMatrix(mRow, 9)
         'Wc = Wc + " and Amount=" & RsTxn_DPIncomeDetails!Amount & ""
      Else
         ErrMsg = ErrMsg + "Amount is blank at row " & mRow + 1 & vbCrLf
         Status = 1
      End If
                        
      If MSHFlexGrid1.TextMatrix(mRow, 10) <> "" Then
         AmtInclST = MSHFlexGrid1.TextMatrix(mRow, 10)
         'Wc = Wc + " and Amt_Incl_ST=" & RsTxn_DPIncomeDetails!Amt_Incl_ST & ""
      Else
         ErrMsg = ErrMsg + "Amt_Incl_ST is blank at row " & mRow + 1 & vbCrLf
         Status = 1
      End If
                            
'      CommodityID is Found above so that no need to find again
'-----------------------------------------------------------------------
'      tmp = "Select * from Txn_DPIncomeDetails where " + wc
'      Call TmpTable
'      If TmpRs.RecordCount > 0 Then
'         Status = 1
'      End If

      If Status = 0 Then
         RsTmp_DPIncomeDetails.AddNew
         RsTmp_DPIncomeDetails!FileName = Namef
         RsTmp_DPIncomeDetails!ISINID = ISIN_ID
         RsTmp_DPIncomeDetails!LocationID = Location_ID
         RsTmp_DPIncomeDetails!CommodityID = Commodity_ID
         RsTmp_DPIncomeDetails!Cycle = Cycle_ID
         RsTmp_DPIncomeDetails!DPID = DP_ID
         RsTmp_DPIncomeDetails!Qty = Qty_xls
         RsTmp_DPIncomeDetails!Amount = Amt
         RsTmp_DPIncomeDetails!Amt_Incl_ST = AmtInclST
         RsTmp_DPIncomeDetails!CMPID = CMP_ID
         RsTmp_DPIncomeDetails.Update
      End If
      ProgressBar1.Value = ProgressBar1.Value + 1
      mRow = mRow + 1
   Loop Until False
   
tmp = "Select * from Tmp_DPIncomeDetails "
'Filename, ISINID, LocationID, CommodityID,CMPID, Cycle
Call TmpTable

For i = 0 To TmpRs.RecordCount - 1
    RsTxn_DPIncomeDetails.Filter = ""
    If RsTxn_DPIncomeDetails.RecordCount > 0 Then
       RsTxn_DPIncomeDetails.MoveFirst
    End If
    
    RsTxn_DPIncomeDetails.Filter = " Filename ='" & TmpRs!FileName & "' And ISINID='" & TmpRs!ISINID & "' And LocationID =" & TmpRs!LocationID & " And CMPID=" & TmpRs!CMPID & " And Cycle=" & TmpRs!Cycle & " "
    If RsTxn_DPIncomeDetails.RecordCount = 0 Then
       RsTxn_DPIncomeDetails.AddNew
       RsTxn_DPIncomeDetails!FileName = TmpRs!FileName
       RsTxn_DPIncomeDetails!Cycle = TmpRs!Cycle
       RsTxn_DPIncomeDetails!LocationID = TmpRs!LocationID
       RsTxn_DPIncomeDetails!ISINID = TmpRs!ISINID
       RsTxn_DPIncomeDetails!DPID = TmpRs!DPID
       RsTxn_DPIncomeDetails!Qty = TmpRs!Qty
       RsTxn_DPIncomeDetails!Amount = TmpRs!Amount
       RsTxn_DPIncomeDetails!Amt_Incl_ST = TmpRs!Amt_Incl_ST
       RsTxn_DPIncomeDetails!CommodityID = TmpRs!CommodityID
       RsTxn_DPIncomeDetails!CMPID = TmpRs!CMPID
       RsTxn_DPIncomeDetails!G_UID = GetGUID
       RsTxn_DPIncomeDetails.Update
    End If
    TmpRs.MoveNext
Next
   
End If
    
tmp = "delete Tmp_DPIncomeDetails"
Call TmpTable
ProgressBar1.Value = 0
If ErrMsg = "" Then
   MsgBox "Data Saved To Database Successfully !!!", vbInformation
Else
   TxtErrorMsg.Text = "Following Data is Not Inserted  into Database Due To Error !..." & vbCrLf & ErrMsg
   MsgBox "Data Saved To Database Successfully Except above error List", vbInformation
End If
CmdSave.Enabled = False
CmdGetFile.Enabled = True
CmdExit.Enabled = True
oXL.Quit
End Sub

Private Sub Form_Load()
LblFile.Caption = ""
Frame1.Enabled = False
Frame2.Enabled = False
CmdSave.Enabled = False
ProgressBar1.Value = 0

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdExit.Enabled = False
CmdReset.Enabled = False

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Index <> 1 Then
   Item.Checked = False
End If
End Sub

