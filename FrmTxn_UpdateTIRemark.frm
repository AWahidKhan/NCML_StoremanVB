VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_UpdateTIRemark 
   Caption         =   "Collection Remark against Tax Invoices "
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   9240
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   15000
      TabIndex        =   23
      Top             =   -120
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   8760
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Frame Frame0 
         Height          =   8565
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   14745
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   345
            Left            =   10845
            TabIndex        =   28
            Top             =   960
            Width           =   1125
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Remarks Updated only"
            Height          =   360
            Left            =   5280
            TabIndex        =   27
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "All Invoices"
            Height          =   375
            Left            =   3000
            TabIndex        =   25
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "Clear"
            Height          =   350
            Left            =   12990
            TabIndex        =   22
            Top             =   960
            Width           =   1000
         End
         Begin VB.TextBox TxtRemark 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   240
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtNAVCode 
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
            Height          =   360
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox TxtSTINo 
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
            Height          =   360
            Left            =   2925
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Search"
            Height          =   345
            Left            =   9720
            TabIndex        =   6
            Top             =   960
            Width           =   1125
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "Update"
            Height          =   350
            Left            =   11970
            TabIndex        =   5
            Top             =   960
            Width           =   1020
         End
         Begin VB.TextBox TxtSPostInvoiceNo 
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
            Height          =   360
            Left            =   5505
            TabIndex        =   4
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox TxtSEntryNo 
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
            Height          =   360
            Left            =   8085
            TabIndex        =   3
            Top             =   360
            Width           =   2655
         End
         Begin VB.ComboBox CmbSLocation 
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
            ItemData        =   "FrmTxn_UpdateTIRemark.frx":0000
            Left            =   10800
            List            =   "FrmTxn_UpdateTIRemark.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   360
            Width           =   3240
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6855
            Left            =   90
            TabIndex        =   8
            Top             =   1560
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12091
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   345
            Left            =   120
            TabIndex        =   9
            Top             =   375
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
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
            Format          =   3211265
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1575
            TabIndex        =   10
            Top             =   375
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   3211265
            CurrentDate     =   39884
         End
         Begin VB.Label Label1 
            Caption         =   "Client NAV Code"
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
            Left            =   840
            TabIndex        =   20
            Top             =   780
            Width           =   1635
         End
         Begin VB.Label Label33 
            Caption         =   "TI No (Voucher No)"
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
            Left            =   3300
            TabIndex        =   18
            Top             =   120
            Width           =   1785
         End
         Begin VB.Label Label16 
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
            Left            =   5520
            TabIndex        =   17
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "_"
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
            Left            =   8160
            TabIndex        =   16
            Top             =   840
            Width           =   135
         End
         Begin VB.Label Label12 
            Caption         =   "From TI Date"
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
            Left            =   172
            TabIndex        =   15
            Top             =   120
            Width           =   1230
         End
         Begin VB.Label Label14 
            Caption         =   "To TI Date"
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
            Left            =   1725
            TabIndex        =   14
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label13 
            Caption         =   "Post Invoice No"
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
            Left            =   6067
            TabIndex        =   13
            Top             =   120
            Width           =   1410
         End
         Begin VB.Label Label18 
            Caption         =   "Entry No"
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
            Left            =   9015
            TabIndex        =   12
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label19 
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
            Height          =   210
            Left            =   11970
            TabIndex        =   11
            Top             =   120
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_UpdateTIRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSelRow As Variant
Dim mSelCol As Variant

Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub CmdClear_Click()
Call Grid_Head
Label15.Caption = "-"
SFrom.Value = Date - 30
STo.Value = Date
TxtNAVCode = ""
TxtRemark = ""
TxtSEntryNo = ""
TxtSPostInvoiceNo = ""
TxtSTINo = ""
CmbSLocation.Text = ""
Check1.Value = 0
ProgressBar1.Value = 0
Check2.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_UpdateTIRemarks.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_UpdateTIRemarks.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub CmdGo_Click()

Dim wc As Variant

wc = ""

If SFrom.Value > STo.Value Then
   MsgBox " From Date should be less than To Date!!"
   Exit Sub
End If

Call Grid_Head
ProgressBar1.Value = 0
MsgBox "Wait till next message!!!"


tmp = " Select TI.TIID,TI.TIDate,TI.TINo,TI.PostInvoiceNo,TI.EntryNo,TI.CollectionRemark,TIAmount,TI.TISTAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecOtherDed"
tmp = tmp & " ,TI.NAVCode,'' ClientName,ML.LocationName,TI.EntryFlag,TI.AdjAmountFAG,TI.PIFlag,TI.PIID,TI.ClientIDRow,TI.DepositoryID"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Where  TINo<> '' "

If Gen_DBType = "MDB" Then
   wc = " And TI.TIDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   wc = " And TI.TIDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If


If TxtSTINo <> "" Then
   wc = wc & " And TI.TINO like '%" & TxtSTINo & "%'"
End If

If TxtSPostInvoiceNo <> "" Then
   wc = wc & " And TI.PostInvoiceNo like '%" & TxtSPostInvoiceNo & "%'"
End If

If TxtSEntryNo <> "" Then
   wc = wc & " And TI.EntryNo like '%" & TxtSEntryNo & "%'"
End If

If CmbSLocation.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbSLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = wc & " AND ML." & Trim(Gen_WcLocation)
   End If
End If

If TxtNAVCode <> "" Then
   wc = wc & " And TI.NAVCOde like '%" & TxtNAVCode & "%'"
End If

If Check2.Value = 1 Then
   wc = wc & " And TI.CollectionRemark is not null And TI.CollectionRemark <> '' "
End If

''If CmbSEntryType.Text <> "" Then
''   If LCase(CmbSEntryType.Text) = "system" Then
''      wc = wc & " And (TI.EntryFlag = '" & CmbSEntryType.Text & "' Or TI.EntryFlag is Null) "
''   Else
''      wc = wc & " And TI.EntryFlag = '" & CmbSEntryType.Text & "'"
''   End If
''End If

tmp = tmp & wc & " Order By TI.TIDate"

Call TmpTable

ProgressBar1.Max = TmpRs.RecordCount + 2

For u = 1 To TmpRs.RecordCount
    mRecAmt = IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount) + IIf(IsNull(TmpRs!RecTDSAmount), 0, TmpRs!RecTDSAmount) + IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed) + IIf(IsNull(TmpRs!AdjAmountFAG), 0, TmpRs!AdjAmountFAG)
    mRemAmt = TmpRs!TIAmount - (IIf(IsNull(TmpRs!RecAmount), 0, TmpRs!RecAmount) + IIf(IsNull(TmpRs!RecTDSAmount), 0, TmpRs!RecTDSAmount) + IIf(IsNull(TmpRs!RecOtherDed), 0, TmpRs!RecOtherDed) + IIf(IsNull(TmpRs!AdjAmountFAG), 0, TmpRs!AdjAmountFAG))
    If Check1.Value = 1 Then
       iu = MSHFlexGrid1.Rows - 1
       For C = 1 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C - 1)), "", TmpRs.Fields(C - 1))
       Next
       MSHFlexGrid1.TextMatrix(iu, 1) = Format(MSHFlexGrid1.TextMatrix(iu, 1), "dd-mmm-yyyy")
       If MSHFlexGrid1.TextMatrix(iu, 12) = "" Then
          MSHFlexGrid1.TextMatrix(iu, 12) = "System"
       End If
        
       If IsNull(TmpRs!ClientIDRow) = True And IsNull(DepositoryID) = True Then
          tmp = "Select ClientID,ClientName"
          tmp = tmp & " from Mst_Client "
          ''tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=TCRD.DPID"
          tmp = tmp & " Where  NAVCode='" & TmpRs!NAVCode & "' or RNAVCode='" & TmpRs!NAVCode & "' "
       
          Call Tmp6Table
       
          If TmpRs6.RecordCount > 0 Then
             MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!ClientName & "~" & TmpRs6!ClientID
          Else
             tmp = "Select DepositoryID,DepositoryName"
             tmp = tmp & " from Mst_Depository  Where NAVCode='" & TmpRs!NAVCode & "'"
             Call Tmp6Table
             If TmpRs6.RecordCount > 0 Then
                MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!DepositoryName & "~" & TmpRs6!DepositoryID
             End If
          End If
       Else
          If IsNull(TmpRs!ClientIDRow) = False Then
             tmp = "Select ClientID,ClientName"
             tmp = tmp & " from Mst_Client "
             tmp = tmp & " Where ClientIDRow = " & TmpRs!ClientIDRow & " "
       
             Call Tmp6Table
       
             If TmpRs6.RecordCount > 0 Then
                MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!ClientName & "~" & TmpRs6!ClientID
             End If
          ElseIf IsNull(TmpRs!DepositoryID) = False Then
             tmp = "Select DepositoryID,DepositoryName"
             tmp = tmp & " from Mst_Depository  Where DepositoryID = '" & TmpRs!DepositoryID & "' "
             Call Tmp6Table
             If TmpRs6.RecordCount > 0 Then
                MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!DepositoryName & "~" & TmpRs6!DepositoryID
             End If
          End If
       End If
         
       MSHFlexGrid1.TextMatrix(iu, 9) = mRecAmt
       MSHFlexGrid1.TextMatrix(iu, 10) = mRemAmt
       MSHFlexGrid1.row = iu
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.Text = strUnChecked
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Else
       If mRemAmt > 0 Then
          iu = MSHFlexGrid1.Rows - 1
                       
          For C = 1 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C - 1)), "", TmpRs.Fields(C - 1))
          Next
          MSHFlexGrid1.TextMatrix(iu, 2) = Format(MSHFlexGrid1.TextMatrix(iu, 2), "dd-mmm-yyyy")
          If MSHFlexGrid1.TextMatrix(iu, 12) = "" Then
             MSHFlexGrid1.TextMatrix(iu, 12) = "System"
          End If
          
            If IsNull(TmpRs!ClientIDRow) = True And IsNull(DepositoryID) = True Then
               tmp = "Select ClientID,ClientName"
               tmp = tmp & " from Mst_Client "
               ''tmp = tmp & " Left join Mst_Depository MD on MD.DepositoryID=TCRD.DPID"
               tmp = tmp & " Where  NAVCode='" & TmpRs!NAVCode & "' or RNAVCode='" & TmpRs!NAVCode & "' "
            
               Call Tmp6Table
            
               If TmpRs6.RecordCount > 0 Then
                  MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!ClientName & "~" & TmpRs6!ClientID
               Else
                  tmp = "Select DepositoryID,DepositoryName"
                  tmp = tmp & " from Mst_Depository  Where NAVCode='" & TmpRs!NAVCode & "'"
                  Call Tmp6Table
                  If TmpRs6.RecordCount > 0 Then
                     MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!DepositoryName & "~" & TmpRs6!DepositoryID
                  End If
               End If
            Else
               If IsNull(TmpRs!ClientIDRow) = False Then
                  tmp = "Select ClientID,ClientName"
                  tmp = tmp & " from Mst_Client "
                  tmp = tmp & " Where ClientIDRow = " & TmpRs!ClientIDRow & " "
            
                  Call Tmp6Table
            
                  If TmpRs6.RecordCount > 0 Then
                     MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!ClientName & "~" & TmpRs6!ClientID
                  End If
               ElseIf IsNull(TmpRs!DepositoryID) = False Then
                  tmp = "Select DepositoryID,DepositoryName"
                  tmp = tmp & " from Mst_Depository  Where DepositoryID = '" & TmpRs!DepositoryID & "' "
                  Call Tmp6Table
                  If TmpRs6.RecordCount > 0 Then
                     MSHFlexGrid1.TextMatrix(iu, 13) = TmpRs6!DepositoryName & "~" & TmpRs6!DepositoryID
                  End If
               End If
            End If
          MSHFlexGrid1.TextMatrix(iu, 9) = mRecAmt
          MSHFlexGrid1.TextMatrix(iu, 10) = mRemAmt
          MSHFlexGrid1.row = iu
          MSHFlexGrid1.Col = 0
          MSHFlexGrid1.CellFontName = "Wingdings"
          MSHFlexGrid1.CellFontSize = 12
          MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid1.Text = strUnChecked
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
     End If
   End If
   TmpRs.MoveNext
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!!"
Label15.Caption = MSHFlexGrid1.Rows - 2

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.RowHeightMin = 600
MSHFlexGrid1.Cols = 19
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 0
MSHFlexGrid1.TextMatrix(0, 0) = "Select"
MSHFlexGrid1.TextMatrix(0, 1) = "TI ID"
MSHFlexGrid1.TextMatrix(0, 2) = "TI Date"
MSHFlexGrid1.TextMatrix(0, 3) = "TI No"
MSHFlexGrid1.TextMatrix(0, 4) = "Post Invoice No"
MSHFlexGrid1.TextMatrix(0, 5) = "Entry No"
MSHFlexGrid1.TextMatrix(0, 6) = "Collection Remark"
MSHFlexGrid1.TextMatrix(0, 7) = "TI Amount"
MSHFlexGrid1.TextMatrix(0, 8) = "TI Service Tax Amount"
MSHFlexGrid1.TextMatrix(0, 9) = "Received Amount"
MSHFlexGrid1.TextMatrix(0, 10) = "Balance Amount"
MSHFlexGrid1.TextMatrix(0, 11) = "Received Other Ded."
MSHFlexGrid1.TextMatrix(0, 12) = "Client NAV Code"
MSHFlexGrid1.TextMatrix(0, 13) = "ClientName"
MSHFlexGrid1.TextMatrix(0, 14) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 15) = "Entry Flag"
MSHFlexGrid1.TextMatrix(0, 16) = "Adjustment Entry by FAG"
MSHFlexGrid1.TextMatrix(0, 17) = "Invoice Type"
MSHFlexGrid1.TextMatrix(0, 18) = "SM ID"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 0
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 800
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 0
MSHFlexGrid1.ColWidth(6) = 2600
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 0
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 0
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1800
MSHFlexGrid1.ColWidth(14) = 1800
MSHFlexGrid1.ColWidth(15) = 0
MSHFlexGrid1.ColWidth(16) = 0
MSHFlexGrid1.ColWidth(17) = 0
MSHFlexGrid1.ColWidth(18) = 0

MSHFlexGrid1.RowHeight(0) = 900

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
'Call TableMst_Location(" Where LocationId > 0 ")

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
End Sub


Private Sub CmdSave_Click()
ProgressBar1.Value = 0
If CheckGridSelect = False Then
   MsgBox "Please select atleast one Record to Update!"
   Exit Sub
End If
Call TableTxn_TIDetails(" Where TIDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "'")
Call TableLog_TICollectionRemark
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       If RsTxn_TIDetails.RecordCount > 0 Then RsTxn_TIDetails.MoveFirst
       RsTxn_TIDetails.Find "TIID=" & Val(MSHFlexGrid1.TextMatrix(I, 1))
       If RsTxn_TIDetails.EOF = False Then
          RsLog_TICollectionRemark.AddNew
          RsLog_TICollectionRemark!TIID = Val(MSHFlexGrid1.TextMatrix(I, 1))
          RsLog_TICollectionRemark!OldCollectionRemark = RsTxn_TIDetails!CollectionRemark
          RsLog_TICollectionRemark!NewCollectionRemark = MSHFlexGrid1.TextMatrix(I, 6)
          RsLog_TICollectionRemark!UpdatedBy = Gen_UserLoginID
          RsLog_TICollectionRemark!UpdatedDate = Now
          RsLog_TICollectionRemark.Update
          RsLog_TICollectionRemark.Requery
          
          RsTxn_TIDetails!CollectionRemark = MSHFlexGrid1.TextMatrix(I, 6)
          RsTxn_TIDetails!CRemarkUpdatedBy = Gen_UserLoginID
          RsTxn_TIDetails!CRemarkUpdatedDate = Now
          RsTxn_TIDetails.Update
          RsTxn_TIDetails.Requery

       End If
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!"
Call CmdClear_Click
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

Call CmdClear_Click
End Sub

Private Sub MSHFlexGrid1_Click()
TxtRemark.Visible = False
With MSHFlexGrid1
    If .row < 1 Or .row = .Rows - 1 Then Exit Sub
    If .Col = 0 Then
      If .TextMatrix(.row, .Col) = "" Then Exit Sub
      .Col = 0
      If .TextMatrix(.row, .Col) = strUnChecked Then
         .TextMatrix(.row, .Col) = strChecked
      Else
         .TextMatrix(.row, .Col) = strUnChecked
      End If
   End If
   
   If .Col = 6 Then
      If .TextMatrix(.row, 0) = strUnChecked Then Exit Sub
      mSelRow = .row
      mSelCol = .Col
      TxtRemark.Visible = True
      TxtRemark.Left = .Left + .CellLeft
      TxtRemark.Top = .Top + .CellTop
      TxtRemark.Width = .CellWidth - 10
      TxtRemark.Height = .CellHeight - 10
      TxtRemark.Text = .Text
      TxtRemark.SelStart = 0
      TxtRemark.SelLength = Len(.Text)
      TxtRemark.ZOrder
      TxtRemark.SetFocus
   End If
   
End With

End Sub

Private Sub TxtRemark_LostFocus()
If TxtRemark.Text = "" Then
   TxtRemark.Visible = False
   Exit Sub
End If

MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = Replace(Trim(TxtRemark.Text), vbCrLf, " ")
TxtRemark.Visible = False
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If

End Sub
Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub

Private Sub HScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l

End Sub

Private Sub SetScrollBars()
'    Frame1.Width = 10000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
End Sub


Private Function CheckGridSelect() As Boolean
Dim Retval As Boolean
Retval = False

For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckGridSelect = Retval
End Function
