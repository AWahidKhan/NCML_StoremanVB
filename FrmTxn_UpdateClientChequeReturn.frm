VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_UpdateClientChequeReturn 
   Caption         =   "Update Client Cheque Return"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8895
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   15030
      Begin VB.TextBox TxtEnterData 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1920
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2400
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CheckBox ChkSelectAll 
         Caption         =   "Select All"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13361
         _Version        =   393216
         FixedCols       =   0
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   8400
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
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
         Left            =   13485
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OptClient 
         Caption         =   "Client"
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
         Left            =   240
         TabIndex        =   1
         Top             =   233
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox TxtClient 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4080
         MaxLength       =   200
         TabIndex        =   4
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox TxtNAVCode 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   8400
         MaxLength       =   200
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton OptDepositor 
         Caption         =   "Depositor"
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
         Left            =   240
         TabIndex        =   2
         Top             =   653
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
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
         Height          =   360
         Left            =   10800
         TabIndex        =   6
         Top             =   600
         Width           =   1335
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
         Height          =   360
         Left            =   13095
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
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
         Height          =   360
         Left            =   12150
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmTxn_UpdateClientChequeReturn.frx":0000
         Left            =   1560
         List            =   "FrmTxn_UpdateClientChequeReturn.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Left            =   5910
         TabIndex        =   14
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NAV Code"
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
         Left            =   8827
         TabIndex        =   13
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label14 
         Caption         =   "Exchange Type"
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
         Left            =   2100
         TabIndex        =   12
         Top             =   233
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Total Record Found :"
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
         Left            =   10560
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LblTotalRecord 
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
         Left            =   12960
         TabIndex        =   10
         Top             =   240
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmTxn_UpdateClientChequeReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Variant
Dim mColsel, mRowsel As Single
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub ChkSelectAll_Click()
If MSHFlexGrid1.Rows = 2 Then
    ChkSelectAll.Value = vbUnchecked
    Exit Sub
End If

If ChkSelectAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub

Private Sub CmdClear_Click()

OptClient.Value = True
Call GridHead
CmbExchangeTypeID.Text = ""
TxtClient.Text = ""
TxtNAVCode.Text = ""

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_UpdateClientChequeReturn.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_UpdateClientChequeReturn.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

Call GridHead
Dim wc As String

wc = ""
If CmbExchangeTypeID.Text <> "" Then
    wc = " And MET.ExchangeTypeID = " & mExchangeTypeID
End If
If OptClient.Value = True Then
    tmp = " Select"
    tmp = tmp & " MCL.ClientIDRow ID, MET.ExchangeType, MCL.ClientName Name, isnull(MCL.NAVCode, MCL.RNAVCode) NAVCode, MCL.ChequeReturn"
    tmp = tmp & " From Mst_Client MCL"
    tmp = tmp & " Inner Join Mst_ExchangeType MET on MET.ExchangeTypeID = MCL.ExchangeTypeID"
    tmp = tmp & " Where MCL.ChequeReturn = 'Y'"
    tmp1 = " order by MET.ExchangeType, MCL.ClientName"
    
    If TxtClient.Text <> "" Then
        wc = wc & " And MCL.ClientName like '%" & TxtClient.Text & "%'"
    End If
    If TxtNAVCode.Text <> "" Then
        wc = wc & " And (MCL.NAVCode like '%" & TxtNAVCode.Text & "%' Or MCL.RNAVCode like '%" & TxtNAVCode.Text & "%') "
    End If
    tmp = tmp & wc & tmp1
    
    Call TmpTable
    
ElseIf OptDepositor.Value = True Then
    tmp = " Select"
    tmp = tmp & " MD.DepositoryID ID, MET.ExchangeType, MD.DepositoryName Name, MD.NAVCode, MD.ChequeReturn"
    tmp = tmp & " From Mst_Depository MD"
    tmp = tmp & " Inner Join Mst_ExchangeType MET on MET.ExchangeTypeID = MD.ExchangeTypeID"
    tmp = tmp & " Where MD.ChequeReturn = 'Y'"
    tmp1 = " order by MET.ExchangeType, MD.DepositoryName"
    
    If TxtClient.Text <> "" Then
        wc = wc & " And MD.DepositoryName like '%" & TxtClient.Text & "%'"
    End If
    If TxtNAVCode.Text <> "" Then
        wc = wc & " And MD.NAVCode like '%" & TxtNAVCode.Text & "%'"
    End If
    tmp = tmp & wc & tmp1
    
    Call TmpTable
End If
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2
LblTotalRecord.Caption = TmpRs.RecordCount
If TmpRs.RecordCount > 0 Then
    CmdExcel.Enabled = True
    CmdUpdate.Enabled = True
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
    With MSHFlexGrid1
        For I = 1 To TmpRs.RecordCount
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!ID), "", TmpRs!ID) '"ID"
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType) '"Exchange Type"
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!Name), "", TmpRs!Name) '"Name"
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode) '"NAV Code"
            .TextMatrix(I, 5) = IIf(IsNull(TmpRs!ChequeReturn), "", TmpRs!ChequeReturn) '"Cheque Return"
            If .TextMatrix(I, 5) = "Y" Then
                .TextMatrix(I, 5) = "Yes"
            Else
                .TextMatrix(I, 5) = "No"
            End If
            MSHFlexGrid1.Row = I
            MSHFlexGrid1.Col = 0
            MSHFlexGrid1.CellFontName = "Wingdings"
            MSHFlexGrid1.CellFontSize = 12
            MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
            MSHFlexGrid1.Text = strUnChecked
            TmpRs.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
        ProgressBar1.Value = ProgressBar1.Max
        MsgBox "Done!!!"
    End With
Else
    MsgBox "No Record Found!!!"
End If
End Sub

Private Sub CmdUpdate_Click()
If ChkRemarks = False Then Exit Sub

tmp = " select * from Log_Mst_CLientDepository"
Call Tmp1Table

ProgressBar1.Value = 0
If OptClient.Value = True Then
    For I = 1 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
            tmp = " Update Mst_CLient Set ChequeReturn = 'N' Where ClientIDRow = " & MSHFlexGrid1.TextMatrix(I, 1)
            Call TmpTable
            
            ''''''For Inserted into Table Log_Mst_ClientDepository''''''''''''''''
            TmpRs1.AddNew
            TmpRs1!ClientDepositoryID = MSHFlexGrid1.TextMatrix(I, 1)
            TmpRs1!ClientDepositorFlag = "C"
            TmpRs1!UpdatedBy = Gen_UserLoginID
            TmpRs1!NAVCode = MSHFlexGrid1.TextMatrix(I, 4)
            TmpRs1!Remarks = MSHFlexGrid1.TextMatrix(I, 6)
            TmpRs1.Update
            
        End If
    ProgressBar1.Value = ProgressBar1.Value + 1
    Next
Else
    For I = 1 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
            tmp = " Update Mst_Depository Set ChequeReturn = 'N' Where DepositoryID = '" & MSHFlexGrid1.TextMatrix(I, 1) & "'"
            Call TmpTable
            
            ''''''For Inserted into Table Log_Mst_ClientDepository''''''''''''''''
            TmpRs1.AddNew
            TmpRs1!ClientDepositoryID = MSHFlexGrid1.TextMatrix(I, 1)
            TmpRs1!ClientDepositorFlag = "D"
            TmpRs1!UpdatedBy = Gen_UserLoginID
            TmpRs1!NAVCode = MSHFlexGrid1.TextMatrix(I, 4)
            TmpRs1!Remarks = MSHFlexGrid1.TextMatrix(I, 6)
            TmpRs1.Update
        End If
    ProgressBar1.Value = ProgressBar1.Value + 1
    Next
End If
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!!"
Call CmdClear_Click
End Sub

Private Function ChkRemarks() As Boolean
ChkRemarks = True
For I = 1 To MSHFlexGrid1.Rows - 1
   If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And Trim(MSHFlexGrid1.TextMatrix(I, 6)) = "" Then
      MsgBox "Blank Remarks not allowed!!!"
      ChkRemarks = False
      Exit Function
   End If
Next
End Function

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Call GridHead
CmdExcel.Enabled = False
CmdUpdate.Enabled = False
End Sub


Private Function GridHead()
ProgressBar1.Value = 0
LblTotalRecord.Caption = ""
CmdExcel.Enabled = False
CmdUpdate.Enabled = False
ChkSelectAll.Value = vbUnchecked
With MSHFlexGrid1
    
    .Clear
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    .Cols = 7
    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "ID"
    .TextMatrix(0, 2) = "Exchange Type"
    .TextMatrix(0, 3) = "Name"
    .TextMatrix(0, 4) = "NAV Code"
    .TextMatrix(0, 5) = "Cheque Return"
    .TextMatrix(0, 6) = "Remarks"
    
    .ColWidth(0) = 700
    .ColWidth(1) = 1000
    .ColWidth(2) = 1600
    .ColWidth(3) = 4500
    .ColWidth(4) = 1200
    .ColWidth(5) = 1000
    .ColWidth(6) = 4000
    '.ColAlignment(0) = ali
    .RowHeight(0) = 600
    .WordWrap = True
End With
End Function

Private Sub MSHFlexGrid1_Click()
mColsel = MSHFlexGrid1.ColSel
mRowsel = MSHFlexGrid1.RowSel
If mColsel = 0 And mRowsel <> MSHFlexGrid1.Rows - 1 Then
   If MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strUnChecked Then
      MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strChecked
   Else
      MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strUnChecked
   End If
ElseIf MSHFlexGrid1.TextMatrix(mRowsel, 0) = strChecked Then
   If mColsel = 6 And mRowsel <> MSHFlexGrid1.Rows - 1 Then
      Call VisibleEnterTxt
   End If
End If
End Sub

Private Sub VisibleEnterTxt()
With TxtEnterData
   .Text = MSHFlexGrid1.TextMatrix(mRowsel, mColsel)
   .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
   .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
   .Width = MSHFlexGrid1.CellWidth - 10
   .Height = MSHFlexGrid1.CellHeight - 10
   .Visible = True
   .ZOrder
   .SetFocus
End With
End Sub


Private Sub OptClient_Click()
Call GridHead
End Sub

Private Sub OptDepositor_Click()
Call GridHead
End Sub

Private Sub TxtEnterData_LostFocus()

MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = TxtEnterData.Text
TxtEnterData.Visible = False

End Sub
