VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_Cust2CustTransfer 
   Caption         =   "Client to Client Transfer"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   9780
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Transfer Log Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10215
      Left            =   0
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   15255
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
         Height          =   525
         Left            =   12480
         TabIndex        =   32
         Top             =   8880
         Width           =   2145
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10320
         TabIndex        =   31
         Top             =   8880
         Width           =   2145
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   8325
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   14684
         _Version        =   393216
         RowHeightMin    =   370
         AllowBigSelection=   0   'False
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Invoice Details"
      TabPicture(0)   =   "FrmTxn_Cust2CustTransfer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Client having Same NAVCode"
      TabPicture(1)   =   "FrmTxn_Cust2CustTransfer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   5895
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   10398
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
      End
      Begin VB.Frame Frame2 
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
         Height          =   6285
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   14730
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   11940
            TabIndex        =   21
            Top             =   60
            Width           =   495
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   5685
            Left            =   0
            TabIndex        =   22
            Top             =   480
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   10028
            _Version        =   393216
            RowHeightMin    =   370
            AllowBigSelection=   0   'False
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
         Begin VB.Label Label6 
            Caption         =   "Payment Pending"
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
            Left            =   12630
            TabIndex        =   25
            Top             =   75
            Width           =   1665
         End
         Begin VB.Label Label2 
            Caption         =   "Total Invoices :"
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
            Left            =   8655
            TabIndex        =   24
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label LblTotRec 
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
            Left            =   10335
            TabIndex        =   23
            Top             =   60
            Width           =   90
         End
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   8640
      Width           =   14970
      Begin VB.CommandButton CmdTransferLogs 
         Caption         =   "Show Transfer Logs"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2520
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Transfer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10500
         TabIndex        =   8
         Top             =   240
         Width           =   2145
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12660
         TabIndex        =   9
         Top             =   240
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   14985
      Begin VB.CommandButton CmdExcelInvoice 
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
         Height          =   435
         Left            =   13020
         TabIndex        =   33
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox TxtTClientName 
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1215
         Width           =   4560
      End
      Begin VB.TextBox TxtFClientName 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1215
         Width           =   4560
      End
      Begin VB.OptionButton OptDP 
         Caption         =   "DP"
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
         Left            =   6000
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptReser 
         Caption         =   "Reservation"
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
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptNCDEX 
         Caption         =   "NCDEX / Non NCDEX"
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
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Search Invoices"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11955
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox TxtToClient 
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
         Left            =   8400
         TabIndex        =   5
         Top             =   720
         Width           =   1860
      End
      Begin VB.ComboBox CmbToClient 
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
         ItemData        =   "FrmTxn_Cust2CustTransfer.frx":0038
         Left            =   8400
         List            =   "FrmTxn_Cust2CustTransfer.frx":003A
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   750
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtFromClient 
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
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   1860
      End
      Begin VB.ComboBox CmbFromClient 
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
         Left            =   1935
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   750
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "To Client Name"
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
         Left            =   6855
         TabIndex        =   19
         Top             =   1275
         Width           =   1785
      End
      Begin VB.Label Label3 
         Caption         =   "From Client Name"
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
         Left            =   90
         TabIndex        =   18
         Top             =   1275
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Client Type"
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
         Left            =   105
         TabIndex        =   15
         Top             =   315
         Width           =   2055
      End
      Begin VB.Label Label45 
         Caption         =   "To Client"
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
         Left            =   6840
         TabIndex        =   14
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label Label44 
         Caption         =   "From Client"
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
         TabIndex        =   13
         Top             =   803
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmTxn_Cust2CustTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp1 As String
Dim mBalFlag As String
Private Sub CmbFromClient_GotFocus()
tmp1 = CmbFromClient.Text

If OptNCDEX.Value = True Then
   tmp = " Select Distinct NAVCode,max(ClientName)ClientName From Mst_Client Where isnull(NAVCode,'')<>'' Group By NAVCode"
ElseIf OptReser.Value = True Then
   tmp = " Select Distinct RNAVCode as NAVCode,max(ClientName)ClientName  From Mst_Client Where isnull(RNAVCode,'')<>'' Group By RNAVCode"
ElseIf OptDP.Value = True Then
   tmp = " Select Distinct NAVCode,max(DepositoryName)ClientName From Mst_Depository Where isnull(NAVCode,'')<>'' Group By NAVCode"
End If
Call Tmp1Table
CmbFromClient.Clear
If TmpRs1.RecordCount = 0 Then
   MsgBox "No NAV Code found"
   CmbFromClient.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbFromClient.AddItem TmpRs1!NAVCode
    TmpRs1.MoveNext
Next
CmbFromClient.Text = tmp1
End Sub

Private Sub CmbFromClient_LostFocus()
If CmbFromClient.Text = "" Then Exit Sub
TmpRs1.MoveFirst
TmpRs1.Find "NAVCode = '" & CmbFromClient.Text & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid NAV Code selection "
   CmbFromClient.SetFocus
   Exit Sub
End If
TxtFClientName.Text = TmpRs1!ClientName
If tmp1 <> CmbFromClient.Text Then
   Call Grid_Head
End If
End Sub

Private Sub CmbToClient_GotFocus()
tmp1 = CmbToClient.Text

If OptNCDEX.Value = True Then
   tmp = " Select Distinct NAVCode,max(ClientName)ClientName From Mst_Client Where isnull(NAVCode,'')<>'' Group By NAVCode"
ElseIf OptReser.Value = True Then
   tmp = " Select Distinct RNAVCode as NAVCode,max(ClientName)ClientName  From Mst_Client Where isnull(RNAVCode,'')<>'' Group By RNAVCode"
ElseIf OptDP.Value = True Then
   tmp = " Select Distinct NAVCode,max(DepositoryName)ClientName From Mst_Depository Where isnull(NAVCode,'')<>'' Group By NAVCode"
End If
Call Tmp1Table
CmbToClient.Clear
If TmpRs1.RecordCount = 0 Then
   MsgBox "No NAV Code found"
   CmbToClient.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
    CmbToClient.AddItem TmpRs1!NAVCode
    TmpRs1.MoveNext
Next
CmbToClient.Text = tmp1
End Sub

Private Sub CmbToClient_LostFocus()
If CmbToClient.Text = "" Then Exit Sub
TmpRs1.MoveFirst
TmpRs1.Find "NAVCode = '" & CmbToClient.Text & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid NAV Code selection "
   CmbToClient.SetFocus
   Exit Sub
End If
TxtTClientName.Text = TmpRs1!ClientName
End Sub


Private Sub Grid_Head()
LblTotRec.Caption = ""
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .Cols = 9
    .FixedRows = 2
    .WordWrap = True

    .TextMatrix(1, 0) = "TI No"
    .TextMatrix(1, 1) = "TI Date"
    .TextMatrix(1, 2) = "TI Amount"
    .TextMatrix(1, 3) = "Received Amt"
    .TextMatrix(1, 4) = "TDS Deducted"
    .TextMatrix(1, 5) = "Other Deduction"
    .TextMatrix(1, 6) = "Balance Amt"
    .TextMatrix(1, 7) = "PostInvoice No"
    .TextMatrix(1, 8) = "Entry No"

    .ColWidth(0) = 2200
    .ColWidth(1) = 1500
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 2200
    .ColWidth(8) = 1500

    .RowHeight(1) = 800
End With
End Sub
Private Sub Grid_HeadTransferLog()

With MSHFlexGrid3
    .Clear
    .Rows = 3
    .Cols = 6
    .FixedRows = 1
    .WordWrap = True

    .TextMatrix(0, 0) = "Log Date"
    .TextMatrix(0, 1) = "Old NAVCode"
    .TextMatrix(0, 2) = "New NAVCode"
    .TextMatrix(0, 3) = "Transfered By"
    .TextMatrix(0, 4) = "Transfered On"
    .TextMatrix(0, 5) = "Client/DP Updated"

    .ColWidth(0) = 1600
    .ColWidth(1) = 1500
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 3000
    .ColWidth(5) = 3000
   
    .RowHeight(0) = 800
    
End With
End Sub
Private Sub GetTransferLogDetail()

'tmp = "Select * from Log_Client2ClientTransfer Order By Concurrency "

tmp = " Select TxnDate,OldNAVCode,NewNAVCode,CreatedBy,Concurrency,ClientIDRows from Log_Client2ClientTransfer Order By Concurrency "

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       For c = 0 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid3.TextMatrix(I, 0) = Format(MSHFlexGrid3.TextMatrix(I, 0), "dd-mmm-yyyy")
       MSHFlexGrid3.TextMatrix(I, 4) = Format(MSHFlexGrid3.TextMatrix(I, 4), "dd-mmm-yyyy hh:mm:ss")
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Log_BlockingClient.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Log_BlockingClient.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid3.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Public Sub Xls_Detail_RptInvoice()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Invoices_BlockingClient.xls")
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
    
        MSHFlexGrid1.Row = I
        MSHFlexGrid1.Col = c
        
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, c + 1), oWS.Cells(mRow, c + 1)).Interior.Color = vbRed
        End If
    
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub CmdExcelInvoice_Click()

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Invoices_BlockingClient.xls")
Call Xls_Detail_RptInvoice

End Sub

Private Sub CmdGo_Click()


tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0)+isnull(TI.AdjAmountFAG,0)RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And (TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo >0))RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID and(TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo > 0))RecOtherDed"
tmp = tmp & " ,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Where TI.NAVCode='" & TxtFromClient.Text & "'" 'tmp = tmp & " Where TI.NAVCode='" & CmbFromClient.Text & "'"
tmp = tmp & " Order By TI.TIDate,TI.NAVCode"

Call Tmp3Table

Call Grid_Head
mBalFlag = "Y"
If TmpRs3.RecordCount > 0 Then
   For gr = 1 To TmpRs3.RecordCount
       rn = MSHFlexGrid1.Rows - 1

       MSHFlexGrid1.TextMatrix(rn, 0) = TmpRs3!TINo
       MSHFlexGrid1.TextMatrix(rn, 1) = IIf(IsNull(TmpRs3!TIDate), "", Format(TmpRs3!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(rn, 2) = IIf(IsNull(TmpRs3!TIAmount), 0, TmpRs3!TIAmount)
       MSHFlexGrid1.TextMatrix(rn, 3) = IIf(IsNull(TmpRs3!RecAmount), 0, TmpRs3!RecAmount)
       MSHFlexGrid1.TextMatrix(rn, 4) = IIf(IsNull(TmpRs3!RecTDSAmount), 0, TmpRs3!RecTDSAmount)
       MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs3!RecOtherDed), 0, TmpRs3!RecOtherDed)
       'MSHFlexGrid1.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 2)) - (Val(MSHFlexGrid1.TextMatrix(rn, 3)) + Val(MSHFlexGrid1.TextMatrix(rn, 4)) + Val(MSHFlexGrid1.TextMatrix(rn, 5))), "#####0.00")
       MSHFlexGrid1.TextMatrix(rn, 7) = IIf(IsNull(TmpRs3!PostInvoiceNo), 0, TmpRs3!PostInvoiceNo)
       MSHFlexGrid1.TextMatrix(rn, 8) = IIf(IsNull(TmpRs3!EntryNo), 0, TmpRs3!EntryNo)
       
      
       TmpRs3.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   Next
   
   tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,"
   tmp = tmp & " isnull(AdjAmountFAG,0) + isnull(Sum(TCRT.RecAmount),0) as RecAmount,"
   tmp = tmp & " isnull(Sum(TCRT.RecTDS),0) as RecTDS,"
   tmp = tmp & " isnull(Sum(TCRT.RecOtherDed), 0) As RecOtherDed"
   tmp = tmp & " From  Txn_NewCashReceiptTIDetail TCRT"
   tmp = tmp & " inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
   tmp = tmp & " inner Join Txn_TIDetails TI on TI.TIID=TCRT.TIID"
   tmp = tmp & " inner Join Txn_CRAdvanceAllocation CRAA on CRAA.CRAAID=TCRT.CRAAID"
   tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
   tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
   tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
   tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
   tmp = tmp & " Where TCR.ReturnFlag<>'Y' And TCR.RejectFlag<>'Y' And isnull(CRAA.BatchID,0)>0"
   tmp = tmp & " And TI.NAVCode='" & TxtFromClient.Text & "'" 'tmp = tmp & " And TI.NAVCode='" & CmbFromClient.Text & "'"
   tmp = tmp & " Group By TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode, AdjAmountFAG"

   Call Tmp5Table
    
   
    For I = 2 To MSHFlexGrid1.Rows - 2
        If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
        TmpRs5.Find "EntryNo=" & Val(MSHFlexGrid1.TextMatrix(I, 8))
        If TmpRs5.EOF = False Then
           MSHFlexGrid1.TextMatrix(I, 3) = Val(MSHFlexGrid1.TextMatrix(I, 3)) + IIf(IsNull(TmpRs5!RecAmount), 0, TmpRs5!RecAmount)
        End If
        MSHFlexGrid1.TextMatrix(I, 6) = Format(Val(MSHFlexGrid1.TextMatrix(I, 2)) - (Val(MSHFlexGrid1.TextMatrix(I, 3)) + Val(MSHFlexGrid1.TextMatrix(I, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 5))), "#####0.00")
        MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(I, 3))
        MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
        MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
        MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
        If Val(MSHFlexGrid1.TextMatrix(I, 6)) > 0 Then
           For c = 0 To MSHFlexGrid1.Cols - 1
               MSHFlexGrid1.Col = c
               MSHFlexGrid1.Row = I
               MSHFlexGrid1.CellBackColor = vbRed
           Next
           mBalFlag = "N"
        End If
    Next
End If

LblTotRec.Caption = MSHFlexGrid1.Rows - 3

End Sub

Private Sub CmdOk_Click()
Frame4.Visible = False
End Sub

Private Sub CmdTransferLogs_Click()

Frame4.Top = Frame1.Top
Frame4.Left = Frame1.Left
Frame4.Height = ButtonFrm.Top + 1000
Frame4.Width = Frame1.Width
Frame4.Visible = True
Call Grid_HeadTransferLog
Call GetTransferLogDetail

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   ButtonFrm.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
ButtonFrm.Enabled = True
Call OptNCDEX_Click
End Sub

Private Sub OptDP_Click()
Call OptNCDEX_Click
End Sub

Private Sub OptNCDEX_Click()
TxtFromClient.Text = "" 'CmbFromClient.Text = ""
TxtToClient.Text = ""  'CmbToClient.Text = ""
TxtTClientName = ""
TxtFClientName = ""
Call Grid_Head
Call Grid_Head1
LblTotRec.Caption = "-"
End Sub

Private Sub OptReser_Click()
Call OptNCDEX_Click
End Sub

Private Sub SaveCmd_Click()
If TxtFromClient.Text = "" Then 'If CmbFromClient.Text = "" Then
   MsgBox "Please select From Client NAV Code!!"
   TxtFromClient.SetFocus '    CmbFromClient.SetFocus
   Exit Sub
End If

If TxtToClient.Text = "" Then  'If CmbToClient.Text = "" Then
   MsgBox "Please select To Client NAV Code!!"
   TxtToClient.SetFocus '    CmbToClient.SetFocus
   Exit Sub
End If

If TxtFromClient.Text = TxtToClient.Text Then ''If CmbFromClient.Text = CmbToClient.Text Then
   MsgBox "To Client should be different from From Client!!"
   TxtToClient.SetFocus ''    CmbToClient.SetFocus
   Exit Sub
End If

'If Left(TxtFromClient.Text, 4) <> Left(TxtToClient.Text, 4) Then  'If Left(CmbFromClient.Text, 4) <> Left(CmbToClient.Text, 4) Then
'   MsgBox "Client type of both Clients should be same!!"
'   TxtToClient.SetFocus   '    CmbToClient.SetFocus
'   Exit Sub
'End If

If Val(MSHFlexGrid1.TextMatrix(0, 6)) <> 0 Then
   MsgBox "Payment Pending for one or more Invoices!!"
   Exit Sub
End If

'If mBalFlag = "N" Then
'   MsgBox "Payment Pending for one or more Invoices!!"
'   Exit Sub
'End If

If OptNCDEX.Value = True Then
   'tmp = " Update Mst_Client Set NAVCode='" & CmbToClient.Text & "' Where NAVCode='" & CmbFromClient.Text & "'"
   tmp = " Update Mst_Client Set NAVCode='" & TxtToClient.Text & "' Where NAVCode='" & TxtFromClient.Text & "'"
ElseIf OptReser.Value = True Then
   'tmp = " Update Mst_Client Set RNAVCode='" & CmbToClient.Text & "' Where RNAVCode='" & CmbFromClient.Text & "'"
   tmp = " Update Mst_Client Set RNAVCode='" & TxtToClient.Text & "' Where RNAVCode='" & TxtFromClient.Text & "'"
ElseIf OptDP.Value = True Then
   'tmp = " Update Mst_Depository Set NAVCode='" & CmbToClient.Text & "' Where NAVCode='" & CmbFromClient.Text & "'"
   tmp = " Update Mst_Depository Set NAVCode='" & TxtToClient.Text & "' Where NAVCode='" & TxtFromClient.Text & "'"
End If

Call TmpTable
Call TableLog_Client2ClientTransfer

mClientIDRows = ""
RsLog_Client2ClientTransfer.AddNew
RsLog_Client2ClientTransfer!TxnDate = Date
RsLog_Client2ClientTransfer!OldNAVCode = TxtFromClient.Text 'CmbFromClient.Text
RsLog_Client2ClientTransfer!NewNAVCode = TxtToClient.Text ' CmbToClient.Text
For I = 1 To MSHFlexGrid2.Rows - 2
    mClientIDRows = mClientIDRows & MSHFlexGrid2.TextMatrix(I, 0) & ","
Next
RsLog_Client2ClientTransfer!CreatedBy = Gen_UserLoginID
RsLog_Client2ClientTransfer!ClientIDRows = Mid(mClientIDRows, 1, Len(mClientIDRows) - 1)
RsLog_Client2ClientTransfer.Update

MsgBox "Done!!"
Call OptNCDEX_Click
End Sub

Private Sub TxtFromClient_LostFocus()
If TxtFromClient.Text = "" Then
   Call Grid_Head
   Call Grid_Head1
   Exit Sub
End If
If OptNCDEX.Value = True Then
   tmp = " Select Distinct NAVCode,max(ClientName)ClientName From Mst_Client Where isnull(NAVCode,'')<>'' And NAVCode = '" & TxtFromClient.Text & "' Group By NAVCode"
ElseIf OptReser.Value = True Then
   tmp = " Select Distinct RNAVCode as NAVCode,max(ClientName)ClientName  From Mst_Client Where isnull(RNAVCode,'')<>'' And RNAVCode = '" & TxtFromClient.Text & "' Group By RNAVCode"
ElseIf OptDP.Value = True Then
   tmp = " Select Distinct NAVCode,max(DepositoryName)ClientName From Mst_Depository Where isnull(NAVCode,'')<>'' And NAVCode = '" & TxtFromClient.Text & "' Group By NAVCode"
End If
Call Tmp1Table

If TmpRs1.RecordCount = 0 Then
    MsgBox "Invalid NAV Code"
    TxtFromClient.SetFocus
    Exit Sub
End If

TxtFClientName.Text = TmpRs1!ClientName
Call Grid_Head
Call FillClientDetails
Call CmdGo_Click

End Sub

Private Sub FillClientDetails()
Call Grid_Head1

If OptDP.Value = True Then
    tmp = "Select MC.DepositoryID,MC.DepositoryName,ME.ExchangeType,MC.NAVCode from Mst_Depository MC"
    tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
    tmp = tmp & " Where MC.NAVCode='" & TxtFromClient & "'"
    Call Tmp2Table
    For I = 1 To TmpRs2.RecordCount
       rn = MSHFlexGrid2.Rows - 1
       MSHFlexGrid2.TextMatrix(rn, 0) = IIf(IsNull(TmpRs2!DepositoryID), "", TmpRs2!DepositoryID)
       MSHFlexGrid2.TextMatrix(rn, 1) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
       MSHFlexGrid2.TextMatrix(rn, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
       MSHFlexGrid2.TextMatrix(rn, 3) = IIf(IsNull(TmpRs2!NAVCode), "", TmpRs2!NAVCode)
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       TmpRs2.MoveNext
    Next
Else
    tmp = " Select MC.ClientIDRow,MC.ClientName,ME.ExchangeType,MC.NAVCode,MC.RNAVCode from Mst_Client MC"
    tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
    If OptReser.Value = True Then
       tmp = tmp & " Where MC.RNAVCode='" & TxtFromClient & "'" ''''MC.NAVCode='" & TxtFromClient & "' or
    Else
       tmp = tmp & " Where MC.NAVCode='" & TxtFromClient & "'"
    End If
    
    Call Tmp2Table
    For I = 1 To TmpRs2.RecordCount
        rn = MSHFlexGrid2.Rows - 1
        MSHFlexGrid2.TextMatrix(rn, 0) = IIf(IsNull(TmpRs2!ClientIDRow), "", TmpRs2!ClientIDRow)
        MSHFlexGrid2.TextMatrix(rn, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid2.TextMatrix(rn, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
        MSHFlexGrid2.TextMatrix(rn, 3) = IIf(IsNull(TmpRs2!NAVCode), "", TmpRs2!NAVCode)
        MSHFlexGrid2.TextMatrix(rn, 4) = IIf(IsNull(TmpRs2!RNAVCode), "", TmpRs2!RNAVCode)
        MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
        TmpRs2.MoveNext
    Next
End If
End Sub
Private Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    If OptDP.Value = False Then
       .Cols = 5
    Else
       .Cols = 4
    End If
    .FixedRows = 1
    .WordWrap = True

    If OptDP.Value = False Then
       .TextMatrix(0, 0) = "Client ID Row"
       .TextMatrix(0, 1) = "Client Name"
       .TextMatrix(0, 2) = "Exchange Type"
       .TextMatrix(0, 3) = "NAV Code"
       .TextMatrix(0, 4) = "Reservation NAV Code"
       .ColWidth(4) = 2500
    Else
       .TextMatrix(0, 0) = "DP ID "
       .TextMatrix(0, 1) = "DP Name"
       .TextMatrix(0, 2) = "Exchange Type"
       .TextMatrix(0, 3) = "NAV Code"
    End If

    .ColWidth(0) = 1800
    .ColWidth(1) = 4500
    .ColWidth(2) = 2000
    .ColWidth(3) = 1800
    
End With
End Sub


Private Sub TxtToClient_LostFocus()
If TxtToClient.Text = "" Then Exit Sub
If OptNCDEX.Value = True Then
   tmp = " Select Distinct NAVCode,max(ClientName)ClientName From Mst_Client Where isnull(NAVCode,'')<>'' And NAVCode = '" & TxtToClient.Text & "' Group By NAVCode"
ElseIf OptReser.Value = True Then
   tmp = " Select Distinct RNAVCode as NAVCode,max(ClientName)ClientName  From Mst_Client Where isnull(RNAVCode,'')<>'' And RNAVCode = '" & TxtToClient.Text & "' Group By RNAVCode"
ElseIf OptDP.Value = True Then
   tmp = " Select Distinct NAVCode,max(DepositoryName)ClientName From Mst_Depository Where isnull(NAVCode,'')<>'' And NAVCode = '" & TxtToClient.Text & "' Group By NAVCode"
End If
Call Tmp1Table

'If TmpRs1.RecordCount = 0 Then
'    MsgBox "Invalid NAV Code"
'    TxtToClient.SetFocus
'    Exit Sub
'End If

If TmpRs1.RecordCount = 0 Then
    'MsgBox "Invalid NAV Code"
    If MsgBox("NAV Code not exist in Client/DP master. Still do you want to continue?", vbYesNo) = vbNo Then
       TxtToClient.SetFocus
    End If
    TxtTClientName.Text = ""
    Exit Sub
End If

TxtTClientName.Text = TmpRs1!ClientName

End Sub
