VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_ClientLedger 
   Caption         =   "Client Ledger Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Clients with same NAV Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   7200
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1695
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   2990
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
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox Check3 
         Caption         =   "Show DP "
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
         Left            =   4620
         TabIndex        =   33
         Top             =   1050
         Width           =   1140
      End
      Begin VB.ComboBox CmbDPName 
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
         Left            =   6255
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   135
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CmbNAVCode 
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
         Left            =   1545
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1035
         Width           =   2970
      End
      Begin VB.ComboBox CmbCNonNcdexClient 
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
         Left            =   5640
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CmbClientDP 
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
         Left            =   9600
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   390
         Width           =   4335
      End
      Begin VB.TextBox TxtCDPName 
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
         Height          =   375
         Left            =   9615
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   390
         Width           =   4335
      End
      Begin VB.TextBox TxtClientID 
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
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   390
         Width           =   1575
      End
      Begin VB.TextBox TxtCClientName 
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
         Height          =   375
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   390
         Width           =   3495
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         Left            =   1545
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   390
         Width           =   2910
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
         Height          =   375
         Left            =   11280
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
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
         Height          =   375
         Left            =   12495
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
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
         Height          =   375
         Left            =   13710
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   90
         TabIndex        =   0
         Top             =   390
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
         Format          =   109051905
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   105
         TabIndex        =   1
         Top             =   1035
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   109051905
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label35 
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
         Height          =   255
         Left            =   2543
         TabIndex        =   31
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Last Uploaded Date and Time :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   3420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "- -"
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
         Left            =   3675
         TabIndex        =   27
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label Label27 
         Caption         =   "Client Name "
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
         Left            =   8640
         TabIndex        =   26
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label LblExchangeType 
         Caption         =   "Exchange Type"
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
         Left            =   2183
         TabIndex        =   22
         Top             =   150
         Width           =   1635
      End
      Begin VB.Label Label24 
         Caption         =   "ClientDPID"
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
         Left            =   10800
         TabIndex        =   21
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "DP Name "
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
         Left            =   12240
         TabIndex        =   20
         Top             =   150
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "DP ID"
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
         Left            =   9975
         TabIndex        =   19
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Client ID"
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
         Left            =   4860
         TabIndex        =   18
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   6960
         TabIndex        =   15
         Top             =   1560
         Width           =   2280
      End
      Begin VB.Label Label2 
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
         Left            =   9315
         TabIndex        =   14
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To  Date"
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
         Left            =   390
         TabIndex        =   13
         Top             =   795
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "From  Date"
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
         Left            =   285
         TabIndex        =   12
         Top             =   150
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ledger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   1905
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4665
         Left            =   120
         TabIndex        =   10
         Top             =   510
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8229
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "- -"
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
         Left            =   7515
         TabIndex        =   30
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label3 
         Caption         =   "Balance Amount :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5400
         TabIndex        =   29
         Top             =   240
         Width           =   1890
      End
   End
End
Attribute VB_Name = "FrmRpt_ClientLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mClientNAVCode As String
Dim mClientRNAVCode As String
Dim mDPNAVCode As String
Dim mBalance As Variant
Private Sub CmdClear_Click()

TxtFromDate = Date - 30
TxtToDate = Date
Call Grid_Head
Call Grid_Head1
Label2.Caption = ""
TxtClientID = ""
TxtCClientName = ""
TxtCDPID = ""
TxtCDPName = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
CmbExchangeTypeID.Text = ""
mExchangeTypeID = 0
CmbCNonNcdexClient.Text = ""
CmdExcel.Enabled = False
Label5.Caption = ""
CmbNAVCode.Clear
CmbNAVCode.Text = ""
Check3.Value = 0
CmbCNonNcdexClient.Width = 9400
CmbCNonNcdexClient.Top = TxtClientID.Top
CmbCNonNcdexClient.Left = TxtClientID.Left
CmbCNonNcdexClient.Visible = True
Call Check3_Click


End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
Call Grid_Head1
Label5.Caption = ""

MsgBox "Wait till next message!!!"

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "'From Date' must be less 'To Date' always!!!"
   TxtToDate.SetFocus
   Exit Sub
End If

If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select Exchange Type!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

''If mExchangeTypeID <> 1 Then
''   If TxtClientID = "" Then
''      MsgBox "Select Client!!!"
''      TxtClientID.SetFocus
''      Exit Sub
''   End If
''Else
''   If CmbCNonNcdexClient.Text = "" Then
''      MsgBox "Select Client!!!"
''      CmbCNonNcdexClient.SetFocus
''      Exit Sub
''   End If
''End If

If Check3.Value = 0 Then
   If CmbCNonNcdexClient.Text = "" Then
      MsgBox "Select Client!!!"
      CmbCNonNcdexClient.SetFocus
      Exit Sub
   End If
Else
   If CmbDPName.Visible = True Then
      If CmbDPName.Text = "" Then
         MsgBox "Blank DP Name Not allowed !!! "
         CmbDPName.SetFocus
         Exit Sub
      End If
   End If
End If

'For Op Balance

' For Ledger
tmp = " Select CustomerNo,CustomerName,EntryNo,EHOPInvoiceNo,DocumentType,DocumentNo,VoucherNo,PostingDate,InvoiceAmount,"
tmp = tmp & "  Narration, ActualDatePosted, ActualTimePosted, CustomerPostingGroup"
tmp = tmp & " From Txn_ClientLedger  "

If Gen_DBType = "MDB" Then
   wc = " Where ActualDatePosted Between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "# "
Else
   wc = " Where ActualDatePosted Between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And '" & Format(TxtToDate.Value, Gen_DatFormat) & "' "
End If

If CmbNAVCode.Text <> "" Then
   wc = wc & " And (CustomerNo = '" & CmbNAVCode.Text & "')"  ' Or CustomerNo = '" & mClientRNAVCode & "') "
Else
   If Check3.Value = 0 Then
      wc = wc & " And (CustomerNo = '" & mClientNAVCode & "' Or CustomerNo = '" & mClientRNAVCode & "') "
   Else
      wc = wc & " And CustomerNo = '" & mDPNAVCode & "' "
   End If
End If



tmp = tmp & wc
tmp = tmp & " Order By PostingDate "

Call TmpTable

Call displedger

Call dispclient


CmdExcel.Enabled = True

MsgBox ("Done!!!")

End Sub
Private Sub displedger()

Dim x As Double

x = 2

tmp = " Select Isnull(Sum(InvoiceAmount),0.00) 'OpBalance' From Txn_ClientLedger "
If Gen_DBType = "MDB" Then
   wc = " Where ActualDatePosted < #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# " ' And #" & TxtToDate.Value & "# "
Else
   wc = " Where ActualDatePosted < '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' " 'And '" & TxtToDate.Value & "' "
End If


If CmbNAVCode.Text <> "" Then
   wc = wc & " And (CustomerNo = '" & CmbNAVCode.Text & "')"  ' Or CustomerNo = '" & mClientRNAVCode & "') "
Else
   If Check3.Value = 0 Then
      wc = wc & " And (CustomerNo = '" & mClientNAVCode & "' Or CustomerNo = '" & mClientRNAVCode & "') "
   Else
      wc = wc & " And CustomerNo = '" & mDPNAVCode & "' "
   End If
End If

'wc = wc & " And (CustomerNo = '" & mClientNAVCode & "' Or CustomerNo = '" & mClientRNAVCode & "') "

tmp = tmp & wc

Call Tmp1Table

mBalance = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x, 0) = "" 'IIf(IsNull(TmpRs!DocumentType), "", TmpRs!DocumentType)
   MSHFlexGrid1.TextMatrix(x, 1) = "" 'IIf(IsNull(TmpRs!DocumentNo), "", TmpRs!DocumentNo)
   MSHFlexGrid1.TextMatrix(x, 2) = "Opening Balance for " & Format(TxtFromDate.Value, "dd-mmm-yyyy") & "" 'IIf(IsNull(TmpRs!Narration), "", TmpRs!Narration) 'IIf(IsNull(TmpRs!Narration), "", TmpRs!Narration)
   
   
   If Val(TmpRs1!OpBalance) < 0 Then
      MSHFlexGrid1.TextMatrix(x, 3) = IIf(IsNull(TmpRs1!OpBalance), 0, TmpRs1!OpBalance)
      MSHFlexGrid1.TextMatrix(x, 3) = Abs(Format(MSHFlexGrid1.TextMatrix(x, 3), "########0.00")) 'Abs(Format(TmpRs1!OpBalance, "########0.00")) 'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
   Else
      MSHFlexGrid1.TextMatrix(x, 4) = IIf(IsNull(TmpRs1!OpBalance), 0, TmpRs1!OpBalance)
      MSHFlexGrid1.TextMatrix(x, 4) = Abs(Format(MSHFlexGrid1.TextMatrix(x, 4), "########0.00")) 'Abs(Format(TmpRs1!OpBalance, "########0.00")) 'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
   End If
   
   'MSHFlexGrid1.TextMatrix(x, 5) = Val(MSHFlexGrid1.TextMatrix(x, 3)) - Val(MSHFlexGrid1.TextMatrix(x, 4)) 'IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
   
   MSHFlexGrid1.TextMatrix(x, 6) = ""
   MSHFlexGrid1.TextMatrix(x, 7) = ""
   MSHFlexGrid1.TextMatrix(x, 8) = ""
   MSHFlexGrid1.TextMatrix(x, 9) = ""
   MSHFlexGrid1.TextMatrix(x, 10) = ""
   MSHFlexGrid1.TextMatrix(x, 11) = ""
   MSHFlexGrid1.TextMatrix(x, 12) = ""
   MSHFlexGrid1.TextMatrix(x, 13) = ""
   'MSHFlexGrid1.TextMatrix(x, 14) = ""
   
   MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(x, 3)) '+ Format(TmpRs!TransactionAmount, "########0.00")  'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
   MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(x, 4)) 'Format(TmpRs!TransactionAmount, "########0.00")  'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
'   MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5))
   MSHFlexGrid1.TextMatrix(0, 3) = Format(MSHFlexGrid1.TextMatrix(0, 3), "########0.00") 'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
   MSHFlexGrid1.TextMatrix(0, 4) = Format(MSHFlexGrid1.TextMatrix(0, 4), "########0.00")
'   MSHFlexGrid1.TextMatrix(0, 5) = Format(MSHFlexGrid1.TextMatrix(0, 5), "########0.00")
   
   mBalance = Val(mBalance) - (Val(MSHFlexGrid1.TextMatrix(x, 3)) + Val(MSHFlexGrid1.TextMatrix(x, 4)))

End If

x = x + 1

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       
       MSHFlexGrid1.TextMatrix(x, 0) = IIf(IsNull(TmpRs!DocumentType), "", TmpRs!DocumentType)
       MSHFlexGrid1.TextMatrix(x, 1) = IIf(IsNull(TmpRs!DocumentNo), "", TmpRs!DocumentNo)
       MSHFlexGrid1.TextMatrix(x, 2) = IIf(IsNull(TmpRs!Narration), "", TmpRs!Narration)
'       MSHFlexGrid1.TextMatrix(x, 3) = IIf(IsNull(TmpRs!InvoiceAmount), 0, TmpRs!InvoiceAmount)
'       MSHFlexGrid1.TextMatrix(x, 4) = IIf(IsNull(TmpRs!BaseAmount), 0, TmpRs!BaseAmount)
''       MSHFlexGrid1.TextMatrix(x, 5) = IIf(IsNull(TmpRs!TaxAmount), 0, TmpRs!TaxAmount)
''       MSHFlexGrid1.TextMatrix(x, 6) = IIf(IsNull(TmpRs!RemainingAmount), 0, TmpRs!RemainingAmount)
       
       If Val(TmpRs!InvoiceAmount) < 0 Then
          MSHFlexGrid1.TextMatrix(x, 3) = Abs(Format(TmpRs!InvoiceAmount, "########0.00")) 'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
'          MSHFlexGrid1.TextMatrix(x, 0) = "Invoice"
       Else
          MSHFlexGrid1.TextMatrix(x, 4) = Format(TmpRs!InvoiceAmount, "########0.00") 'IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
'          MSHFlexGrid1.TextMatrix(x, 0) = "Payment"
       End If
'       MSHFlexGrid1.TextMatrix(x, 5) = Val(MSHFlexGrid1.TextMatrix(x, 3)) - Val(MSHFlexGrid1.TextMatrix(x, 4)) 'IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
       
       MSHFlexGrid1.TextMatrix(x, 5) = IIf(IsNull(TmpRs!CustomerName), "", TmpRs!CustomerName)
       MSHFlexGrid1.TextMatrix(x, 6) = IIf(IsNull(TmpRs!CustomerNo), "", TmpRs!CustomerNo)
       MSHFlexGrid1.TextMatrix(x, 7) = IIf(IsNull(TmpRs!PostingDate), "", Format(TmpRs!PostingDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(x, 8) = IIf(IsNull(TmpRs!VoucherNo), "", TmpRs!VoucherNo)
       MSHFlexGrid1.TextMatrix(x, 9) = IIf(IsNull(TmpRs!EHOPInvoiceNo), "", TmpRs!EHOPInvoiceNo)
       MSHFlexGrid1.TextMatrix(x, 10) = IIf(IsNull(TmpRs!ActualDatePosted), "", Format(TmpRs!ActualDatePosted, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(x, 11) = IIf(IsNull(TmpRs!ActualTimePosted), "", TmpRs!ActualTimePosted)
       MSHFlexGrid1.TextMatrix(x, 12) = IIf(IsNull(TmpRs!CustomerPostingGroup), "", TmpRs!CustomerPostingGroup)
       MSHFlexGrid1.TextMatrix(x, 13) = IIf(IsNull(TmpRs!EntryNo), "", TmpRs!EntryNo)
       
       MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(x, 3)) 'IIf(IsNull(TmpRs!InvoiceAmount), 0, TmpRs!InvoiceAmount)
       MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(x, 4))
'       MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5))
       'mBalance = Val(mBalance) + Val(MSHFlexGrid1.TextMatrix(x, 5))
       mBalance = Val(mBalance) - (Val(MSHFlexGrid1.TextMatrix(x, 3)) + Val(MSHFlexGrid1.TextMatrix(x, 4)))
       
       TmpRs.MoveNext
       x = x + 1

   Next
End If

Label2.Refresh
Label2.Caption = TmpRs.RecordCount
Label5.Refresh
Label5.Caption = Val(MSHFlexGrid1.TextMatrix(0, 4)) - Val(MSHFlexGrid1.TextMatrix(0, 3)) 'Val(mBalance)
Label5.Caption = Format(Label5.Caption, "########0.00")

End Sub
Private Sub dispclient()
Call Grid_Head1

If Check3.Value = 1 Then
    tmp = "Select MC.DepositoryID,MC.DepositoryName,ME.ExchangeType,MC.NAVCode from Mst_Depository MC"
    tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
    tmp = tmp & " Where MC.NAVCode='" & mDPNAVCode & "'"
    Call Tmp2Table
    For I = 1 To TmpRs2.RecordCount
       rn = MSHFlexGrid2.Rows - 1
       MSHFlexGrid2.TextMatrix(rn, 0) = IIf(IsNull(TmpRs2!DepositoryID), "", TmpRs2!DepositoryID)
       MSHFlexGrid2.TextMatrix(rn, 1) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
       MSHFlexGrid2.TextMatrix(rn, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
       MSHFlexGrid2.TextMatrix(rn, 3) = IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode)
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       TmpRs2.MoveNext
    Next
    
       
Else
        
    tmp = " Select ClientID , ClientName,'',NAVCode, RNAVCode From Mst_Client "
    tmp = tmp & " Where ExchangeTypeID = " & mExchangeTypeID & " And (NAVCode = '" & mClientNAVCode & "' Or RNAVCode = '" & mClientRNAVCode & "')  "
    
    Call TmpTable
    
    MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
    
    For I = 1 To TmpRs.RecordCount
        For C = 0 To MSHFlexGrid2.Cols - 1
            MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
        Next
        MSHFlexGrid2.TextMatrix(I, 2) = CmbExchangeTypeID.Text
        TmpRs.MoveNext
    Next
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ClientLedger.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ClientLedger.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Ledger Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 2
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Client Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 2
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        End If
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Form_Load()
Call Grid_Head
Call Grid_Head1
Dim FrmLoadAccess() As Boolean

FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

Call CmdClear_Click

tmp = "Select Max(ActualDatePosted) as 'LastUpdation' from Txn_ClientLedger Where ActualDatePosted <= GetDate() "
Call TmpTable

Label8.Caption = "-"
If TmpRs.RecordCount > 0 Then
   Label8.Caption = Format(TmpRs!LastUpdation, "dd-mmm-yyyy")
End If

If Label8.Caption <> "-" Then
   tmp = "Select Max(ActualTimePosted) as 'LastUpdation' from Txn_ClientLedger Where ActualDatePosted = '" & Label8.Caption & "' "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      Label8.Caption = Label8.Caption & " " & TmpRs!LastUpdation
   End If
End If


CmdExcel.Enabled = False

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 14
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.WordWrap = True

'MSHFlexGrid1.TextMatrix(1, 0) = "Document Type"
'MSHFlexGrid1.TextMatrix(1, 1) = "Document No"
'MSHFlexGrid1.TextMatrix(1, 2) = "Narration"
'MSHFlexGrid1.TextMatrix(1, 3) = "Credit" '"Invoice Amount"
'MSHFlexGrid1.TextMatrix(1, 4) = "Debit" '"Base Amount"
'MSHFlexGrid1.TextMatrix(1, 5) = "Balance" '"Tax Amount"
''MSHFlexGrid1.TextMatrix(1, 6) = "Remaining Amount"
'MSHFlexGrid1.TextMatrix(1, 6) = "Customer Name"
'MSHFlexGrid1.TextMatrix(1, 7) = "Customer No"
'MSHFlexGrid1.TextMatrix(1, 8) = "Posting Date"
'MSHFlexGrid1.TextMatrix(1, 9) = "Voucher No"
'MSHFlexGrid1.TextMatrix(1, 10) = "EHOP Invoice No"
'MSHFlexGrid1.TextMatrix(1, 11) = "Actual Date Posted"
'MSHFlexGrid1.TextMatrix(1, 12) = "Actual Time Posted"
'MSHFlexGrid1.TextMatrix(1, 13) = "Customer Posting Group"
'MSHFlexGrid1.TextMatrix(1, 14) = "Entry No"

MSHFlexGrid1.TextMatrix(1, 0) = "Document Type"
MSHFlexGrid1.TextMatrix(1, 1) = "Document No"
MSHFlexGrid1.TextMatrix(1, 2) = "Narration"
MSHFlexGrid1.TextMatrix(1, 3) = "Credit" '"Invoice Amount"
MSHFlexGrid1.TextMatrix(1, 4) = "Debit" '"Base Amount"
'MSHFlexGrid1.TextMatrix(1, 5) = "Balance" '"Tax Amount"
'MSHFlexGrid1.TextMatrix(1, 6) = "Remaining Amount"
MSHFlexGrid1.TextMatrix(1, 5) = "Customer Name"
MSHFlexGrid1.TextMatrix(1, 6) = "Customer No"
MSHFlexGrid1.TextMatrix(1, 7) = "Posting Date"
MSHFlexGrid1.TextMatrix(1, 8) = "Voucher No"
MSHFlexGrid1.TextMatrix(1, 9) = "EHOP Invoice No"
MSHFlexGrid1.TextMatrix(1, 10) = "Actual Date Posted"
MSHFlexGrid1.TextMatrix(1, 11) = "Actual Time Posted"
MSHFlexGrid1.TextMatrix(1, 12) = "Customer Posting Group"
MSHFlexGrid1.TextMatrix(1, 13) = "Entry No"



MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1500
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1400
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 1400
'MSHFlexGrid1.ColWidth(14) = 1400

MSHFlexGrid1.RowHeight(1) = 800

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
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   Call Grid_Head
   Call Grid_Head1
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label27.Visible = True
   Label28.Visible = False
      Check3.Value = 0
   Check3.Enabled = False
Else
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   
   Check3.Enabled = True
   Call Check3_Click
End If
CmbNAVCode.Clear
CmbNAVCode.Text = ""

End Sub

Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then
   TxtCClientName = ""
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
   
End If
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Select Exchange!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
   CmbClientDP.SetFocus
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.NAVCode,Mst_Client.RNAVCode " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    If CmbClientDP.Text = "" Then
       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    End If
    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    TmpRs1.MoveNext
Next
If ans(3, 0) <= 1 Then
   CmbClientDP_LostFocus
End If
End Sub


Private Sub CmbCNonNcdexClient_GotFocus()

tmp1 = CmbCNonNcdexClient.Text
''Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)
tmp = "Select MC.ClientName,MC.ClientID,MD.DepositoryName,MC.ClientIDRow,MC.NAVCode,MC.RNAVCode From Mst_Client MC"
tmp = tmp & " inner join Mst_Depository MD on MD.DepositoryID=MC.DepositoryID"
tmp = tmp & " Where MC.ExchangeTypeId=" & mExchangeTypeID

Call TmpTable

CmbCNonNcdexClient.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbCNonNcdexClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientID & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientIDRow
    TmpRs.MoveNext
Next
CmbCNonNcdexClient.Text = tmp1
End Sub

Private Sub CmbCNonNcdexClient_LostFocus()
If CmbCNonNcdexClient.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbCNonNcdexClient, "~")

If UBound(strSplitString) <> 3 Then
   MsgBox "Invalid Selection!!!"
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If

ClientName_ = Trim(strSplitString(0))
ClientNo_ = Trim(strSplitString(1))
DPName_ = Trim(strSplitString(2))
ClientIDRow_ = Trim(strSplitString(3))

TmpRs.MoveFirst
TmpRs.Find "ClientID = " & ClientNo_ & ""

If TmpRs.EOF Then
   MsgBox "Invalid Selection "
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If

mClientIDRow = TmpRs!ClientID
mClientNAVCode = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)

If mClientNAVCode = "" And mClientRNAVCode = "" Then
   MsgBox "No NAV Code is mapped to selected Client!!!"
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If

CmbNAVCode.Clear
If mClientRNAVCode <> "" And mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.AddItem mClientRNAVCode
ElseIf mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.Text = mClientNAVCode
ElseIf mClientRNAVCode <> "" Then
   CmbNAVCode.AddItem mClientRNAVCode
   CmbNAVCode.Text = mClientRNAVCode
End If

'CmdGetReport.SetFocus

End Sub
Private Sub CmbClientDP_LostFocus()
Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      CmbClientDP.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtCClientName = TmpRs1!ClientName
TxtCDPID = TmpRs1!DepositoryID
TxtCDPName = TmpRs1!DepositoryName
mClientNAVCode = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs1!RNAVCode), "", TmpRs1!RNAVCode)
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow

CmbNAVCode.Clear

If mClientNAVCode = "" And mClientRNAVCode = "" Then
   MsgBox "No NAV Code is mapped to selected Client!!!"
   If CmbCNonNcdexClient.Visible = True Then
      CmbCNonNcdexClient.SetFocus
   Else
      TxtClientID.SetFocus
   End If
   Exit Sub
End If

If mClientRNAVCode <> "" And mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.AddItem mClientRNAVCode
ElseIf mClientNAVCode <> "" Then
   CmbNAVCode.AddItem mClientNAVCode
   CmbNAVCode.Text = mClientNAVCode
ElseIf mClientRNAVCode <> "" Then
   CmbNAVCode.AddItem mClientRNAVCode
   CmbNAVCode.Text = mClientRNAVCode
End If

End Sub

Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 5
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.TextMatrix(0, 0) = "Client ID"
MSHFlexGrid2.TextMatrix(0, 1) = "Client Name"
MSHFlexGrid2.TextMatrix(0, 2) = "Exchange Type"
MSHFlexGrid2.TextMatrix(0, 3) = "NAV Code"
MSHFlexGrid2.TextMatrix(0, 4) = "Reservation NAV Code"

MSHFlexGrid2.ColWidth(0) = 1800
MSHFlexGrid2.ColWidth(1) = 3000
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 1600
MSHFlexGrid2.ColWidth(4) = 1600

End Sub

Private Sub CmbDPName_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select Exchange Type!!!"
   CmbDPName.Clear
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
tmp = CmbDPName.Text

Call TableMst_Depository("Where ExchangeTypeId = " & mExchangeTypeID & " ")

CmbDPName.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDPName.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDPName.Text = tmp

End Sub

Private Sub CmbDPName_LostFocus()

Dim strSplitString() As String
Dim DPID_, DPName_ As String

If CmbDPName.Text = "" Then
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
End If
strSplitString = Split(CmbDPName, "~")
DPName_ = Trim(strSplitString(0))
DPID_ = Trim(strSplitString(1))

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDPName.SetFocus
   Exit Sub
End If

mDPID = RsMst_Depository!DepositoryID
TxtCDPID = mDPID
TxtCDPName = DPName_ 'CmbDPName.Text
mDPNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
If mDPNAVCode <> "" Then
   CmbNAVCode.AddItem mDPNAVCode
   CmbNAVCode.Text = mDPNAVCode
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbDPName.Visible = True
   CmbDPName.Width = 9400 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbCNonNcdexClient.Visible = False
   CmbDPName.Text = ""
   CmbNAVCode.Text = ""
Else
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbCNonNcdexClient.Visible = True
   CmbDPName.Visible = False
   CmbCNonNcdexClient.Text = ""
   CmbNAVCode.Text = ""
End If

End Sub
