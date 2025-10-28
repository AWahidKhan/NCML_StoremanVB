VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_ClientLocationOutStandingInvoice 
   Caption         =   "Client-Location Wise OutStanding Invoice"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7155
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   150
         Left            =   120
         TabIndex        =   19
         Top             =   6885
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6390
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   11271
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
      Height          =   2145
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   15000
      Begin VB.CheckBox ChkBatchNo 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   24
         Top             =   1680
         Value           =   1  'Checked
         Width           =   225
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
         Left            =   3440
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.CheckBox ChkDP 
         Caption         =   "Show DP"
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
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   1253
         Width           =   1125
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   1950
      End
      Begin VB.TextBox TxtNAVCode 
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
         Left            =   9240
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox CmbClientID 
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
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   5775
      End
      Begin VB.CheckBox ChkWithStock 
         Caption         =   "With Stock Details"
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
         Left            =   11340
         TabIndex        =   7
         Top             =   960
         Width           =   1965
      End
      Begin VB.ComboBox CmbStateID 
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
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   500
         Width           =   5775
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
         Left            =   12480
         TabIndex        =   9
         Top             =   493
         Width           =   1215
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
         Left            =   11265
         TabIndex        =   8
         Top             =   480
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
         Left            =   13695
         TabIndex        =   10
         Top             =   493
         Width           =   1215
      End
      Begin VB.ComboBox CmbLocation 
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
         Left            =   6000
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   500
         Width           =   5175
      End
      Begin VB.Label Label30 
         Caption         =   "Approved By FAG"
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
         TabIndex        =   25
         Top             =   1680
         Width           =   1710
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DP Name"
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
         Left            =   5760
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   383
         TabIndex        =   22
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label2 
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
         Left            =   9720
         TabIndex        =   21
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   5640
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   2535
         TabIndex        =   15
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label5 
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
         Left            =   13560
         TabIndex        =   14
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label4 
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
         Left            =   11280
         TabIndex        =   13
         Top             =   1320
         Width           =   2325
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   8160
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
      Height          =   4335
      Left            =   1800
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7646
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
Attribute VB_Name = "FrmRpt_ClientLocationOutStandingInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID As Variant

'''-- Ullhas Patil 16 March 2018
Dim mClientIDRow As Variant, mClientNAVCode As Variant, mExchangeTypeID As Variant, mDPID As Variant, mDPNAVCode As Variant

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 7
    .Rows = 3
    .FixedRows = 2
    .FixedCols = 0
    .WordWrap = True
    
    .TextMatrix(1, 0) = "State Name":               .ColWidth(0) = 2000
    .TextMatrix(1, 1) = "Location Name":            .ColWidth(1) = 2500
    .TextMatrix(1, 2) = "NAVCode":                  .ColWidth(2) = 1800
    .TextMatrix(1, 3) = "Client Name":              .ColWidth(3) = 5500
    .TextMatrix(1, 4) = "TI Count":                 .ColWidth(4) = 1000
    .TextMatrix(1, 5) = "OutStanding Amount":       .ColWidth(5) = 1500
    .TextMatrix(1, 6) = "Stock Value":              .ColWidth(6) = 1500

    .RowHeight(1) = 700
End With
End Sub

Private Sub ChkDP_Click()
If ChkDP.Value = vbUnchecked Then
   Label1.Visible = True
   CmbClientID.Visible = True
   Label8.Visible = False
   CmbDPName.Visible = False
Else
   Label1.Visible = False
   CmbClientID.Visible = False
   Label8.Visible = True
   CmbDPName.Visible = True
End If
End Sub

Private Sub CmbClientId_GotFocus()
If CmbExchangeTypeID.Text = "" Then
    MsgBox "Select Exchange"
    CmbExchangeTypeID.SetFocus
    Exit Sub
End If

tmp = CmbClientID.Text
Call TableMst_Client("Where ExchangeTypeID = " & Val(mExchangeTypeID) & "")
CmbClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClientID.AddItem RsMst_Client!ClientName & "#" & RsMst_Client!ClientIDRow & "#" & RsMst_Client!ClientID
    RsMst_Client.MoveNext
Next
CmbClientID.Text = tmp
End Sub
Private Sub CmbClientId_LostFocus()
Dim mSplit As Variant
Dim mR As Double

If CmbClientID.Text = "" Then
   mClientNAVCode = ""
   Exit Sub
End If

mSplit = Split(CmbClientID, "#")
If UBound(mSplit) <> 2 Then
   MsgBox "Invalid selection "
   CmbClientID.SetFocus
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientID = '" & mSplit(2) & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbClientID.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID
mClientNAVCode = "''"
mClientNAVCode = mClientNAVCode & IIf(IsNull(RsMst_Client!NavCode), ",''", ",'" & RsMst_Client!NavCode & "'") & IIf(IsNull(RsMst_Client!RNAVCode), ",''", ",'" & RsMst_Client!RNAVCode & "'")

tmp = " SELECT GSTNavCode,ClientID,RGSTNavCode FROM Mst_ClientStateWiseGST Where ClientID = " & mClientIDRow & " "
Call TmpTable
If TmpRs.RecordCount > 0 Then
    For mR = 1 To TmpRs.RecordCount
        mClientNAVCode = mClientNAVCode & IIf(IsNull(TmpRs!GSTNavCode), ",''", ",'" & TmpRs!GSTNavCode & "'") & IIf(IsNull(TmpRs!RGSTNavCode), ",''", ",'" & TmpRs!RGSTNavCode & "'")
    Next mR
End If
TxtNAVCode = ""
End Sub

Private Sub CmbDPName_GotFocus()
If CmbExchangeTypeID.Text = "" Then
    MsgBox "Select Exchange"
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
Dim mR As Double

If CmbDPName.Text = "" Then
   mDPNAVCode = ""
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
mDPNAVCode = "''"
mDPID = RsMst_Depository!DepositoryID
mDPNAVCode = mDPNAVCode & IIf(IsNull(RsMst_Depository!NavCode), ",''", ",'" & RsMst_Depository!NavCode & "'")

tmp = " SELECT GSTNavCode,DepositoryID FROM Mst_DepositoryStateWiseGST Where DepositoryID = '" & mDPID & "'"
Call TmpTable
If TmpRs.RecordCount > 0 Then
    For mR = 1 To TmpRs.RecordCount
        mDPNAVCode = mDPNAVCode & IIf(IsNull(TmpRs!GSTNavCode), ",''", ",'" & TmpRs!GSTNavCode & "'")
    Next mR
End If
TxtNAVCode = ""
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
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   ChkDP.Value = vbUnchecked: ChkDP.Enabled = False
   Label1.Visible = True
   CmbClientID = "": CmbClientID.Visible = True
   Label8.Visible = False
   CmbDPName = "": CmbDPName.Visible = False
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   ChkDP.Value = vbUnchecked: ChkDP.Enabled = False
   Label1.Visible = True
   CmbClientID = "": CmbClientID.Visible = True
   Label8.Visible = False
   CmbDPName = "": CmbDPName.Visible = False
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
If mExchangeTypeID = 1 Then         ''-- Non-NCDEX
   ChkDP.Value = vbUnchecked: ChkDP.Enabled = False
ElseIf mExchangeTypeID = 2 Then     ''-- NCDEX
   ChkDP.Value = vbUnchecked: ChkDP.Enabled = True
Else
    ChkDP.Value = vbUnchecked: ChkDP.Enabled = False
End If

Label1.Visible = True
CmbClientID = "": CmbClientID.Visible = True
Label8.Visible = False
CmbDPName = "": CmbDPName.Visible = False

End Sub


Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" where " & Gen_WcState & "")
Else
   Call TableMst_State
End If

CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocation.Clear
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.Clear
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmbLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State !!"
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text

Call TableMst_Location(" Where LocationID <> 0 And StateID= " & mStateID & "")

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_LostFocus()

If CmbLocation.Text = "" Then
   mLocationID = 0
   TxtPreFix = ""
   Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocation.Text <> tmp Then
   TxtPreFix = ""
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Label5.Caption = ""
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbStateID.Text = ""
ProgressBar1.Value = 0
ChkWithStock.Value = 0
''-- Ullhas Patil 16 March 2018
CmbExchangeTypeID.Text = ""
ChkDP.Value = vbUnchecked: ChkDP.Enabled = False

Label1.Visible = True
CmbClientID.Text = "": CmbClientID.Visible = True
Label8.Visible = False
CmbDPName.Text = "": CmbDPName.Visible = False

TxtNAVCode.Text = ""
''----------
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ClientLocationOutStandingInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ClientLocationOutStandingInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
       oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
Dim mWC As String

MsgBox "Please wait till next Message."

Call Grid_Head
ProgressBar1.Value = 0
Label5.Caption = ""

If TxtNAVCode <> "" Then
    mWC = " Where A.NAVCode = '" & TxtNAVCode & "' "
Else
    If CmbExchangeTypeID <> "" Then
        If ChkDP.Value = vbChecked Then
            If CmbDPName.Text <> "" Then
                mWC = " Where A.NAVCode In (" & mDPNAVCode & ") "
            End If
        Else
            If CmbClientID.Text <> "" Then
                mWC = " Where A.NAVCode In (" & mClientNAVCode & ") "
            End If
        End If
    End If
End If

wc = " Where TINo<>''  And "
If ChkBatchNo.Value = vbChecked Then
   wc = wc & " TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y' And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+ (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R' And TCR.BatchID > 0)+AdjAmountFAG)>0 "
Else
   wc = wc & " TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+ (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R')+AdjAmountFAG)>0 "
End If

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp = Replace(Gen_WcState, "StateID", "MS.StateID")
      wc = GenWc(wc, tmp, "", "N", "I")
   End If
End If

'tmp = " Select"
'tmp = tmp & " isnull(MS.StateName,'Unknown') as StateName, isnull(ML.LocationName,'Unknown') as LocationName, count(*) as TICount,"
'tmp = tmp & " TI.NAVCode, sum(TIAmount) as TIAmount , isnull(sum(RecAmount)+sum(AdjAmountFAG)+sum(RecTDSAmount)+sum(RecOtherDed),0) as ReceivedAmount,"
'tmp = tmp & " sum(TIAmount) - isnull(sum(RecAmount)+sum(AdjAmountFAG)+sum(RecTDSAmount)+sum(RecOtherDed),0) as OutStandingAmount, "
'tmp = tmp & " (Select top 1 ClientName From Mst_Client Where NAVCode = TI.NAVCode Or RNAVCode = TI.NAVCode) as ClientName,"
'tmp = tmp & " (Select top 1 DepositoryName From Mst_Depository Where NAVCode = TI.NAVCode) as DepositoryName"
'tmp = tmp & " from Txn_TIDetails TI"
'tmp = tmp & " left join Mst_Location ML ON ML.LocationID = TI.LocationID"
'tmp = tmp & " left join Mst_State MS ON MS.StateID = ML.StateID "
'tmp = tmp & wc
'tmp = tmp & " group by MS.StateName, ML.LocationName, TI.NAVCode, ML.LocationID"
'tmp = tmp & " Having Sum(TIAmount) > IsNull(Sum(RecAmount)+ sum(AdjAmountFAG) + Sum(RecTDSAmount) + Sum(RecOtherDed), 0)"
'tmp = tmp & " order by MS.StateName, ML.LocationName"

tmp = " Select StateName, LocationName, Count(TIID) TICount, NAVCode, Sum(TIAmount) TIAmount, sum(RecAmount) ReceivedAmount,"
tmp = tmp & " sum(TIAmount - RecAmount) OutStandingAmount, ClientName, DepositoryName"
tmp = tmp & " From ("
tmp = tmp & " Select MS.StateName, ML.LocationName, TI.TIID, TI.NAVCode, TI.TIAmount,"
If ChkBatchNo.Value = vbChecked Then
   tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+ isnull(TI.AdjAmountFAG,0) +"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0) +"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.BatchID > 0 )+"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.FAGPostingDate Is Not Null And TCR.BatchID > 0)+"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.BatchID > 0) as RecAmount,"
Else
   tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ isnull(TI.AdjAmountFAG,0) +"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) +"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )+"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+"
   tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID) as RecAmount,"
End If
'''-- Ullhas Patil 18 Sep 2017 ''-- Change the query to get the Client Name / Depositor Name from State Wise GSTNavCode
'''tmp = tmp & " (Select top 1 ClientName From Mst_Client Where NAVCode = TI.NAVCode Or RNAVCode = TI.NAVCode) as ClientName,"
'''tmp = tmp & " (Select top 1 DepositoryName From Mst_Depository Where NAVCode = TI.NAVCode) as DepositoryName"
tmp = tmp & " (SELECT top 1 MC.ClientName FROM Mst_ClientStateWiseGST MCGST Inner Join Mst_Client MC On MCGST.ClientID = MC.ClientID Where MCGST.GSTNavCode = TI.NAVCode Or MCGST.RGSTNavCode = TI.NAVCode ) as ClientName, "
tmp = tmp & " (SELECT top 1 MD.DepositoryName FROM Mst_DepositoryStateWiseGST MDGST Inner Join Mst_Depository MD On MDGST.DepositoryID = MD.DepositoryID And MDGST.ExchangeTypeID = MD.ExchangeTypeID Where MDGST.GSTNavCode = TI.NAVCode ) as DepositoryName "
'''---------
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID "
tmp = tmp & wc

'If wc = "" Then
'   tmp = tmp & " Where"
'Else
'   tmp = tmp & " And "
'End If
'tmp = tmp & " TINo<>''  And"
'tmp = tmp & " TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+ (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID = TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R')+AdjAmountFAG)>0"

tmp = tmp & " group by ML.LocationName, MS.StateName, TI.TIID, TI.NAVCode, TI.TIAmount, TI.AdjAmountFAG, TI.NAVCode"
tmp = tmp & " )A "

If mWC <> "" Then
    tmp = tmp & mWC
End If

tmp = tmp & " Group By StateName, LocationName, ClientName, DepositoryName, NAVCode"
tmp = tmp & " Order by StateName, LocationName, ClientName, DepositoryName, NAVCode"

Call TmpTable

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
    For I = 2 To TmpRs.RecordCount + 1
        MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
        
        If IsNull(TmpRs!ClientName) = True Then
            MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
        Else
            MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!ClientName
        End If
        
        If Trim(MSHFlexGrid1.TextMatrix(I, 3)) = "" Then
            Call GetClientName(I)
        End If
        
        MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!TICount), 0, TmpRs!TICount)
        MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
        MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!OutStandingAmount), "", TmpRs!OutStandingAmount)
        MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
        If IsNull(TmpRs!Locationname) = False And IsNull(TmpRs!NavCode) = False Then
           If ChkWithStock.Value = 1 Then
              MSHFlexGrid1.TextMatrix(I, 6) = FillGSLBalanceStockWithPrice(TmpRs!NavCode, TmpRs!Locationname, GetStateID(TmpRs!StateName))
           End If
        End If
        TmpRs.MoveNext
        ProgressBar1.Value = ProgressBar1.Value + 1
    Next
End If
ProgressBar1.Value = ProgressBar1.Max
Label5.Caption = TmpRs.RecordCount
Label5.Refresh
CmdExcel.Enabled = True
MsgBox "Done!!"
End Sub

Private Sub GetClientName(x_ As Variant)
tmp = " Select ClientName From Mst_Client where NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "' Or RNAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "'"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 3) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
Else
    tmp = " Select DepositoryName From Mst_Depository where NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 2) & "' "
    If TmpRs1.RecordCount > 0 Then
        MSHFlexGrid1.TextMatrix(x_, 3) = IIf(IsNull(TmpRs1!DepositoryName), "", TmpRs1!DepositoryName)
    End If
End If
End Sub


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
Grid_Head
CmdExcel.Enabled = False
End Sub


Private Function FillGSLBalanceStockWithPrice(mClientNAVCode_ As Variant, mLocName_ As Variant, mStateID_ As Variant) As Double
Dim mSM_Prefix_  As String
Call Grid_Head_Stock
'Call TableTxn_CxTF_GSL("  Where  sm_prefix = '" & TxtPreFix & "' And CxTFNo = " & TxtCxTFNo & " And TxnType = 'W'")
'Call TableMst_GSL("  Where  sm_prefix = '" & TxtPreFix & "' And ClientIDRow in (Select ClientIDRow from Txn_CxTF_ClientDetail Where SM_Prefix='" & TxtPreFix & "' And CxTFNo=" & Val(TxtCxTFNo) & " And TxnType='W') And BalanceWeight>0")
tmp = "select SM_Prefix from Mst_location Where LocationName='" & mLocName_ & "'"
Call Tmp3Table
If TmpRs3.RecordCount > 0 Then
   mSM_Prefix_ = TmpRs3!SM_Prefix
Else
   mSM_Prefix_ = ""
End If
tmp = "Select GSL.CommodityID,MC.Commodity,sum(GSL.BalanceBags)BalanceBags,sum(GSL.BalanceWeight)BalanceWeight,MCL.ClientIDRow+'~'+MCL.ClientName as ClientName"
tmp = tmp & " from Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID=GSL.GodownID"
tmp = tmp & " Inner Join Mst_Commodity MC on MC.CommodityID=GSL.CommodityID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientID=GSL.ClientID"
tmp = tmp & " Where  GSL.ExchangeTypeID in (1,2) And GSL.sm_prefix = '" & mSM_Prefix_ & "' And GSL.ClientID "
''--- Ullhas Patil 18 Sep 2017 Query change due to New GSTNAV Code Mapping
tmp = tmp & " in ((Select ClientID from Mst_Client Where NAVCode ='" & mClientNAVCode_ & "' or RNAVCode ='" & mClientNAVCode_ & "') Union "
tmp = tmp & " (Select MCGST.ClientID from Mst_ClientStateWiseGST MCGST Inner Join Mst_Client MC On MCGST.ClientID = MC.ClientID Where MCGST.GSTNavCode ='" & mClientNAVCode_ & "' or MCGST.RGSTNavCode ='" & mClientNAVCode_ & "')) "
''--------
tmp = tmp & " And GSL.Status not in ('E','Z') And GSL.BalanceBags > 0 And GSL.BalanceWeight > 0"
tmp = tmp & " Group by MG.GodownNo,GSL.CommodityID,MC.Commodity,MCL.ClientIDRow,MCL.ClientName"

Call Tmp5Table

MSHFlexGrid8.Rows = MSHFlexGrid8.Rows + TmpRs5.RecordCount

If TmpRs5.RecordCount = 0 Then Exit Function

For I = 2 To MSHFlexGrid8.Rows - 2
    MSHFlexGrid8.TextMatrix(I, 1) = TmpRs5!Commodity
    MSHFlexGrid8.TextMatrix(I, 2) = TmpRs5!ClientName
    MSHFlexGrid8.TextMatrix(I, 3) = TmpRs5!BalanceBags
    MSHFlexGrid8.TextMatrix(I, 4) = TmpRs5!BalanceWeight
    mCommPrice = GetCommodityLatestRate(TmpRs5!CommodityID, mStateID_)
    MSHFlexGrid8.TextMatrix(I, 5) = mCommPrice
    MSHFlexGrid8.TextMatrix(I, 6) = Format(mCommPrice * TmpRs5!BalanceWeight, "#####0.00")
    MSHFlexGrid8.TextMatrix(0, 3) = Val(MSHFlexGrid8.TextMatrix(0, 3)) + Val(MSHFlexGrid8.TextMatrix(I, 3))
    MSHFlexGrid8.TextMatrix(0, 4) = Val(MSHFlexGrid8.TextMatrix(0, 4)) + Val(MSHFlexGrid8.TextMatrix(I, 4))
    MSHFlexGrid8.TextMatrix(0, 6) = Val(MSHFlexGrid8.TextMatrix(0, 6)) + Val(MSHFlexGrid8.TextMatrix(I, 6))
    TmpRs5.MoveNext
Next

FillGSLBalanceStockWithPrice = Val(MSHFlexGrid8.TextMatrix(0, 6))

End Function

Private Sub Grid_Head_Stock()

MSHFlexGrid8.Clear
MSHFlexGrid8.Rows = 3
MSHFlexGrid8.FixedRows = 2
MSHFlexGrid8.Cols = 7

MSHFlexGrid8.TextMatrix(1, 0) = "Godown No"
MSHFlexGrid8.TextMatrix(1, 1) = "Commodity"
MSHFlexGrid8.TextMatrix(1, 2) = "Client Name"
MSHFlexGrid8.TextMatrix(1, 3) = "Balance Bags"
MSHFlexGrid8.TextMatrix(1, 4) = "Balance Weight"
MSHFlexGrid8.TextMatrix(1, 5) = "Commodity Price"
MSHFlexGrid8.TextMatrix(1, 6) = "Total Price"

MSHFlexGrid8.ColWidth(0) = 1200
MSHFlexGrid8.ColWidth(1) = 1400
MSHFlexGrid8.ColWidth(2) = 2500
MSHFlexGrid8.ColWidth(3) = 1400
MSHFlexGrid8.ColWidth(4) = 1400
MSHFlexGrid8.ColWidth(5) = 1400
MSHFlexGrid8.ColWidth(6) = 1400

MSHFlexGrid8.RowHeight(1) = 600

End Sub


Private Function GetCommodityLatestRate(mCommodityID_ As Variant, mState_ As Variant) As Variant

Dim mComRate, mStateID As Variant

'tmp = " Select StateID from Mst_Location Where LocationID= " & mLocationID & ""
'Call Tmp2Table
'mStateID = 0
'If TmpRs2.RecordCount > 0 Then
'   mStateID = TmpRs2!StateID
'End If

'tmp = " Select MCP.YearMonth, MCP.Day28,MCP.Day29,MCP.Day30,MCP.Day31 from Mst_CommodityPrice MCP"
'tmp = tmp & " Where MCP.StateID = " & mState_ & " And MCP.CommodityID = " & mCommodityID_ & ""
'tmp = tmp & " Order By MCP.YearMonth "

tmp = " Select MCP.StateID,MCP.CommodityID,MCP.YearMonth,"
tmp = tmp & " Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
tmp = tmp & " Then"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " Else"
tmp = tmp & "   COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & " End [Rate]"
tmp = tmp & " from Mst_CommodityPrice MCP"
tmp = tmp & " Where StateID = " & mState_ & " And CommodityID = " & mCommodityID_ & "  And Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Group By StateID,CommodityID )"


Call Tmp2Table

mComRate = 0

'If TmpRs2.RecordCount > 0 Then
'   TmpRs2.MoveLast
'   If Right(TmpRs2!YearMonth, 2) = "01" Or Right(TmpRs2!YearMonth, 2) = "03" Or Right(TmpRs2!YearMonth, 2) = "05" Or Right(TmpRs2!YearMonth, 2) = "07" Or Right(TmpRs2!YearMonth, 2) = "08" Or Right(TmpRs2!YearMonth, 2) = "10" Or Right(TmpRs2!YearMonth, 2) = "12" Then
'      mComRate = TmpRs2!Day31
'   ElseIf Right(TmpRs2!YearMonth, 2) = "04" Or Right(TmpRs2!YearMonth, 2) = "06" Or Right(TmpRs2!YearMonth, 2) = "09" Or Right(TmpRs2!YearMonth, 2) = "11" Then
'      mComRate = TmpRs2!Day30
'   Else
'      If IsNull(TmpRs2!Day29) = False Then
'         mComRate = TmpRs2!Day29
'      Else
'         mComRate = TmpRs2!Day28
'      End If
'   End If
'End If

If TmpRs2.RecordCount > 0 Then
   mComRate = IIf(IsNull(TmpRs2!Rate), 0, TmpRs2!Rate)
End If

GetCommodityLatestRate = mComRate
End Function

