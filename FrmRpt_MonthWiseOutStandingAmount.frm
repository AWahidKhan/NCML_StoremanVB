VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_MonthWiseOutStandingAmount 
   Caption         =   "Month Wise OutStanding Amount"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   15000
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   1800
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   13095
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
            Left            =   2370
            TabIndex        =   5
            Top             =   660
            Width           =   1140
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
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   2160
         End
         Begin VB.ComboBox CmbClient 
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
            Left            =   3720
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   9255
         End
         Begin VB.Label LblExchangeType 
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
            Left            =   435
            TabIndex        =   20
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label Label7 
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
            Left            =   7800
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Client"
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
         Top             =   1380
         Width           =   1485
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
         Left            =   6900
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   4335
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
         TabIndex        =   9
         Top             =   555
         Width           =   1095
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
         Left            =   11385
         TabIndex        =   7
         Top             =   555
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
         Left            =   12600
         TabIndex        =   8
         Top             =   555
         Width           =   1095
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
         Left            =   1740
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   570
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   570
         Width           =   1455
         _ExtentX        =   2566
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   61276163
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Starting Month"
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
         Left            =   210
         TabIndex        =   17
         Top             =   240
         Width           =   1260
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
         Left            =   8640
         TabIndex        =   16
         Top             =   240
         Width           =   855
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
         Left            =   11520
         TabIndex        =   15
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Left            =   13920
         TabIndex        =   14
         Top             =   240
         Width           =   75
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
         Left            =   3735
         TabIndex        =   13
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7035
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6615
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   11668
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
End
Attribute VB_Name = "FrmRpt_MonthWiseOutStandingAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mMonthCount  As Variant
Dim mClientRNAVCode, mClientNAVCode, mClientIDRow, mDepositoryID, mDPNAVCode, mExchangeTypeID  As Variant

Public Sub Grid_Head()
'MSHFlexGrid1.Clear
'If Check1.Value = 1 Then
'    MSHFlexGrid1.Cols = 20
'Else
'    MSHFlexGrid1.Cols = 18
'End If
'
'MSHFlexGrid1.Rows = 2
'MSHFlexGrid1.FixedRows = 1
'MSHFlexGrid1.FixedCols = 0
'
'MSHFlexGrid1.TextMatrix(0, 0) = "RM "
'MSHFlexGrid1.TextMatrix(0, 1) = "AM"
'MSHFlexGrid1.TextMatrix(0, 2) = "State"
'MSHFlexGrid1.TextMatrix(0, 3) = "Location"
'mMonth = TxtFromDate.Month
'mYear = TxtFromDate.Year
'
'If Check1.Value = 1 Then
'    MSHFlexGrid1.TextMatrix(0, 4) = "NAVCode"
'    MSHFlexGrid1.TextMatrix(0, 5) = "Client Name"
'    For c = 7 To 18
'        MSHFlexGrid1.TextMatrix(0, c) = Format(TxtFromDate.Value, "MMM-yyyy")
'        If TxtFromDate.Month = 12 Then
'            TxtFromDate.Month = 1
'            TxtFromDate.Year = TxtFromDate.Year + 1
'        Else
'            TxtFromDate.Month = TxtFromDate.Month + 1
'        End If
'    Next
'    MSHFlexGrid1.TextMatrix(0, 19) = "Total"
'Else
'    For c = 5 To 16
'        MSHFlexGrid1.TextMatrix(0, c) = Format(TxtFromDate.Value, "MMM-yyyy")
'        If TxtFromDate.Month = 12 Then
'            TxtFromDate.Month = 1
'            TxtFromDate.Year = TxtFromDate.Year + 1
'        Else
'            TxtFromDate.Month = TxtFromDate.Month + 1
'        End If
'    Next
'    MSHFlexGrid1.TextMatrix(0, 17) = "Total"
'End If
'
'TxtFromDate.Month = mMonth
'TxtFromDate.Year = mYear
'
''    MSHFlexGrid1.TextMatrix(0, 7) = "Jan-" & mYear
''MSHFlexGrid1.TextMatrix(0, 8) = "Feb-" & mYear
''MSHFlexGrid1.TextMatrix(0, 9) = "Mar-" & mYear
''MSHFlexGrid1.TextMatrix(0, 10) = "Apr-" & mYear
''MSHFlexGrid1.TextMatrix(0, 11) = "May-" & mYear
''MSHFlexGrid1.TextMatrix(0, 12) = "Jun-" & mYear
''MSHFlexGrid1.TextMatrix(0, 13) = "Jul-" & mYear
''MSHFlexGrid1.TextMatrix(0, 14) = "Aug-" & mYear
''MSHFlexGrid1.TextMatrix(0, 15) = "Sep-" & mYear
''MSHFlexGrid1.TextMatrix(0, 16) = "Oct-" & mYear
''MSHFlexGrid1.TextMatrix(0, 17) = "Nov-" & mYear
''MSHFlexGrid1.TextMatrix(0, 18) = "Dec-" & mYear
'
'MSHFlexGrid1.ColWidth(0) = 2500
'MSHFlexGrid1.ColWidth(1) = 2500
'MSHFlexGrid1.ColWidth(2) = 1800
'MSHFlexGrid1.ColWidth(3) = 1700
'MSHFlexGrid1.ColWidth(4) = 1800
'If Check1.Value = 1 Then
'    MSHFlexGrid1.ColWidth(5) = 2500
'    MSHFlexGrid1.ColWidth(6) = 2500
'Else
'    MSHFlexGrid1.ColWidth(5) = 1000
'    MSHFlexGrid1.ColWidth(6) = 1000
'End If
'MSHFlexGrid1.ColWidth(7) = 1000
'MSHFlexGrid1.ColWidth(8) = 1000
'MSHFlexGrid1.ColWidth(9) = 1000
'MSHFlexGrid1.ColWidth(10) = 1000
'MSHFlexGrid1.ColWidth(11) = 1000
'MSHFlexGrid1.ColWidth(12) = 1000
'MSHFlexGrid1.ColWidth(13) = 1000
'MSHFlexGrid1.ColWidth(14) = 1000
'MSHFlexGrid1.ColWidth(15) = 1000
'MSHFlexGrid1.ColWidth(16) = 1000
'MSHFlexGrid1.RowHeight(0) = 700


MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 5

MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.FixedCols = 0

MSHFlexGrid1.TextMatrix(1, 0) = "RM "
MSHFlexGrid1.TextMatrix(1, 1) = "AM"
MSHFlexGrid1.TextMatrix(1, 2) = "State"
MSHFlexGrid1.TextMatrix(1, 3) = "Location"
TxtFromDate.Day = 1
mMonth = TxtFromDate.Month
mYear = TxtFromDate.Year
    
If Check1.Value = 1 Then
    MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 2
    MSHFlexGrid1.TextMatrix(1, 4) = "NAVCode"
    MSHFlexGrid1.TextMatrix(1, 5) = "Client Name"
    For c = 7 To 18
        If TxtFromDate.Value > Date Then Exit For
        MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 1
        MSHFlexGrid1.TextMatrix(1, c) = Format(TxtFromDate.Value, "MMM-yyyy")
        If TxtFromDate.Month = 12 Then
            TxtFromDate.Month = 1
            TxtFromDate.Year = TxtFromDate.Year + 1
        Else
            TxtFromDate.Month = TxtFromDate.Month + 1
        End If
    Next
Else
    For c = 5 To 16
        If TxtFromDate.Value > Date Then Exit For
        MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 1
        MSHFlexGrid1.TextMatrix(1, c) = Format(TxtFromDate.Value, "MMM-yyyy")
        If TxtFromDate.Month = 12 Then
            TxtFromDate.Month = 1
            TxtFromDate.Year = TxtFromDate.Year + 1
        Else
            TxtFromDate.Month = TxtFromDate.Month + 1
        End If
    Next
End If
MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + 1
MSHFlexGrid1.TextMatrix(1, MSHFlexGrid1.Cols - 1) = "Total"
TxtFromDate.Month = mMonth
TxtFromDate.Year = mYear


MSHFlexGrid1.ColWidth(0) = 2500
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1700
MSHFlexGrid1.ColWidth(4) = 1800

If Check1.Value = 1 Then
    MSHFlexGrid1.ColWidth(5) = 2500
    MSHFlexGrid1.ColWidth(6) = 2500
    
Else
    MSHFlexGrid1.ColWidth(5) = 1100
    MSHFlexGrid1.ColWidth(6) = 1100
End If

For c = 7 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.ColWidth(c) = 1100
Next

End Sub


Private Sub Check1_Click()
If Check1.Value = 1 Then
    Frame3.Visible = True
    CmbExchangeTypeID.Text = ""
    Check3.Enabled = True
    CmbClient.Text = ""
Else
    Frame3.Visible = False
End If
End Sub

Private Sub CmbClient_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please Select Exchage Type !!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
If Check3.Value = 0 Then
    tmp1 = CmbClient.Text
    'Call TableMst_Client("Where ExchangeTypeID = " & mExchangeTypeID)   ' 'And ClientGroupID=" & mClientGroupID
    tmp = "Select MC.ClientName,MC.ClientID,MD.DepositoryName,MC.ClientIDRow,MC.NAVCode,MC.RNAVCode From Mst_Client MC"
    tmp = tmp & " inner join Mst_Depository MD on MD.DepositoryID=MC.DepositoryID"
    tmp = tmp & " Where MC.ExchangeTypeId=" & mExchangeTypeID

    Call TmpTable

    CmbClient.Clear
    If TmpRs.RecordCount = 0 Then
       MsgBox "No Record found !!!! "
       Exit Sub
    End If
    'For I = 1 To RsMst_Client.RecordCount
    '    'CmbCNonNcdexClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientID & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientIDRow
    '    CmbClient.AddItem RsMst_Client!ClientName & "~" & TmpRs!ClientID & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientIDRow
    '    RsMst_Client.MoveNext
    'Next
    For I = 1 To TmpRs.RecordCount
        CmbClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientID & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientIDRow
        TmpRs.MoveNext
    Next
    CmbClient.Text = tmp1
ElseIf Check3.Value = 1 Then
    tmp = CmbClient.Text
    Call TableMst_Depository("Where ExchangeTypeID = " & mExchangeTypeID)
    CmbClient.Clear
    If RsMst_Depository.RecordCount = 0 Then
        MsgBox "No record found !!!! "
        Exit Sub
    End If
    For I = 1 To RsMst_Depository.RecordCount
        CmbClient.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
        RsMst_Depository.MoveNext
    Next
    CmbClient.Text = tmp
End If
End Sub

Private Sub CmbClient_LostFocus()

If CmbClient.Text = "" Then
   Exit Sub
End If

If Check3.Value = 0 Then
   strSplitString = Split(CmbClient, "~")
   If UBound(strSplitString) <> 3 Then
      MsgBox "Invalid Selection!!!"
      CmbClient.SetFocus
      Exit Sub
   End If
    
   ClientName_ = Trim(strSplitString(0))
   ClientNo_ = Trim(strSplitString(1))
   DPName_ = Trim(strSplitString(2))
   ClientIDRow_ = Trim(strSplitString(3))
   TmpRs.MoveFirst
   TmpRs.Find "ClientIDRow = " & ClientIDRow_ & ""
   
   If TmpRs.EOF Then
       MsgBox "Invalid Selection "
       CmbClient.SetFocus
       Exit Sub
   End If
   mClientIDRow = TmpRs!ClientIDRow
   mClientNAVCode = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
   mClientRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)

ElseIf Check3.Value = 1 Then
    
   strSplitString = Split(CmbClient, "~")
   'DPName_ = Trim(strSplitString(0))
   DPID_ = Trim(strSplitString(1))
    
   RsMst_Depository.MoveFirst
   RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
   If RsMst_Depository.EOF Then
      MsgBox "Invalid selection "
      CmbClient.SetFocus
      Exit Sub
   End If
   mDepositoryID = RsMst_Depository!DepositoryID
   mDPNAVCode = IIf(IsNull(RsMst_Depository!NAVCode), "", RsMst_Depository!NAVCode)
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
   CmbClient.Clear
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   Check3.Enabled = False
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

If mExchangeTypeID = 1 Then
   Check3.Value = 0
   Check3.Enabled = False
Else
   Check3.Enabled = True
End If
End Sub

'Private Sub CmbClient_GotFocus()
'tmp = CmbClient.Text
'Call TableMst_Client
'CmbClient.Clear
'If RsMst_Client.RecordCount = 0 Then
'    MsgBox "No record found !!!! "
'    Exit Sub
'End If
'For I = 1 To RsMst_Client.RecordCount
'    CmbClient.AddItem RsMst_Client!ClientName
'    RsMst_Client.MoveNext
'Next
'CmbClient.Text = tmp
'End Sub
'
'Private Sub CmbClient_LostFocus()
'If CmbClient.Text = "" Then
'   Exit Sub
'End If
'RsMst_Client.MoveFirst
'RsMst_Client.Find "ClientName = '" & CmbClient.Text & "'"
'If RsMst_Client.EOF Then
'   MsgBox "Invalid selection "
'   CmbClient.SetFocus
'   Exit Sub
'End If
'mClientIDRow = RsMst_Client!ClientIDRow
'mNAVCode = RsMst_Client!NAVCode
'mRNavCode = RsMst_Client!RNavCode
'End Sub

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
CmbLocation.Clear
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
    CmbLocation.AddItem RsMst_Location!LocationName
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
Label5.Caption = ""
TxtFromDate.Value = Date - 30
'TxtFromDate.Value = Date - 30 'TxtFromDate.Year - 1
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbStateID.Text = ""
Call Grid_Head
Check1.Value = 0
Call Check1_Click
End Sub

Private Sub CmdExcel_Click()

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ClientLocationOutStandingInvoice.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ClientLocationOutStandingInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows

oWS.Range("a:a").ColumnWidth = 9
oWS.Range("b:b").ColumnWidth = 9
oWS.Range("c:c").ColumnWidth = 9
oWS.Range("d:d").ColumnWidth = 9
If Check1.Value = 1 Then
    oWS.Range("e:e").ColumnWidth = 10#
    oWS.Range("f:f").ColumnWidth = 11.29
    oWS.Range("g:g").ColumnWidth = 26.71
Else
    oWS.Range("e:e").ColumnWidth = 26.71
End If


For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = I + 1 ' + 6
    If (mRow + 2) Mod 5 = 0 Then 'If (mRow + 3) Mod 5 = 0 Then
        tmp = "a" & mRow & ":" & "a" & mRow + 4
        oWS.Range(tmp).Merge
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        tmp = "b" & mRow & ":" & "b" & mRow + 4
        oWS.Range(tmp).Merge
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        tmp = "c" & mRow & ":" & "c" & mRow + 4
        oWS.Range(tmp).Merge
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        tmp = "d" & mRow & ":" & "d" & mRow + 4
        oWS.Range(tmp).Merge
        oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        If Check1.Value = 1 Then
            tmp = "e" & mRow & ":" & "e" & mRow + 4
            oWS.Range(tmp).Merge
            oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
            tmp = "f" & mRow & ":" & "f" & mRow + 4
            oWS.Range(tmp).Merge
            oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
                
            For c = 0 To 5
               oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
               oWS.Cells(mRow, c + 1).WrapText = True
            Next
        Else
            For c = 0 To 3
               oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
               oWS.Cells(mRow, c + 1).WrapText = True
            Next
        End If
    End If
    If mRow = 1 Or mRow = 2 Then
       For c = 0 To MSHFlexGrid1.Cols - 1
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
       Next
    Else
       If Check1.Value = 1 Then
          For c = 6 To MSHFlexGrid1.Cols - 1
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
              oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
          Next
       Else
          For c = 4 To MSHFlexGrid1.Cols - 1
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
              oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
          Next
       End If
    End If
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String

Call Grid_Head
TxtFromDate.Day = 1

If TxtFromDate.Value > Date Then
    MsgBox "Future Date Not Allowed"
    TxtFromDate.SetFocus
    Exit Sub
End If

StartDate = Format(TxtFromDate.Value, "yyyy-MM-dd")

mdiff = DateDiff("M", TxtFromDate, Date)
        
If mdiff > 11 Then
    EndDate = Format(GetLastDateOfMonth(DateAdd("M", 11, TxtFromDate)), "yyyy-MM-dd")
Else
    EndDate = Format(GetLastDateOfMonth(Date), "yyyy-MM-dd")
End If


MsgBox "Please wait till next Message."

'StartDate = TxtFromDate.Year & "-01-01"
'EndDate = TxtFromDate.Year & "-12-31"

'If CDate(StartDate) > Date Then
'   MsgBox "Future date not allowed!!!"
'   TxtFromDate.SetFocus
'   Exit Sub
'End If


Label5.Caption = ""

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
'tmp = tmp & " TIDate, ML.LocationID,  MS.StateName, ML.LocationName,"
'tmp = tmp & " TI.NAVCode, TIAmount as OutStandingAmount, RecAmount as ReceiptAmount, RecTDSAmount as TDS, RecOtherDed as Adjustment,"
'tmp = tmp & " TIAmount-RecAmount-RecTDSAmount-RecOtherDed as BalanceOutStandingAmount,"
tmp = " Select ML.LocationID,  isnull(MS.StateName,'Unknown') as StateName, isnull(ML.LocationName,'Unknown') as LocationName"
If Check1.Value = 1 Then
    tmp = tmp & " , TI.NAVCode, (Select top 1 ClientName From Mst_Client Where NAVCode = TI.NAVCode Or RNAVCode = TI.NAVCode) as ClientName,"
    tmp = tmp & " (Select top 1 DepositoryName From Mst_Depository Where NAVCode = TI.NAVCode) as DepositoryName"
End If
tmp = tmp & " from Txn_TIDetails TI"
tmp = tmp & " left join Mst_Location ML ON ML.LocationID = TI.LocationID"
tmp = tmp & " left join Mst_State MS ON MS.StateID = ML.StateID "
tmp = tmp & wc
If wc = "" Then
    tmp = tmp & " Where TIAmount > IsNull(RecAmount+ AdjAmountFAG + RecTDSAmount + RecOtherDed, 0)"
Else
    tmp = tmp & " And TIAmount > IsNull(RecAmount + AdjAmountFAG + RecTDSAmount + RecOtherDed, 0)"
End If

If Check1.Value = 1 Then
    If CmbClient.Text <> "" Then
        If Check3.Value = 0 Then
           If mClientNAVCode <> "" Then
              tmp = tmp & " And (TI.NAVCode = '" & mClientNAVCode & "' ) " 'or MC.RNAVCode = '" & mClientRNAVCode & "')"
           End If
        ElseIf Check3.Value = 1 Then
           If mDPNAVCode <> "" Then
              tmp = tmp & " And TI.NAVCode='" & mDPNAVCode & "'"
           End If
        End If
    End If
End If


If Gen_DBType = "MDB" Then
    tmp = tmp & " And (TI.TIDate Between #" & StartDate & "# And #" & EndDate & "#) "
Else
    tmp = tmp & " And (TI.TIDate Between '" & StartDate & "' And '" & EndDate & "') "
End If
tmp = tmp & " group by ML.LocationID, MS.StateName, ML.LocationName"
If Check1.Value = 1 Then
    tmp = tmp & " , TI.NavCode "
End If
tmp = tmp & " Order by MS.StateName, ML.LocationName "

Call TmpTable

If TmpRs.RecordCount > 0 Then
    'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
    I = 2
    For k = 1 To TmpRs.RecordCount
        
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 5
        If IsNull(TmpRs!LocationID) = False Then
            Call GetAMRM(I)
        End If
        MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateName), "Unknown", TmpRs!StateName)
        MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!LocationName), "Unknown", TmpRs!LocationName)
        If Check1.Value = 1 Then
           MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!NAVCode), "", TmpRs!NAVCode)
           MSHFlexGrid1.TextMatrix(I, 6) = "Out Standing Invoice Amount"
           MSHFlexGrid1.TextMatrix(I + 1, 6) = "Receipt Amount"
           MSHFlexGrid1.TextMatrix(I + 2, 6) = "TDS"
           MSHFlexGrid1.TextMatrix(I + 3, 6) = "AdjustMent"
           MSHFlexGrid1.TextMatrix(I + 4, 6) = "Balance OutStanding Amount"
           
           mMonth = TxtFromDate.Month
           mYear = TxtFromDate.Year
           For l = 7 To 18
              If TxtFromDate.Value > Date Then Exit For
              Call GetDetail(I, l)
           Next
           TxtFromDate.Month = mMonth
           TxtFromDate.Year = mYear
           If IsNull(TmpRs!ClientName) = True Then
              MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
           Else
              MSHFlexGrid1.TextMatrix(I, 5) = TmpRs!ClientName
           End If
       Else
           MSHFlexGrid1.TextMatrix(I, 4) = "Out Standing Invoice Amount"
           MSHFlexGrid1.TextMatrix(I + 1, 4) = "Receipt Amount"
           MSHFlexGrid1.TextMatrix(I + 2, 4) = "TDS"
           MSHFlexGrid1.TextMatrix(I + 3, 4) = "AdjustMent"
           MSHFlexGrid1.TextMatrix(I + 4, 4) = "Balance OutStanding Amount"
           
           mMonth = TxtFromDate.Month
           mYear = TxtFromDate.Year
           For l = 5 To 16
               If TxtFromDate.Value > Date Then Exit For
               Call GetDetail(I, l)
           Next
           TxtFromDate.Month = mMonth
           TxtFromDate.Year = mYear
       End If
          
       MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Cols - 1) = Val(MSHFlexGrid1.TextMatrix(0, MSHFlexGrid1.Cols - 1)) + Val(MSHFlexGrid1.TextMatrix(I + 4, MSHFlexGrid1.Cols - 1))
       TmpRs.MoveNext
       
       I = I + 5
    Next
End If

Label5.Caption = TmpRs.RecordCount
Label5.Refresh
CmdExcel.Enabled = True
MsgBox "Done!!"
End Sub

Private Sub GetAMRM(x_)

tmp = " Select ML.LocationID,ML.LocationName, " ', EMP1.EmployeeID AS 'CCID', EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as 'CCName',"
tmp = tmp & " EMP2.EmployeeID as 'AMID',EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMName'"
tmp = tmp & " ,EMP3.EmployeeID as 'RMID',EMP3.EmployeeName +'~'+EMP3.EmployeeNo as 'RMName'"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID"
'tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
'tmp = tmp & " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD1 On EMP2.DesignationID = MD1.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD2 On EMP3.DesignationID = MD2.DesignationID"
tmp = tmp & " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID"
tmp = tmp & " Where ML.LocationID = " & TmpRs!LocationID
tmp = tmp & " Group By ML.LocationID,ML.LocationName, " ' EMP1.EmployeeID , EMP1.EmployeeName, EMP1.EmployeeNo ,"
tmp = tmp & " EMP2.EmployeeID , EMP2.EmployeeName, EMP2.EmployeeNo"
tmp = tmp & " ,EMP3.EmployeeID ,EMP3.EmployeeName ,EMP3.EmployeeNo"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 0) = IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName)
    MSHFlexGrid1.TextMatrix(x_, 1) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname)
End If
End Sub

Private Sub GetDetail(x_, l_)

Dim EndDate As Date

tmp = " Select"
tmp = tmp & " ML.LocationID,  MS.StateName, ML.LocationName,"
tmp = tmp & " sum(TIAmount) as OutStandingAmount, isnull(sum(RecAmount)+sum(AdjAmountFAG),0) as ReceiptAmount, isnull(sum(RecTDSAmount),0) as TDS, isnull(sum(RecOtherDed),0) as Adjustment,"
tmp = tmp & " sum(TIAmount)-isnull(sum(RecAmount)+sum(AdjAmountFAG)+sum(RecTDSAmount)+sum(RecOtherDed),0) as BalanceOutStandingAmount"
If Check1.Value = 1 Then
    tmp = tmp & " , TI.NAVCode"
End If
tmp = tmp & " from Txn_TIDetails TI"
tmp = tmp & " left join Mst_Location ML ON ML.LocationID = TI.LocationID"
tmp = tmp & " left join Mst_State MS ON MS.StateID = ML.StateID"
tmp = tmp & " Where ML.LocationID = " & IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID) & " "
StartDate = Format(TxtFromDate.Year & "-" & TxtFromDate.Month & "-01", "yyyy-MM-dd")

EndDate = Format(GetLastDateOfMonth(TxtFromDate.Value), "yyyy-MM-dd")
If Check1.Value = 1 Then
    tmp = tmp & " And TI.NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 4) & "' "
End If
tmp = tmp & " And TI.TIDate between '" & StartDate & "' And '" & EndDate & "' And TIAmount > IsNull(RecAmount + AdjAmountFAG + RecTDSAmount + RecOtherDed, 0) "
tmp = tmp & " group by ML.LocationID,  MS.StateName, ML.LocationName"
If Check1.Value = 1 Then
    tmp = tmp & " , TI.NAVCode"
End If
'tmp = tmp & " Having Sum(TIAmount) > IsNull(Sum(RecAmount) + Sum(RecTDSAmount) + Sum(RecOtherDed), 0)"

Call Tmp1Table


'If TmpRs1.RecordCount > 0 Then
'        If Check1.Value = 1 Then
'            MSHFlexGrid1.TextMatrix(x_, l_) = IIf(IsNull(TmpRs1!OutStandingAmount), "", TmpRs1!OutStandingAmount)
'            MSHFlexGrid1.TextMatrix(x_, 19) = Val(MSHFlexGrid1.TextMatrix(x_, 19)) + Val(MSHFlexGrid1.TextMatrix(x_, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 1, l_) = IIf(IsNull(TmpRs1!ReceiptAmount), "", TmpRs1!ReceiptAmount)
'            MSHFlexGrid1.TextMatrix(x_ + 1, 19) = Val(MSHFlexGrid1.TextMatrix(x_ + 1, 19)) + Val(MSHFlexGrid1.TextMatrix(x_ + 1, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 2, l_) = IIf(IsNull(TmpRs1!TDS), "", TmpRs1!TDS)
'            MSHFlexGrid1.TextMatrix(x_ + 2, 19) = Val(MSHFlexGrid1.TextMatrix(x_ + 2, 19)) + Val(MSHFlexGrid1.TextMatrix(x_ + 2, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 3, l_) = IIf(IsNull(TmpRs1!Adjustment), "", TmpRs1!Adjustment)
'            MSHFlexGrid1.TextMatrix(x_ + 3, 19) = Val(MSHFlexGrid1.TextMatrix(x_ + 3, 19)) + Val(MSHFlexGrid1.TextMatrix(x_ + 3, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 4, l_) = IIf(IsNull(TmpRs1!BalanceOutStandingAmount), "", TmpRs1!BalanceOutStandingAmount)
'            MSHFlexGrid1.TextMatrix(x_ + 4, 19) = Val(MSHFlexGrid1.TextMatrix(x_ + 4, 19)) + Val(MSHFlexGrid1.TextMatrix(x_ + 4, l_))
'        Else
'
'            MSHFlexGrid1.TextMatrix(x_, l_) = IIf(IsNull(TmpRs1!OutStandingAmount), "", TmpRs1!OutStandingAmount)
'            MSHFlexGrid1.TextMatrix(x_, 17) = Val(MSHFlexGrid1.TextMatrix(x_, 17)) + Val(MSHFlexGrid1.TextMatrix(x_, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 1, l_) = IIf(IsNull(TmpRs1!ReceiptAmount), "", TmpRs1!ReceiptAmount)
'            MSHFlexGrid1.TextMatrix(x_ + 1, 17) = Val(MSHFlexGrid1.TextMatrix(x_ + 1, 17)) + Val(MSHFlexGrid1.TextMatrix(x_ + 1, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 2, l_) = IIf(IsNull(TmpRs1!TDS), "", TmpRs1!TDS)
'            MSHFlexGrid1.TextMatrix(x_ + 2, 17) = Val(MSHFlexGrid1.TextMatrix(x_ + 2, 17)) + Val(MSHFlexGrid1.TextMatrix(x_ + 2, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 3, l_) = IIf(IsNull(TmpRs1!Adjustment), "", TmpRs1!Adjustment)
'            MSHFlexGrid1.TextMatrix(x_ + 3, 17) = Val(MSHFlexGrid1.TextMatrix(x_ + 3, 17)) + Val(MSHFlexGrid1.TextMatrix(x_ + 3, l_))
'            MSHFlexGrid1.TextMatrix(x_ + 4, l_) = IIf(IsNull(TmpRs1!BalanceOutStandingAmount), "", TmpRs1!BalanceOutStandingAmount)
'            MSHFlexGrid1.TextMatrix(x_ + 4, 17) = Val(MSHFlexGrid1.TextMatrix(x_ + 4, 17)) + Val(MSHFlexGrid1.TextMatrix(x_ + 4, l_))
'        End If
'End If

If TmpRs1.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, l_) = IIf(IsNull(TmpRs1!OutStandingAmount), "", TmpRs1!OutStandingAmount)
    MSHFlexGrid1.TextMatrix(x_, MSHFlexGrid1.Cols - 1) = Val(MSHFlexGrid1.TextMatrix(x_, MSHFlexGrid1.Cols - 1)) + Val(MSHFlexGrid1.TextMatrix(x_, l_))
    MSHFlexGrid1.TextMatrix(x_ + 1, l_) = IIf(IsNull(TmpRs1!ReceiptAmount), "", TmpRs1!ReceiptAmount)
    MSHFlexGrid1.TextMatrix(x_ + 1, MSHFlexGrid1.Cols - 1) = Val(MSHFlexGrid1.TextMatrix(x_ + 1, MSHFlexGrid1.Cols - 1)) + Val(MSHFlexGrid1.TextMatrix(x_ + 1, l_))
    MSHFlexGrid1.TextMatrix(x_ + 2, l_) = IIf(IsNull(TmpRs1!TDS), "", TmpRs1!TDS)
    MSHFlexGrid1.TextMatrix(x_ + 2, MSHFlexGrid1.Cols - 1) = Val(MSHFlexGrid1.TextMatrix(x_ + 2, MSHFlexGrid1.Cols - 1)) + Val(MSHFlexGrid1.TextMatrix(x_ + 2, l_))
    MSHFlexGrid1.TextMatrix(x_ + 3, l_) = IIf(IsNull(TmpRs1!Adjustment), "", TmpRs1!Adjustment)
    MSHFlexGrid1.TextMatrix(x_ + 3, MSHFlexGrid1.Cols - 1) = Val(MSHFlexGrid1.TextMatrix(x_ + 3, MSHFlexGrid1.Cols - 1)) + Val(MSHFlexGrid1.TextMatrix(x_ + 3, l_))
    MSHFlexGrid1.TextMatrix(x_ + 4, l_) = IIf(IsNull(TmpRs1!BalanceOutStandingAmount), "", TmpRs1!BalanceOutStandingAmount)
    MSHFlexGrid1.TextMatrix(x_ + 4, MSHFlexGrid1.Cols - 1) = Val(MSHFlexGrid1.TextMatrix(x_ + 4, MSHFlexGrid1.Cols - 1)) + Val(MSHFlexGrid1.TextMatrix(x_ + 4, l_))
    MSHFlexGrid1.TextMatrix(0, l_) = Val(MSHFlexGrid1.TextMatrix(x_ + 4, l_)) + Val(MSHFlexGrid1.TextMatrix(0, l_))
End If


If TxtFromDate.Month = 12 Then
    TxtFromDate.Month = 1
    TxtFromDate.Year = TxtFromDate.Year + 1
Else
    TxtFromDate.Month = TxtFromDate.Month + 1
End If

End Sub

Private Function GetLastDateOfMonth(mdate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mdate_)
mSYear = Year(mdate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
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
CmdClear_Click
CmdExcel.Enabled = False
End Sub

