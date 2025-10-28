VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_TruckwiseCxtfDetail 
   Caption         =   "Truck wise Cxtf Detail"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   9855
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   14895
      Begin VB.Frame Frame1 
         Height          =   9615
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   14655
         Begin VB.ComboBox CmbTransactiontype 
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
            ItemData        =   "FrmRpt_TruckwiseCxtfDetail.frx":0000
            Left            =   11880
            List            =   "FrmRpt_TruckwiseCxtfDetail.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   1245
            Width           =   2535
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
            ItemData        =   "FrmRpt_TruckwiseCxtfDetail.frx":0021
            Left            =   9000
            List            =   "FrmRpt_TruckwiseCxtfDetail.frx":0023
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   1245
            Width           =   2775
         End
         Begin VB.ComboBox CmbAgent 
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
            Left            =   3000
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   1245
            Width           =   5910
         End
         Begin VB.ComboBox CmbCommodity 
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
            TabIndex        =   6
            Top             =   1245
            Width           =   2760
         End
         Begin VB.ComboBox CmbCMP 
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
            Left            =   9000
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   487
            Width           =   5460
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
            Left            =   13230
            TabIndex        =   12
            Top             =   1680
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
            Height          =   360
            Left            =   12015
            TabIndex        =   11
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton CmdGo 
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
            Height          =   360
            Left            =   10800
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1680
            Width           =   1215
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
            Left            =   3000
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   487
            Width           =   2775
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
            Left            =   5880
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   487
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   480
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
            Format          =   110886913
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7065
            Left            =   120
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2100
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   12462
            _Version        =   393216
            FixedCols       =   0
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
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   375
            Left            =   1530
            TabIndex        =   2
            Top             =   480
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
            Format          =   110886913
            CurrentDate     =   36494
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   9240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Transaction Type"
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
            Left            =   12352
            TabIndex        =   31
            Top             =   960
            Width           =   1590
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
            Left            =   9675
            TabIndex        =   30
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agent"
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
            Left            =   5693
            TabIndex        =   29
            Top             =   960
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Commodity"
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
            Left            =   990
            TabIndex        =   28
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "CMP"
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
            Left            =   11513
            TabIndex        =   27
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   315
            TabIndex        =   26
            Top             =   180
            Width           =   945
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
            Left            =   8895
            TabIndex        =   25
            Top             =   1740
            Width           =   90
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   6495
            TabIndex        =   24
            Top             =   1740
            Width           =   2325
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1830
            TabIndex        =   23
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Following Guards are not linked to any godowns and hence attendance can not be marked."
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
            Left            =   120
            TabIndex        =   22
            Top             =   5760
            Width           =   9360
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   10680
            TabIndex        =   21
            Top             =   5760
            Width           =   2325
         End
         Begin VB.Label Label24 
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
            Left            =   13080
            TabIndex        =   20
            Top             =   5760
            Width           =   90
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
            Left            =   3855
            TabIndex        =   19
            Top             =   180
            Width           =   1065
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
            Left            =   6840
            TabIndex        =   18
            Top             =   173
            Width           =   855
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10680
      Width           =   10935
   End
End
Attribute VB_Name = "FrmRpt_TruckwiseCxtfDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mCommodityID, mAgentID, mExchangeTypeID  As Variant

Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Agent.RecordCount
   CmbAgent.AddItem RsMst_Agent!AgentName
   RsMst_Agent.MoveNext
Next
CmbAgent.Text = tmp
End Sub
'----
Private Sub CmbAgent_LostFocus()
If CmbAgent.Text = "" Then Exit Sub

RsMst_Agent.MoveFirst
RsMst_Agent.Find "agentName = '" & CmbAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection!!!"
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID

End Sub


Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No record found!!!"
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
If CmbExchangeTypeID.Text = "" Then Exit Sub

RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub

Private Sub CmbCMP_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMP.Text
wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If
Call TableMst_CMP(" Where CloseDate is Null And LocationID = " & mLocationID & " " & wc)
CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
   CmbCMP.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbCMP.Text = tmp

End Sub

Private Sub CmbCMP_LostFocus()
If CmbCMP.Text = "" Then Exit Sub

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCMP.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
   CmbCommodity.AddItem RsMst_Commodity!Commodity
   RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then Exit Sub

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State(" where " & Gen_WcState)
Else
   Call TableMst_State
End If

CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found!!!"
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
   CmbLocation.Text = ""
   CmbCMP.Text = ""
   Exit Sub
End If
If tmp <> CmbStateID.Text Then
   CmbLocation.Text = ""
   CmbCMP.Text = ""
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection!!!"
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
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where StateID= " & mStateID & " AND " & Gen_WcLocation)
Else
   Call TableMst_Location(" Where StateID= " & mStateID)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found!!!"
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
   CmbCMP.Text = ""
   Exit Sub
End If

If tmp <> CmbLocation.Text Then
   CmbCMP.Text = ""
End If
' Code for clearing Godown combox in case of change

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Selection!!! "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbTransactiontype_Change()
If Trim(CmbTransactiontype.Text) = "" Then Exit Sub
   
If LCase(CmbTransactiontype.Text) <> "deposit" And LCase(CmbTransactiontype.Text) <> "withdraw" Then
   MsgBox "Invalid selection!!!"
   CmbTransactiontype.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbStateID.Text = ""
CmbLocation.Text = ""
CmbCMP.Text = ""
CmbCommodity.Text = ""
CmbAgent.Text = ""
CmbExchangeTypeID.Text = ""
CmbTransactiontype.Text = ""
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_TruckwiseCxtfDetail.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_TruckwiseCxtfDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
If TxtFromDate.Value > Date Then
   MsgBox "Future Date not Allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
If TxtToDate.Value > Date Then
   MsgBox "Future Date not Allowed!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From Date should not greater than To Date!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
Call Grid_Head
MsgBox "Wait till Next Message!!!"

If CmbLocation.Text <> "" Then
   wc = " AND ML.LocationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = " AND ML. " & Gen_WcLocation
End If
If CmbStateID.Text <> "" Then
   wc = wc & " AND MS.StateID = " & mStateID
ElseIf Gen_WcState <> "" Then
   wc = wc & " AND MS. " & Gen_WcState
End If
If CmbCMP.Text <> "" Then
   wc = wc & " AND MCMP.CMPID = " & mCMPID
ElseIf Gen_WcCMP <> "" Then
   wc = wc & " AND MCMP. " & Gen_WcCMP
End If
If CmbCommodity.Text <> "" Then
   wc = wc & " AND MC.CommodityID = " & mCommodityID
End If
If CmbAgent.Text <> "" Then
   wc = wc & " AND MA.AgentID = " & mAgentID
End If
If CmbExchangeTypeID.Text <> "" Then
   wc = wc & " AND ME.ExchangeTypeID = " & mExchangeTypeID
End If
If Trim(CmbTransactiontype.Text) <> "" Then
   wc = wc & " AND TCD.TxnType = '" & Left$(CmbTransactiontype.Text, 1) & "'"
End If


tmp = " SELECT MS.StateName,ML.LocationName, MC.Commodity,ME.ExchangeType,MCMP.CMPName,MA.AgentName,"
tmp = tmp & " TCD.CxTF_DetailsID,TCD.TxnType,TCD.CxTFDate,"
tmp = tmp & " TCT.TruckNo , WeighBridgeSlipno, TCT.TotalWeight, TCT.TareWeight, TCT.GrossWeight, TCT.Remarks"
tmp = tmp & " FROM Txn_CxTF_TruckDetail TCT"
tmp = tmp & " LEFT JOIN Txn_CxTF_Details TCD ON TCT.CxTF_DetailsID = TCD.CxTF_DetailsID AND TCT.TxnType = TCD.TxnType AND TCT.SM_Prefix = TCD.SM_Prefix"
tmp = tmp & " LEFT JOIN Mst_Commodity MC ON MC.CommodityID=TCD.CommodityID"
tmp = tmp & " LEFT JOIN Mst_ExchangeType ME ON ME.ExchangeTypeID=TCD.ExchangeTypeID"
tmp = tmp & " LEFT JOIN Mst_CMP MCMP ON MCMP.CMPID=TCD.CMPID"
tmp = tmp & " LEFT JOIN Mst_Location ML ON ML.SM_PreFix=TCD.SM_PreFix"
tmp = tmp & " LEFT JOIN Mst_State MS ON Ms.StateID=ML.StateID"
tmp = tmp & " LEFT JOIN Mst_Agent MA ON TCD.AgentID=MA.AgentID"
tmp = tmp & " WHERE TCD.CxTFDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "' "
tmp = tmp & wc
tmp = tmp & " ORDER BY MS.StateName, ML.Locationname, MCMP.CMPName, ME.ExchangeType"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Label22.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      ProgressBar1.Max = .Rows
      For I = 2 To TmpRs.RecordCount + 1
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!AgentName), "", TmpRs!AgentName)
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID)
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!TxnType), "", TmpRs!TxnType)
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs!CxTFDate), "", Format(TmpRs!CxTFDate, "dd-MMM-yyyy"))
         .TextMatrix(I, 9) = IIf(IsNull(TmpRs!TruckNo), "", TmpRs!TruckNo)
         .TextMatrix(I, 10) = IIf(IsNull(TmpRs!WeighBridgeSlipno), "", TmpRs!WeighBridgeSlipno)
         .TextMatrix(I, 11) = IIf(IsNull(TmpRs!TotalWeight), "", Format(TmpRs!TotalWeight, "#0.000"))
         .TextMatrix(I, 12) = IIf(IsNull(TmpRs!TareWeight), "", Format(TmpRs!TareWeight, "#0.000"))
         .TextMatrix(I, 13) = IIf(IsNull(TmpRs!GrossWeight), "", Format(TmpRs!GrossWeight, "#0.000"))
         .TextMatrix(I, 14) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
         
         .TextMatrix(0, 11) = Format(Val(.TextMatrix(I, 11)) + Val(.TextMatrix(0, 11)), "#0.000")
         .TextMatrix(0, 12) = Format(Val(.TextMatrix(I, 12)) + Val(.TextMatrix(0, 12)), "#0.000")
         .TextMatrix(0, 13) = Format(Val(.TextMatrix(I, 13)) + Val(.TextMatrix(0, 13)), "#0.000")
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
   MsgBox "No Record Found!!!"
End If
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
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
HScroll1.Max = Me.ScaleWidth - Frame1.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame1.Width >= Me.Width Then HScroll1.Enabled = True
If Frame1.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

End Sub

Private Sub VScroll1_Scroll()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

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

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call CmdClear_Click
End Sub

Private Sub Grid_Head()
Label22.Caption = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 15
    .FixedCols = 1
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    
    .TextMatrix(1, 0) = "State"
    .TextMatrix(1, 1) = "Location"
    .TextMatrix(1, 2) = "Commodity"
    .TextMatrix(1, 3) = "Exchange Type"
    .TextMatrix(1, 4) = "CMP"
    .TextMatrix(1, 5) = "Agent Name"
    .TextMatrix(1, 6) = "CxTF No"
    .TextMatrix(1, 7) = "D/W"
    .TextMatrix(1, 8) = "CxTF Date"
    .TextMatrix(1, 9) = "Truck No"
    .TextMatrix(1, 10) = "Weigh Bridge Slip No"
    .TextMatrix(1, 11) = "Total Weight"
    .TextMatrix(1, 12) = "Tare Weight"
    .TextMatrix(1, 13) = "Gross Weight"
    .TextMatrix(1, 14) = "Remark"
         
    .ColWidth(0) = 1800
    .ColWidth(1) = 1800
    .ColWidth(2) = 1600
    .ColWidth(3) = 1200
    .ColWidth(4) = 1900
    .ColWidth(5) = 2200
    .ColWidth(6) = 1000
    .ColWidth(7) = 800
    .ColWidth(8) = 1100
    .ColWidth(9) = 1400
    
    .RowHeight(1) = 600
End With
End Sub

