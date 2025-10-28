VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_NNNRBOEMonitoringMIS 
   Caption         =   "NNNR BOE Monitoring MIS"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   15015
      TabIndex        =   19
      Top             =   315
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   9840
      Width           =   10935
   End
   Begin VB.PictureBox Frame4 
      Height          =   9660
      Left            =   120
      ScaleHeight     =   9600
      ScaleWidth      =   14895
      TabIndex        =   0
      Top             =   0
      Width           =   14955
      Begin VB.Frame Frame11 
         Height          =   1380
         Left            =   45
         TabIndex        =   1
         Top             =   -30
         Width           =   14760
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
            Left            =   5550
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   390
            Width           =   3750
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
            Left            =   13440
            TabIndex        =   5
            Top             =   390
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
            Left            =   11010
            TabIndex        =   4
            Top             =   390
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
            Left            =   12225
            TabIndex        =   3
            Top             =   390
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
            Left            =   1680
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   390
            Width           =   3750
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   330
            Left            =   120
            TabIndex        =   7
            Top             =   390
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            Format          =   109379587
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin VB.Label Label1 
            Caption         =   "Month Year"
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
            TabIndex        =   12
            Top             =   120
            Width           =   1095
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
            Left            =   7005
            TabIndex        =   11
            Top             =   120
            Width           =   855
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
            Left            =   13740
            TabIndex        =   10
            Top             =   1020
            Width           =   90
         End
         Begin VB.Label Label5 
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
            Left            =   11415
            TabIndex        =   9
            Top             =   1020
            Width           =   2280
         End
         Begin VB.Label Label4 
            Caption         =   "State"
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
            Left            =   3315
            TabIndex        =   8
            Top             =   120
            Width           =   480
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   8055
         Left            =   45
         TabIndex        =   13
         Top             =   1440
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   14208
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "NNNR BOE Monitoring MIS"
         TabPicture(0)   =   "FrmRpt_NNNRBOEMonitoringMIS.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "BOEs, Generated/Pending by CMPM"
         TabPicture(1)   =   "FrmRpt_NNNRBOEMonitoringMIS.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "BOEs, Generated/Pending by FAG"
         TabPicture(2)   =   "FrmRpt_NNNRBOEMonitoringMIS.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSHFlexGrid5"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "BOEs, TI Uploaded/Pending "
         TabPicture(3)   =   "FrmRpt_NNNRBOEMonitoringMIS.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "MSHFlexGrid7"
         Tab(3).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7545
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   13309
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   7545
            Left            =   -74880
            TabIndex        =   15
            Top             =   420
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   13309
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   7545
            Left            =   -74880
            TabIndex        =   16
            Top             =   420
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   13309
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
            Height          =   7425
            Left            =   -74880
            TabIndex        =   17
            Top             =   420
            Width           =   14520
            _ExtentX        =   25612
            _ExtentY        =   13097
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
End
Attribute VB_Name = "FrmRpt_NNNRBOEMonitoringMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID, mTotalDays As Variant
'Dim mTotBOEs, mGenBOECMPM, mPenBOECMPM, mProBOEFAG, mPenBOEFAG, mBOETISM, mBOETI As Variant
Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
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
VScroll1.Max = Me.ScaleHeight - Frame4.Height '4
VScroll1.LargeChange = Frame4.ScaleHeight
VScroll1.SmallChange = Frame4.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width '4
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True ''4
If Frame4.Height < Me.Height Then VScroll1.Enabled = False '4
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True '4
If Frame4.Width < Me.Width Then HScroll1.Enabled = False '4
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbLocation.Clear
   CmbStateID.SetFocus
   Exit Sub
End If

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID = " & mStateID & " ")
Else
   Call TableMst_Location(" Where StateID = " & mStateID & " And " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
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
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text

If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
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
   CmbLocation.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date
CmbStateID = ""
CmbLocation = ""
CmdExcel.Enabled = False
Label22.Caption = "-"
Call Grid_Head
Call Grid_HeadTotalBOE
Call Grid_HeadProcessedFAG
Call Grid_HeadTI
SSTab1.Tab = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NNNRBOEMonitoringMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NNNRBOEMonitoringMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "NNNR BOE Monitoring MIS"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
    Next
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "BOEs, Generated-Pending by CMPM"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
    Next
Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "BOEs, Generated-Pending by FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid5.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid5.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid5.TextMatrix(I, C)
    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "BOEs, TI Uploaded-Pending "
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid7.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid7.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid7.TextMatrix(I, C)
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
Call Grid_HeadTotalBOE
Call Grid_HeadProcessedFAG
Call Grid_HeadTI

SSTab1.Tab = 0

Dim wc As String
Dim tmp1 As String
wc = ""


TxtFromDate.Value = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")

mTotalDays = Day(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1, 0))
mToDate = Format(Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-" & mTotalDays, "yyyy-mm-dd")


If CmbLocation.Text <> "" Then
   wc = wc & " And ML.LocationID = " & mLocationID
   
Else
   wc = wc & Gen_Location
End If

tmp = " Select Distinct MS.StateID , MS.StateName, ML.LocationID, ML.LocationName, ME.EmployeeName, ME.EmployeeNo"
tmp = tmp & " From Mst_Location ML"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_CMP MCMP On ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_Employee ME on ME.EmployeeID = ML.AM_EmployeeID"
tmp = tmp & " Where (MCMP.CloseDate is Null or MCMP.CloseDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "') And MCMP.StartDate <= '" & Format(mToDate, "yyyy-MM-dd") & "' "
'tmp = "Select MS.StateID , MS.StateName, ML.LocationID, ML.LocationName, ME.EmployeeName, ME.EmployeeNo"
'tmp = tmp & " From Mst_Location ML"
'tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
'tmp = tmp & " Inner Join Mst_Employee ME on ME.EmployeeID = ML.AM_EmployeeID"
If CmbLocation.Text <> "" Then
    tmp = tmp & " And ML.LocationID = " & mLocationID
Else
    If CmbStateID.Text <> "" Then
        tmp = tmp & " And MS.StateID = " & mStateID
    End If
End If

If CmbLocation = "" And CmbStateID = "" Then
   If Gen_Location <> "" Then
       tmp = tmp & " And " & Gen_Location
   End If
End If
tmp = tmp & " Order by  MS.StateName,ML.LocationName"

Call TmpTable

''''''''--For Minimum BOD

'tmp = " Select ML.LocationID, ML.LocationName, Min(NewBODDate) MinDate"
'tmp = tmp & " from Log_BODUpdate LBU"
'tmp = tmp & " Inner Join Mst_CMP MCMP On LBU.CMPID = MCMP.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
'tmp = tmp & " group by ML.LocationID, ML.LocationName"


'Commented On 13 Dec 2012 411 PM to show max BOD Date of Open CMP only
'tmp = " Select ML.LocationID, ML.LocationName, min(MCMP.CurrentDate)MinDate"
'tmp = tmp & " from Mst_CMP MCMP"
'tmp = tmp & " left join Log_BODUpdate LBU on LBU.CMPID = MCMP.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
'tmp = tmp & " group by ML.LocationID, ML.LocationName"


tmp = " Select ML.LocationID, ML.LocationName, Min(MCMP.CurrentDate)MinDate,Max(MCMP.CurrentDate)MaxDate"
tmp = tmp & " from Mst_CMP MCMP"
tmp = tmp & " left join Log_BODUpdate LBU on LBU.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Where MCMP.CloseDate is null "
tmp = tmp & " group by ML.LocationID, ML.LocationName"



Call Tmp1Table

'''''''''''''''--For Invoice Process initiated Date
tmp = " Select TINV.YearMonth,ML.LocationID, ML.LocationName, min(TINV.ConCurrency) StartDate"
tmp = tmp & " From Txn_InvMonGSLDetail TINV"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Where YearMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TINV.ExchangeTypeID=1 And (TINV.ContractID = 0 Or TINV.ContractID is Null)"
tmp = tmp & " group by ML.LocationID, ML.LocationName,TINV.YearMonth"

Call Tmp2Table

'''''''''''''''''--For Total BOEs
tmp = " Select YearMonth,ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TInv.YearMonth , Mst.NavCode, TInv.PINo, Sum(InvoiceAmount) as Amount"
tmp = tmp & " From Txn_InvMonGSLDetail TINV"
tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MClient On TINV.ClientID = MClient.ClientID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"
tmp = tmp & " Where YearMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And Me.ExchangeTypeID=1 And (TINV.ContractID = 0 Or TINV.ContractID is Null)"
tmp = tmp & " group by ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TInv.YearMonth , Mst.NavCode, TInv.PINo, YearMonth"

Call Tmp3Table


''''''''''''''''''''''---For Generated BOEs
tmp = " Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TPI.InvoiceDate , TPI.InvoiceNo, TPI.ServiceTaxNAVCode, TPI.Amount, TPI.Discount, TPI.Premium"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientID = MClient.ClientID"
tmp = tmp & " Where BillingMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And Me.ExchangeTypeID=1 And TPI.Flag = 'C'"

Call Tmp4Table

'''''''''''''''''''''''''''''''--For Pending BOE
tmp = " Select YearMonth,ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TInv.YearMonth , Mst.NavCode, TInv.PINo, Sum(InvoiceAmount) as Amount, "
''added by rakesh on 15/03/2013 for discount and premium column
tmp = tmp & " (select sum(TPI.discount) From Txn_ProInvoice TPI where TPI.invoiceno = TINv.PINo) Discount,"
tmp = tmp & " (select sum(TPI.Premium) From Txn_ProInvoice TPI where TPI.invoiceno = TINv.PINo) Premium"
tmp = tmp & " From Txn_InvMonGSLDetail TINV"
tmp = tmp & " Inner Join Mst_Commodity MC On TINV.CommodityID = MC.CommodityID"
tmp = tmp & " Inner Join Mst_CMP MCMP On TINV.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Client MClient On TINV.ClientID = MClient.ClientID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TINV.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Left Join Mst_ServiceTax MST On MC.ServiceTaxID = MST.ServiceTaxID"
tmp = tmp & " Where YearMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And Me.ExchangeTypeID=1 And (TINV.ContractID = 0 Or TINV.ContractID is Null) And TINV.PINo is null"
tmp = tmp & " group by ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TInv.YearMonth , Mst.NavCode, TInv.PINo, YearMonth"

Call Tmp5Table

''''''''''''''''''''--For BOEs Processed By FAG
tmp = " Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TPI.InvoiceDate , TPI.InvoiceNo, TPI.ServiceTaxNAVCode, TPI.Amount, TPI.Discount, TPI.Premium,   BatchID"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientID = MClient.ClientID"
'tmp = tmp & " Where Me.ExchangeTypeID = 1 And BatchID > 0 and BillingMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Where Me.ExchangeTypeID = 1 And BillingMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.Flag = 'C'"
Call Tmp6Table

'''''''''''''''''''''--For BOEs Pending from FAG

'tmp = " Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
'tmp = tmp & " TPI.InvoiceDate , TPI.InvoiceNo, TPI.ServiceTaxNAVCode, Amount, BatchID"
'tmp = tmp & " From Txn_ProInvoice TPI"
'tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
'tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientIDRow = MClient.ClientIDRow"
'tmp = tmp & " Where Me.ExchangeTypeID = 1 And BatchID = 0 and BillingMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "'"

'Call Tmp7Table

'''''''''''''''''''''''--For TI Uploaded
tmp = " Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType, "
tmp = tmp & " TPI.InvoiceDate , TPI.InvoiceNo, TPI.ServiceTaxNAVCode, TPI.Amount, TPI.Discount, TPI.Premium, BatchID,"
tmp = tmp & " TDI.LocationID , TDI.TINO, TDI.TIAmount, TDI.TIDate, TDI.TIID"
tmp = tmp & " From Txn_TIDetails TDI"
tmp = tmp & " Inner Join Txn_ProInvoice TPI On TDI.PIID = TPI.InvoiceNo"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientID = MClient.ClientID"
tmp = tmp & " Where TDI.PIFLag='NNNR' And Me.ExchangeTypeID = 1 And BillingMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.Flag = 'C'"

Call Tmp7Table


''''''''''''''''''''''''''--For TI Pending
tmp = " Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TPI.InvoiceDate , TPI.InvoiceNo, TPI.ServiceTaxNAVCode, TPI.Amount, TPI.Discount, TPI.Premium, BatchID"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientID = MClient.ClientID"
tmp = tmp & " Where Me.ExchangeTypeID = 1 And BillingMonth='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.Flag = 'C'"
tmp = tmp & " And TPI.InvoiceNo not in (Select PIID From Txn_TIDetails TDI Where TDI.PIFLag='NNNR') And BatchID > 0 "

Call Tmp0Table

If TmpRs.RecordCount > 0 Then
   Label22.Caption = TmpRs.RecordCount
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid1.TextMatrix(I, 0) = Format(TxtFromDate.Value, "MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!StateName
       MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!Locationname
       
       TmpRs1.Filter = " LocationID = " & TmpRs!LocationID
       If TmpRs1.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 3) = Format(TmpRs1!MinDate, "dd-MMM-yyyy")
           MSHFlexGrid1.TextMatrix(I, 13) = Format(TmpRs1!MaxDate, "dd-MMM-yyyy")
       End If
       TmpRs1.Filter = ""

       TmpRs2.Filter = " LocationID = " & TmpRs!LocationID
       If TmpRs2.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs2!StartDate, "dd-MMM-yyyy")
       End If
       TmpRs2.Filter = ""
       
       TmpRs3.Filter = " LocationID = " & TmpRs!LocationID
       If TmpRs3.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 5) = TmpRs3.RecordCount
           MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + TmpRs3.RecordCount
           'k = MSHFlexGrid2.Rows - 1
           'MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs3.RecordCount
           'For j = 1 To TmpRs3.RecordCount
           '    MSHFlexGrid2.TextMatrix(k, 0) = TmpRs3!LocationName
           '    MSHFlexGrid2.TextMatrix(k, 1) = TmpRs3!ClientName
           '    MSHFlexGrid2.TextMatrix(k, 2) = TmpRs3!ExchangeType
           '    MSHFlexGrid2.TextMatrix(k, 3) = IIf(IsNull(TmpRs3!PINo), "", TmpRs3!PINo)
           '    MSHFlexGrid2.TextMatrix(k, 5) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
           '    MSHFlexGrid2.TextMatrix(k, 6) = IIf(IsNull(TmpRs3!Amount), "", TmpRs3!Amount)
           '    k = k + 1
           '    TmpRs3.MoveNext
           'Next j
       End If
       TmpRs3.Filter = ""
       
       TmpRs4.Filter = " LocationID = " & TmpRs!LocationID & ""
       If TmpRs4.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 6) = TmpRs4.RecordCount
           MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + TmpRs4.RecordCount
           k = MSHFlexGrid2.Rows - 1
           MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs4.RecordCount
           For j = 1 To TmpRs4.RecordCount
               MSHFlexGrid2.TextMatrix(k, 0) = TmpRs4!Locationname
               MSHFlexGrid2.TextMatrix(k, 1) = TmpRs4!ClientName
               MSHFlexGrid2.TextMatrix(k, 2) = TmpRs4!ExchangeType
               MSHFlexGrid2.TextMatrix(k, 3) = IIf(IsNull(TmpRs4!InvoiceNo), "", TmpRs4!InvoiceNo)
               MSHFlexGrid2.TextMatrix(k, 4) = Format(TmpRs4!InvoiceDate, "dd-MMM-yyyy")
               MSHFlexGrid2.TextMatrix(k, 5) = IIf(IsNull(TmpRs4!ServiceTaxNAVCode), "", TmpRs4!ServiceTaxNAVCode)
               MSHFlexGrid2.TextMatrix(k, 6) = IIf(IsNull(TmpRs4!Amount), "", TmpRs4!Amount)
               MSHFlexGrid2.TextMatrix(k, 7) = IIf(IsNull(TmpRs4!Discount), "", TmpRs4!Discount)
               MSHFlexGrid2.TextMatrix(k, 8) = IIf(IsNull(TmpRs4!Premium), "", TmpRs4!Premium)
               MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(k, 6))
               MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(k, 7))
               MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(k, 8))
               MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + 1
               k = k + 1
               TmpRs4.MoveNext
           Next j
       End If
       TmpRs4.Filter = ""

       TmpRs5.Filter = " LocationID = " & TmpRs!LocationID
       If TmpRs5.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 7) = TmpRs5.RecordCount
           MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + TmpRs5.RecordCount
           k = MSHFlexGrid2.Rows - 1
           MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs5.RecordCount
           For j = 1 To TmpRs5.RecordCount
               MSHFlexGrid2.TextMatrix(k, 0) = TmpRs5!Locationname
               MSHFlexGrid2.TextMatrix(k, 1) = TmpRs5!ClientName
               MSHFlexGrid2.TextMatrix(k, 2) = TmpRs5!ExchangeType
               MSHFlexGrid2.TextMatrix(k, 5) = IIf(IsNull(TmpRs5!NavCode), "", TmpRs5!NavCode)
               MSHFlexGrid2.TextMatrix(k, 6) = IIf(IsNull(TmpRs5!Amount), "", TmpRs5!Amount)
               MSHFlexGrid2.TextMatrix(k, 7) = IIf(IsNull(TmpRs5!Discount), "", TmpRs5!Discount)
               MSHFlexGrid2.TextMatrix(k, 8) = IIf(IsNull(TmpRs5!Premium), "", TmpRs5!Premium)
               MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(k, 6))
               MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(k, 7))
               MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(k, 8))
               MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + 1
               k = k + 1
               TmpRs5.MoveNext
           Next j
       End If
       TmpRs5.Filter = ""

       TmpRs6.Filter = " LocationID = " & TmpRs!LocationID & "" ' And BatchID > 0"
       If TmpRs6.RecordCount > 0 Then
           'MSHFlexGrid1.TextMatrix(I, 8) = TmpRs6.RecordCount
          
           k = MSHFlexGrid5.Rows - 1
           MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + TmpRs6.RecordCount
           For j = 1 To TmpRs6.RecordCount
               If TmpRs6!BatchID > 0 Then
                  MSHFlexGrid1.TextMatrix(I, 8) = Val(MSHFlexGrid1.TextMatrix(I, 8)) + 1
                  MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + 1
               Else
                  MSHFlexGrid1.TextMatrix(I, 9) = Val(MSHFlexGrid1.TextMatrix(I, 9)) + 1
                  MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + 1
               End If
               MSHFlexGrid5.TextMatrix(k, 0) = TmpRs6!Locationname
               MSHFlexGrid5.TextMatrix(k, 1) = TmpRs6!ClientName
               MSHFlexGrid5.TextMatrix(k, 2) = TmpRs6!ExchangeType
               MSHFlexGrid5.TextMatrix(k, 3) = IIf(IsNull(TmpRs6!InvoiceNo), "", TmpRs6!InvoiceNo)
               MSHFlexGrid5.TextMatrix(k, 4) = Format(TmpRs6!InvoiceDate, "dd-MMM-yyyy")
               MSHFlexGrid5.TextMatrix(k, 5) = IIf(IsNull(TmpRs6!ServiceTaxNAVCode), "", TmpRs6!ServiceTaxNAVCode)
               MSHFlexGrid5.TextMatrix(k, 6) = IIf(IsNull(TmpRs6!Amount), "", TmpRs6!Amount)
               MSHFlexGrid5.TextMatrix(k, 7) = IIf(IsNull(TmpRs6!Discount), "", TmpRs6!Discount)
               MSHFlexGrid5.TextMatrix(k, 8) = IIf(IsNull(TmpRs6!Premium), "", TmpRs6!Premium)
               MSHFlexGrid5.TextMatrix(k, 9) = IIf(IsNull(TmpRs6!BatchID), "", TmpRs6!BatchID)
               MSHFlexGrid5.TextMatrix(0, 6) = Val(MSHFlexGrid5.TextMatrix(0, 6)) + Val(MSHFlexGrid5.TextMatrix(k, 6))
               MSHFlexGrid5.TextMatrix(0, 7) = Val(MSHFlexGrid5.TextMatrix(0, 7)) + Val(MSHFlexGrid5.TextMatrix(k, 7))
               MSHFlexGrid5.TextMatrix(0, 8) = Val(MSHFlexGrid5.TextMatrix(0, 8)) + Val(MSHFlexGrid5.TextMatrix(k, 8))
               MSHFlexGrid5.TextMatrix(0, 3) = Val(MSHFlexGrid5.TextMatrix(0, 3)) + 1
               
               k = k + 1
               
               TmpRs6.MoveNext
           Next j
           
       End If
       TmpRs6.Filter = ""

''       TmpRs6.Filter = " LocationID = " & TmpRs!LocationID & " And BatchID = 0"
''       If TmpRs6.RecordCount > 0 Then
''           MSHFlexGrid1.TextMatrix(I, 9) = TmpRs6.RecordCount
''           MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + TmpRs6.RecordCount
''           k = MSHFlexGrid5.Rows - 1
''           MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + TmpRs6.RecordCount
''           For j = 1 To TmpRs6.RecordCount
''               MSHFlexGrid5.TextMatrix(k, 0) = TmpRs6!LocationName
''               MSHFlexGrid5.TextMatrix(k, 1) = TmpRs6!ClientName
''               MSHFlexGrid5.TextMatrix(k, 2) = TmpRs6!ExchangeType
''               MSHFlexGrid5.TextMatrix(k, 3) = IIf(IsNull(TmpRs6!InvoiceNo), "", TmpRs6!InvoiceNo)
''               MSHFlexGrid5.TextMatrix(k, 4) = Format(TmpRs6!InvoiceDate, "dd-MMM-yyyy")
''               MSHFlexGrid5.TextMatrix(k, 5) = IIf(IsNull(TmpRs6!ServiceTaxNAVCode), "", TmpRs6!ServiceTaxNAVCode)
''               MSHFlexGrid5.TextMatrix(k, 6) = IIf(IsNull(TmpRs6!Amount), "", TmpRs6!Amount)
''               MSHFlexGrid5.TextMatrix(k, 7) = IIf(IsNull(TmpRs6!BatchID), "", TmpRs6!BatchID)
''               k = k + 1
''               TmpRs6.MoveNext
''           Next j
''       End If
''       TmpRs6.Filter = ""
       
       TmpRs7.Filter = " LocationID = " & TmpRs!LocationID
       If TmpRs7.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 10) = TmpRs7.RecordCount
           MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + TmpRs7.RecordCount
           k = MSHFlexGrid7.Rows - 1
           MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + TmpRs7.RecordCount
           For j = 1 To TmpRs7.RecordCount
               MSHFlexGrid7.TextMatrix(k, 0) = TmpRs7!Locationname
               MSHFlexGrid7.TextMatrix(k, 1) = TmpRs7!ClientName
               MSHFlexGrid7.TextMatrix(k, 2) = TmpRs7!ExchangeType
               MSHFlexGrid7.TextMatrix(k, 3) = IIf(IsNull(TmpRs7!InvoiceNo), "", TmpRs7!InvoiceNo)
               MSHFlexGrid7.TextMatrix(k, 4) = Format(TmpRs7!InvoiceDate, "dd-MMM-yyyy")
               MSHFlexGrid7.TextMatrix(k, 5) = IIf(IsNull(TmpRs7!ServiceTaxNAVCode), "", TmpRs7!ServiceTaxNAVCode)
               MSHFlexGrid7.TextMatrix(k, 6) = IIf(IsNull(TmpRs7!Amount), "", TmpRs7!Amount)
               MSHFlexGrid7.TextMatrix(k, 7) = IIf(IsNull(TmpRs7!Discount), "", TmpRs7!Discount)
               MSHFlexGrid7.TextMatrix(k, 8) = IIf(IsNull(TmpRs7!Premium), "", TmpRs7!Premium)
               MSHFlexGrid7.TextMatrix(k, 9) = IIf(IsNull(TmpRs7!BatchID), "", TmpRs7!BatchID)
               MSHFlexGrid7.TextMatrix(k, 10) = IIf(IsNull(TmpRs7!TINo), "", TmpRs7!TINo)
               MSHFlexGrid7.TextMatrix(k, 11) = Format(TmpRs7!TIDate, "dd-MMM-yyyy")
               MSHFlexGrid7.TextMatrix(k, 12) = IIf(IsNull(TmpRs7!TIAmount), "", TmpRs7!TIAmount)
               
               MSHFlexGrid7.TextMatrix(0, 6) = Val(MSHFlexGrid7.TextMatrix(0, 6)) + Val(MSHFlexGrid7.TextMatrix(k, 6))
               MSHFlexGrid7.TextMatrix(0, 7) = Val(MSHFlexGrid7.TextMatrix(0, 7)) + Val(MSHFlexGrid7.TextMatrix(k, 7))
               MSHFlexGrid7.TextMatrix(0, 8) = Val(MSHFlexGrid7.TextMatrix(0, 8)) + Val(MSHFlexGrid7.TextMatrix(k, 8))
               MSHFlexGrid7.TextMatrix(0, 12) = Val(MSHFlexGrid7.TextMatrix(0, 12)) + Val(MSHFlexGrid7.TextMatrix(k, 12))
               MSHFlexGrid7.TextMatrix(0, 3) = Val(MSHFlexGrid7.TextMatrix(0, 3)) + 1
               k = k + 1
               TmpRs7.MoveNext
           Next j
       End If
       TmpRs7.Filter = ""

       TmpRs0.Filter = " LocationID = " & TmpRs!LocationID
       If TmpRs0.RecordCount > 0 Then
           MSHFlexGrid1.TextMatrix(I, 11) = TmpRs0.RecordCount
           MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + TmpRs0.RecordCount
           k = MSHFlexGrid7.Rows - 1
           MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + TmpRs0.RecordCount
           For j = 1 To TmpRs0.RecordCount
               MSHFlexGrid7.TextMatrix(k, 0) = TmpRs0!Locationname
               MSHFlexGrid7.TextMatrix(k, 1) = TmpRs0!ClientName
               MSHFlexGrid7.TextMatrix(k, 2) = TmpRs0!ExchangeType
               MSHFlexGrid7.TextMatrix(k, 3) = IIf(IsNull(TmpRs0!InvoiceNo), "", TmpRs0!InvoiceNo)
               MSHFlexGrid7.TextMatrix(k, 4) = Format(TmpRs0!InvoiceDate, "dd-MMM-yyyy")
               MSHFlexGrid7.TextMatrix(k, 5) = IIf(IsNull(TmpRs0!ServiceTaxNAVCode), "", TmpRs0!ServiceTaxNAVCode)
               MSHFlexGrid7.TextMatrix(k, 6) = IIf(IsNull(TmpRs0!Amount), "", TmpRs0!Amount)
               MSHFlexGrid7.TextMatrix(k, 7) = IIf(IsNull(TmpRs0!Discount), "", TmpRs0!Discount)
               MSHFlexGrid7.TextMatrix(k, 8) = IIf(IsNull(TmpRs0!Premium), "", TmpRs0!Premium)
               MSHFlexGrid7.TextMatrix(k, 9) = IIf(IsNull(TmpRs0!BatchID), "", TmpRs0!BatchID)
               MSHFlexGrid7.TextMatrix(0, 6) = Val(MSHFlexGrid7.TextMatrix(0, 6)) + Val(MSHFlexGrid7.TextMatrix(k, 6))
               MSHFlexGrid7.TextMatrix(0, 7) = Val(MSHFlexGrid7.TextMatrix(0, 7)) + Val(MSHFlexGrid7.TextMatrix(k, 7))
               MSHFlexGrid7.TextMatrix(0, 8) = Val(MSHFlexGrid7.TextMatrix(0, 8)) + Val(MSHFlexGrid7.TextMatrix(k, 8))
               MSHFlexGrid7.TextMatrix(0, 3) = Val(MSHFlexGrid7.TextMatrix(0, 3)) + 1
               
               k = k + 1
               TmpRs0.MoveNext
           Next j
       End If
       TmpRs0.Filter = ""
       
       MSHFlexGrid1.TextMatrix(I, 12) = TmpRs!EmployeeName & "~" & TmpRs!EmployeeNo
       
       
       TmpRs.MoveNext
   Next
End If
       

CmdExcel.Enabled = True
MsgBox "Done!!!"

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
'   CmdGetReport.Enabled = False
'   CmdExcel.Enabled = False
   Frame11.Enabled = False
   Exit Sub
End If

Frame11.Enabled = True
Call CmdClear_Click

End Sub

Private Sub Grid_HeadTI()
MSHFlexGrid7.Clear
MSHFlexGrid7.Rows = 3
MSHFlexGrid7.FixedRows = 2
MSHFlexGrid7.Cols = 13

MSHFlexGrid7.TextMatrix(1, 0) = "Location"
MSHFlexGrid7.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid7.TextMatrix(1, 2) = "ExchangeType"
MSHFlexGrid7.TextMatrix(1, 3) = "Invoice No"
MSHFlexGrid7.TextMatrix(1, 4) = "Invoice Date"
MSHFlexGrid7.TextMatrix(1, 5) = "Service Tax"
MSHFlexGrid7.TextMatrix(1, 6) = "Amount"
MSHFlexGrid7.TextMatrix(1, 7) = "Discount"
MSHFlexGrid7.TextMatrix(1, 8) = "Premium"
MSHFlexGrid7.TextMatrix(1, 9) = "Batch ID"
MSHFlexGrid7.TextMatrix(1, 10) = "TI No"
MSHFlexGrid7.TextMatrix(1, 11) = "TI Date"
MSHFlexGrid7.TextMatrix(1, 12) = "TI Amount"

MSHFlexGrid7.ColWidth(0) = 1800
MSHFlexGrid7.ColWidth(1) = 4000
MSHFlexGrid7.ColWidth(2) = 1700
MSHFlexGrid7.ColWidth(3) = 1500
MSHFlexGrid7.ColWidth(4) = 1500
MSHFlexGrid7.ColWidth(5) = 1500
MSHFlexGrid7.ColWidth(6) = 1500
MSHFlexGrid7.ColWidth(7) = 1500
MSHFlexGrid7.ColWidth(8) = 2000
MSHFlexGrid7.ColWidth(9) = 1500

End Sub


Private Sub Grid_HeadProcessedFAG()
MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 3
MSHFlexGrid5.FixedRows = 2
MSHFlexGrid5.Cols = 10

MSHFlexGrid5.TextMatrix(1, 0) = "Location"
MSHFlexGrid5.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid5.TextMatrix(1, 2) = "ExchangeType"
MSHFlexGrid5.TextMatrix(1, 3) = "Invoice No"
MSHFlexGrid5.TextMatrix(1, 4) = "Invoice Date"
MSHFlexGrid5.TextMatrix(1, 5) = "Service Tax"
MSHFlexGrid5.TextMatrix(1, 6) = "Amount"
MSHFlexGrid5.TextMatrix(1, 7) = "Discount"
MSHFlexGrid5.TextMatrix(1, 8) = "Premium"
MSHFlexGrid5.TextMatrix(1, 9) = "Batch ID"

MSHFlexGrid5.ColWidth(0) = 1800
MSHFlexGrid5.ColWidth(1) = 4000
MSHFlexGrid5.ColWidth(2) = 1700
MSHFlexGrid5.ColWidth(3) = 1500
MSHFlexGrid5.ColWidth(4) = 1500
MSHFlexGrid5.ColWidth(5) = 1500
MSHFlexGrid5.ColWidth(6) = 1500
End Sub



Private Sub Grid_HeadTotalBOE()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 9

MSHFlexGrid2.TextMatrix(1, 0) = "Location"
MSHFlexGrid2.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid2.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid2.TextMatrix(1, 3) = "Invoice No"
MSHFlexGrid2.TextMatrix(1, 4) = "Invoice Date"
MSHFlexGrid2.TextMatrix(1, 5) = "Service Tax"
MSHFlexGrid2.TextMatrix(1, 6) = "Amount"
MSHFlexGrid2.TextMatrix(1, 7) = "Discount"
MSHFlexGrid2.TextMatrix(1, 8) = "Premium"

MSHFlexGrid2.ColWidth(0) = 1200
MSHFlexGrid2.ColWidth(1) = 4000
MSHFlexGrid2.ColWidth(2) = 1500
MSHFlexGrid2.ColWidth(3) = 1700
MSHFlexGrid2.ColWidth(4) = 1500
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500


End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 14
MSHFlexGrid1.FixedRows = 2

MSHFlexGrid1.TextMatrix(1, 0) = "Month-Year"
MSHFlexGrid1.TextMatrix(1, 1) = "State"
MSHFlexGrid1.TextMatrix(1, 2) = "Location"
MSHFlexGrid1.TextMatrix(1, 3) = "Min BOD"
MSHFlexGrid1.TextMatrix(1, 4) = "Date on which monthly invoicing process is initiated by CMG"
MSHFlexGrid1.TextMatrix(1, 5) = "No. of BOEs required to be generated by CMPM"
MSHFlexGrid1.TextMatrix(1, 6) = "No. of BOEs Generated By CMPM"
MSHFlexGrid1.TextMatrix(1, 7) = "No. of BOEs pending with CMPM"
MSHFlexGrid1.TextMatrix(1, 8) = "No. of BOEs processed by FAG"
MSHFlexGrid1.TextMatrix(1, 9) = "No. of BOEs pending with FAG"
MSHFlexGrid1.TextMatrix(1, 10) = "No. of BOEs for which TI is uploaded into SM by FAG"
MSHFlexGrid1.TextMatrix(1, 11) = "No. of BOEs for which TI is yet to be uploaded by FAG"
MSHFlexGrid1.TextMatrix(1, 12) = "Name of AM"
MSHFlexGrid1.TextMatrix(1, 13) = "Max BOD"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1500

MSHFlexGrid1.ColWidth(12) = 3500

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(1) = 1500

End Sub

'Private Sub TxtToDate_LostFocus()
'mTotalDays = Day(DateSerial(Year(TxtToDate.Value), Month(TxtToDate.Value) + 1, 0))
'TxtToDate.Value = (Year(TxtToDate.Value) & "-" & Month(TxtToDate.Value) & "-" & mTotalDays)
'End Sub

