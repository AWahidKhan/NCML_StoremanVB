VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_NewGodownMapping 
   Caption         =   "New Godown Mapping"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   10680
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9765
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   14895
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   165
         TabIndex        =   5
         Top             =   120
         Width           =   14655
         Begin VB.ComboBox CmbGodown 
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
            TabIndex        =   24
            Top             =   1170
            Width           =   2685
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
            Left            =   10815
            TabIndex        =   12
            Top             =   1185
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
            Left            =   12030
            TabIndex        =   11
            Top             =   1185
            Width           =   1215
         End
         Begin VB.ComboBox CmbCMPName 
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
            Left            =   8580
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   480
            Width           =   5895
         End
         Begin VB.ComboBox CmbUnitTypeID 
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
            Left            =   2955
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   1170
            Width           =   3525
         End
         Begin VB.ComboBox CmbLocationID 
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
            Left            =   5610
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   2880
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
            Left            =   13245
            TabIndex        =   7
            Top             =   1185
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
            Left            =   2955
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   480
            Width           =   2595
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   90
            TabIndex        =   20
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
            Format          =   64552961
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   375
            Left            =   1470
            TabIndex        =   21
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
            Format          =   64552961
            CurrentDate     =   36494
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Godown"
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
            Left            =   1087
            TabIndex        =   25
            Top             =   885
            Width           =   750
         End
         Begin VB.Label Label3 
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
            Left            =   1770
            TabIndex        =   23
            Top             =   210
            Width           =   735
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
            Left            =   285
            TabIndex        =   22
            Top             =   210
            Width           =   945
         End
         Begin VB.Label LblLocationID 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   6675
            TabIndex        =   19
            Top             =   210
            Width           =   765
         End
         Begin VB.Label lblCMP 
            AutoSize        =   -1  'True
            Caption         =   "CMP Name"
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
            Left            =   11010
            TabIndex        =   18
            Top             =   210
            Width           =   1035
         End
         Begin VB.Label LblStateID 
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
            Left            =   3720
            TabIndex        =   17
            Top             =   210
            Width           =   1065
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
            Left            =   10920
            TabIndex        =   15
            Top             =   870
            Width           =   2295
         End
         Begin VB.Label Label2 
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
            Left            =   13320
            TabIndex        =   14
            Top             =   870
            Width           =   90
         End
         Begin VB.Label LblUnitTypeID 
            AutoSize        =   -1  'True
            Caption         =   "Unit Type"
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
            Left            =   4282
            TabIndex        =   13
            Top             =   885
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         Height          =   7695
         Left            =   165
         TabIndex        =   3
         Top             =   1740
         Width           =   14670
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7335
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   12938
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   240
         Left            =   165
         TabIndex        =   16
         Top             =   9480
         Width           =   14670
         _ExtentX        =   25876
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_NewGodownMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mUnitTypeID, mCMPID, mGodownID As Variant

Private Sub CmbCMPName_GotFocus()
If Trim(CmbLocationID.Text) = "" Then
   MsgBox "Please Select Location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
tmp = CmbCMPName.Text
If Gen_WcCMP = "" Then
   Call TableMst_CMP(" where LocationID=" & mLocationID)
Else
   Call TableMst_CMP(" where LocationID=" & mLocationID & " And " & Gen_WcCMP)
End If

CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
   CmbCMPName.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbCMPName_LostFocus()
If Trim(CmbCMPName.Text) = "" Then
   CmbGodown.Text = ""
   Exit Sub
End If
If tmp <> CmbCMPName.Text Then
   CmbGodown.Text = ""
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodown_GotFocus()
If Trim(CmbCMPName.Text) = "" Then
   MsgBox "Please Select CMP!!!"
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID)
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
   CmbGodown.AddItem RsMst_Godown!GodownNo
   RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If Trim(CmbGodown.Text) = "" Then Exit Sub

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!!"
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbLocationID_GotFocus()
If Trim(CmbStateID.Text) = "" Then
   MsgBox "Please Select State!!!"
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID= " & mStateID)
Else
   Call TableMst_Location(" Where StateID= " & mStateID & " And " & Gen_WcLocation)
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "Record not found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
   CmbLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()
If Trim(CmbLocationID.Text) = "" Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If
If tmp <> CmbLocationID.Text Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState = "" Then
   Call TableMst_State
Else
   Call TableMst_State(" Where " & Gen_WcState)
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
   CmbStateID.AddItem RsMst_State!StateName
   RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If Trim(CmbStateID.Text) = "" Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If
If tmp <> CmbStateID.Text Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
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

Private Sub CmbUnitTypeID_GotFocus()
tmp = CmbUnitTypeID.Text
Call TableMst_UnitType
CmbUnitTypeID.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_UnitType.RecordCount
   CmbUnitTypeID.AddItem RsMst_UnitType!UnitType
   RsMst_UnitType.MoveNext
Next
CmbUnitTypeID.Text = tmp
End Sub

Private Sub CmbUnitTypeID_LostFocus()
If CmbUnitTypeID.Text = "" Then Exit Sub

RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbUnitTypeID.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid selection!!!"
   CmbUnitTypeID.SetFocus
   Exit Sub
End If
mUnitTypeID = RsMst_UnitType!UnitTypeID
End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbUnitTypeID.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
Call Grid_Head
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
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
MsgBox "Wait till next message"
Call Grid_Head

'wc = " Where MG.StartDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"

wc = " Where (MG.UpdatedDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value + 1, "yyyy-mm-dd") & "' Or MG.Concurrency between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value + 1, "yyyy-mm-dd") & "')"

If CmbStateID.Text <> "" Then
   wc = wc & " And MS.StateID = " & mStateID
ElseIf Gen_WcState <> "" Then
   wc = wc & " And MS." & Gen_WcState
End If
If CmbLocationID.Text <> "" Then
   wc = wc & " And ML.LocationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   wc = wc & " And ML." & Gen_WcLocation
End If
If CmbCMPName.Text <> "" Then
   wc = wc & " And MCMP.CMPID = " & mCMPID
ElseIf Gen_WcCMP <> "" Then
   wc = wc & " And MCMP." & Gen_WcCMP
End If
If CmbGodown.Text <> "" Then
   wc = wc & " And MG.GodownID = " & mGodownID
End If
If CmbUnitTypeID.Text <> "" Then
   wc = wc & " And MG.UnitTypeID = " & mUnitTypeID
End If
tmp = " Select MG.GodownID, MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo, MG.Address, MG.StartDate,"
tmp = tmp & " MG.CloseDate, MU.UnitType, MCMP.CMPID,MG.Concurrency,MG.UpdatedDate "
tmp = tmp & " from Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MG.CMPID = MCMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MCMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & wc
tmp = tmp & " order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      ProgressBar1.Max = .Rows
      For i = 1 To TmpRs.RecordCount
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs!Concurrency), "", Format(TmpRs!Concurrency, "dd-MMM-yyyy"))
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!UpdatedDate), "", Format(TmpRs!UpdatedDate, "dd-MMM-yyyy"))
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
   MsgBox "Done!!!"
   ProgressBar1.Value = ProgressBar1.Max
Else
    MsgBox "No Record Found!!!"
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NewGodownMappingforInsurance.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NewGodownMappingforInsurance.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For i = 0 To MSHFlexGrid1.Rows - 1
   mRow = i + 1
   For c = 1 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, c) = MSHFlexGrid1.TextMatrix(i, c)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
Call CmdClear_Click
End Sub

Private Sub Grid_Head()
Label2.Caption = ""
ProgressBar1.Value = 0
CmdExcel.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 2
   .Cols = 12
   
   .TextMatrix(0, 0) = "Sr No"
   .TextMatrix(0, 1) = "State"
   .TextMatrix(0, 2) = "Location"
   .TextMatrix(0, 3) = "Unit Type"
   .TextMatrix(0, 4) = "CMP ID"
   .TextMatrix(0, 5) = "CMP Name  "
   .TextMatrix(0, 6) = "Godown No "
   .TextMatrix(0, 7) = "Godown Address "
   .TextMatrix(0, 8) = "Godown Start Date"
   .TextMatrix(0, 9) = "Godown End Date"
   .TextMatrix(0, 10) = "Godown Created Date"
   .TextMatrix(0, 11) = "Last Updated Date"
   
   .ColWidth(0) = 800
   .ColWidth(1) = 1600
   .ColWidth(2) = 1600
   .ColWidth(3) = 1500
   .ColWidth(4) = 800
   .ColWidth(5) = 2200
   .ColWidth(6) = 800
   .ColWidth(7) = 2500
   .ColWidth(8) = 1000
   .ColWidth(9) = 1000
   .ColWidth(10) = 1000
   .ColWidth(11) = 1000
   
   
   .RowHeight(0) = 600
   .WordWrap = True
End With
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

   Dim L As Double
   Dim a As Double
   Dim X As Double
   
   X = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * X) / 1000
   Frame4.Top = -L

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()

   Dim L As Double
   Dim a As Double
   Dim X As Double
   
   X = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * X) / 1000
   Frame4.Top = -L

'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim L As Double
   Dim a As Double
   Dim X As Double
   
   X = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * X) / 1000
   Frame4.Left = -L
End Sub

Private Sub HScroll1_Scroll()
   Dim L As Double
   Dim a As Double
   Dim X As Double
   
   X = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * X) / 1000
   Frame4.Left = -L
End Sub

