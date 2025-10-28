VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_MonthwiseGodownRent 
   Caption         =   "Month wise Godown Rent"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   135
      TabIndex        =   11
      Top             =   1545
      Width           =   14970
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   7545
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   255
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   12726
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
      Height          =   1425
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   15000
      Begin VB.CheckBox Check1 
         Caption         =   "Only Paid Rent"
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
         Left            =   4320
         TabIndex        =   21
         Top             =   1013
         Width           =   1695
      End
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
         Left            =   1395
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   2580
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
         Left            =   8160
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   510
         Width           =   3465
      End
      Begin VB.ComboBox CmbStateName 
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
         Left            =   2745
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   510
         Width           =   2790
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
         Height          =   350
         Left            =   14010
         TabIndex        =   8
         Top             =   965
         Width           =   885
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
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
         Left            =   12240
         TabIndex        =   6
         Top             =   965
         Width           =   885
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
         Height          =   350
         Left            =   13125
         TabIndex        =   7
         Top             =   965
         Width           =   885
      End
      Begin VB.ComboBox CmbLocationName 
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
         Left            =   5625
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   510
         Width           =   2460
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   115081219
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1395
         TabIndex        =   1
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   115081219
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Added Accrual Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7080
         TabIndex        =   23
         Top             =   1020
         Width           =   2370
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   22
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   13920
         TabIndex        =   20
         Top             =   195
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
         Left            =   11520
         TabIndex        =   19
         Top             =   195
         Width           =   2325
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
         Left            =   9375
         TabIndex        =   18
         Top             =   195
         Width           =   1035
      End
      Begin VB.Label LblGodown 
         AutoSize        =   -1  'True
         Caption         =   "Godown No"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   495
         TabIndex        =   16
         Top             =   195
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To "
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
         Left            =   1845
         TabIndex        =   15
         Top             =   195
         Width           =   300
      End
      Begin VB.Label Label4 
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
         Left            =   3608
         TabIndex        =   14
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location Name"
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
         Left            =   6173
         TabIndex        =   10
         Top             =   195
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FrmRpt_MonthwiseGodownRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLOcationID, mCMPID, mGodownID, mStartDate, mEndDate, wc As Variant

Private Sub CmbLocationName_GotFocus()
If CmbStateName.Text = "" Then
   MsgBox "Plase Select State !!!"
   CmbStateName.SetFocus
   Exit Sub
End If

tmp = CmbLocationName.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location(" where StateID=" & mStateID)
Else
   Call TableMst_Location("where StateID=" & mStateID & " and " & Gen_WcSMLocation)
End If

CmbLocationName.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocationName.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationName.Text = tmp

End Sub

Private Sub CmbLocationName_LostFocus()
If CmbLocationName.Text = "" Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If


RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationName.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid PreFix Selection !!!!"
   CmbLocationName.SetFocus
   Exit Sub
End If
If tmp <> CmbLocationName.Text Then
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
End If
mLOcationID = RsMst_Location!LocationID
End Sub

Private Sub CmbCMPName_GotFocus()
If CmbLocationName.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbLocationName.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" Where LocationID = " & mLOcationID)
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp

End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   CmbGodown.Text = ""
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
If tmp <> CmbCMPName.Text Then
   CmbGodown.Text = ""
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown("Where UnitTypeID in (1,6) And CMPID = " & mCMPID)
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!godownno
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then Exit Sub

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbStateName_GotFocus()
tmp = CmbStateName.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateName.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateName.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateName.Text = tmp
End Sub

Private Sub CmbStateName_LostFocus()
If CmbStateName.Text = "" Then
   CmbLocationName.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateName.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateName.SetFocus
   Exit Sub
End If
If tmp <> CmbStateName.Text Then
   CmbLocationName.Text = ""
   CmbCMPName.Text = ""
   CmbGodown.Text = ""
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
TxtToDate.Value = Date
TxtFromDate.Value = Date - 365
CmbStateName.Text = ""
CmbLocationName.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_MonthwiseGodownRent.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_MonthwiseGodownRent.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.row = I
        MSHFlexGrid1.Col = C
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
        If MSHFlexGrid1.CellBackColor = -2147483643 Or MSHFlexGrid1.CellBackColor = -2147483633 Then
            oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbWhite
        Else
            oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = MSHFlexGrid1.CellBackColor
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

If Format(TxtFromDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future Date not allowed!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If
If Format(TxtToDate.Value, "yyyy-mm") > Format(Date, "yyyy-mm") Then
   MsgBox "Future Date not allowed!!!"
   TxtToDate.SetFocus
   Exit Sub
End If
If Format(TxtFromDate.Value, "yyyy-mm") > Format(TxtToDate.Value, "yyyy-mm") Then
   MsgBox "From Date should not greater than To Date!!!"
   TxtToDate.SetFocus
   Exit Sub
End If

wc = ""
If CmbGodown.Text <> "" Then
   wc = " Where MG.GodownID = " & mGodownID
ElseIf CmbCMPName.Text <> "" Then
   wc = " Where MCMP.CMPID = " & mCMPID
Else
   If CmbLocationName.Text <> "" Then
      wc = " Where ML.LocationID = " & mLOcationID
   ElseIf Gen_WcLocation <> "" Then
      wc = " Where ML." & Gen_WcLocation
   ElseIf CmbStateName.Text <> "" Then
      wc = " Where MS.StateID = " & mStateID
   ElseIf Gen_WcState <> "" Then
      wc = " Where MS." & Gen_WcState
   End If
End If

MsgBox "Wait till next message!!!"
Call Grid_Head
mStartDate = Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01"
mEndDate = Format(DateSerial(Year(TxtToDate.Value), Month(TxtToDate.Value) + 1, 0), "yyyy-mm-dd")
MSHFlexGrid1.Cols = MSHFlexGrid1.Cols + DateDiff("m", CDate(mStartDate), CDate(mEndDate)) + 1
For j = 10 To MSHFlexGrid1.Cols - 1 '9 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.ColWidth(j) = 1500
    MSHFlexGrid1.TextMatrix(1, j) = Format(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1 + j - 10, 0), "MMM-yyyy") 'Format(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1 + j - 9, 0), "MMM-yyyy")
Next

tmp = " Select MS.StateName, ML.LocationID, ML.LocationName, MCMP.CMPName, MG.GodownID, MG.GodownNo, MG.Address, MG.StartDate, MG.CloseDate, MUT.UnitType,  MG.GodownCapacity, MS.StateID, "
tmp = tmp & " (select top 1 advancegiven from Txn_LessorAgreement TLA inner join Txn_GodownLessorAgreeDetail tlad on tlad.laid = tla.laid where tlad.godownid = mg.godownid order by startdate desc) securitydeposite"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If wc <> "" Then
    tmp = tmp & wc & " And " '(MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
Else
    tmp = tmp & " Where " '(MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
End If
tmp = tmp & " (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'Commented on 29 Jan 2014 during meeting in conf room
'tmp = tmp & " And MUT.UnitTypeID in (1,6) " ' 1 and 6 means franchisee or leased franchisee
tmp = tmp & " And MUT.UnitTypeID in (1) "
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Call TmpTable

LblTotalRecords.Caption = TmpRs.RecordCount

If TmpRs.RecordCount > 0 Then
    CmdExcel.Enabled = True
    With MSHFlexGrid1
        .Rows = .Rows + TmpRs.RecordCount
        ProgressBar1.Max = .Rows
        For I = 2 To TmpRs.RecordCount + 1
            .TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName) '"State Name"
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location Name"
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) '"CMP Name"
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!godownno), "", TmpRs!godownno) '"Godown No"
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy")) '"Start Date"
            .TextMatrix(I, 5) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy")) '"Close Date"
            .TextMatrix(I, 6) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address) '"Godown Address"
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID) '"Godown ID"
            .TextMatrix(I, 8) = IIf(IsNull(TmpRs!securitydeposite), 0, TmpRs!securitydeposite) '"Security Deposite"
            .TextMatrix(I, 9) = IIf(IsNull(TmpRs!GodownCapacity), 0, TmpRs!GodownCapacity) '"Godown Capacity"
            .TextMatrix(0, 8) = Val(.TextMatrix(I, 8)) + Val(.TextMatrix(0, 8))
            TmpRs.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
    End With
    For j = 10 To MSHFlexGrid1.Cols - 1 '9 To MSHFlexGrid1.Cols - 1
        MonthwiseData (j)
    Next
    MsgBox "Done!!!"
    ProgressBar1.Value = ProgressBar1.Max
Else
    MsgBox "No record found!!!"
End If
End Sub

Private Sub MonthwiseData(x_ As Variant)

mStartDate = Format(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + x_ - 10, 1), "yyyy-mm-dd")
mEndDate = Format(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1 + x_ - 10, 0), "yyyy-mm-dd")
mTotalDays = Day(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1 + x_ - 10, 0))

tmp = " Select TGRD.GodownID, sum(TGRD.Amount) As IncomeAmt"
tmp = tmp & " From Txn_GodownRentDetail TGRD"
tmp = tmp & " Inner Join Txn_GodownRent TGR on TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TGRD.GodownID"
tmp = tmp & " Where TGR.Flag in ('A', 'P','S') And  TGR.ExpenseFrom Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by TGRD.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
    For I = 2 To MSHFlexGrid1.Rows - 2
        TmpRs1.MoveFirst
        TmpRs1.Find "GodownID = " & MSHFlexGrid1.TextMatrix(I, 7)
        If Not TmpRs1.EOF Then
            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
        End If
    Next
End If

If Check1.Value = vbChecked Then Exit Sub

tmp = " select TLAD.GodownID, MG.CLoseDate, TLA.StartDate , TLA.EXpiryDate, sum(TLAD.RentAmount) Amount  from Txn_GodownLessorAgreeDetail TLAD"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TLAD.GodownID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID = TLAD.LAID"
tmp = tmp & " Where TLA.ExpiryDate >= '" & mStartDate & "' and TLA.StartDate <= '" & mEndDate & "'  and  TLA.FLag in ('A', 'P','S')"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TLA.LAID not in"
tmp = tmp & " (Select LAID from Txn_GodownRentDetail TGRD"
tmp = tmp & " Inner Join Txn_GodownRent TGR on TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TGRD.GodownID"
tmp = tmp & " Where TGR.Flag in ('A', 'P','S') And"
tmp = tmp & " TGR.ExpenseFrom = '" & mStartDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'  )"
tmp = tmp & " Group by TLAD.GodownID, TLA.EXpiryDate, TLA.StartDate , MG.CLoseDate "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
    For I = 2 To MSHFlexGrid1.Rows - 2
        TmpRs1.MoveFirst
        TmpRs1.Find "GodownID = " & MSHFlexGrid1.TextMatrix(I, 7)
        If Not TmpRs1.EOF Then
            MSHFlexGrid1.row = I
            MSHFlexGrid1.Col = x_
            MSHFlexGrid1.CellBackColor = &HFF&
            If IsNull(TmpRs1!CloseDate) = True And Format(TmpRs1!StartDate, "yyyy-mm-dd") <= Format(mStartDate, "yyyy-mm-dd") Then
                MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
            Else
                If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") > Format(TmpRs1!CloseDate, "yyyy-mm-dd") Then
                    If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                        End If
                    Else
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!ExpiryDate) + 1
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            mdays = DateDiff("d", CDate(mStartDate), TmpRs1!ExpiryDate) + 1
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        End If
                    End If
                Else
                    If Format(TmpRs1!CloseDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                        End If
                    Else
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!CloseDate) + 1
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            mdays = DateDiff("d", CDate(mStartDate), TmpRs1!CloseDate) + 1
                            MSHFlexGrid1.TextMatrix(I, x_) = Val(MSHFlexGrid1.TextMatrix(I, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid1.TextMatrix(0, x_) = Val(MSHFlexGrid1.TextMatrix(0, x_)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        End If
                    End If
                End If
            End If
        End If
    Next
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click
End Sub

Private Sub Grid_Head()
ProgressBar1.Value = 0
LblTotalRecords.Caption = ""
CmdExcel.Enabled = False
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    .WordWrap = True
    .Cols = 10 '9
       
    .TextMatrix(1, 0) = "State Name"
    .TextMatrix(1, 1) = "Location Name"
    .TextMatrix(1, 2) = "CMP Name"
    .TextMatrix(1, 3) = "Godown No"
    .TextMatrix(1, 4) = "Start Date"
    .TextMatrix(1, 5) = "Close Date"
    .TextMatrix(1, 6) = "Address"
    .TextMatrix(1, 7) = "GodownID"
    .TextMatrix(1, 8) = "Security Deposite"
    .TextMatrix(1, 9) = "Capacity"
    
    .ColWidth(0) = 1900
    .ColWidth(1) = 1900
    .ColWidth(2) = 2200
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 4500
    .ColWidth(7) = 0
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    
    .RowHeight(1) = 600
End With
End Sub

