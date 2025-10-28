VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_SecurityAttendance 
   Caption         =   "Security Attendance Report"
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
      TabIndex        =   12
      Top             =   0
      Width           =   15015
      Begin VB.Frame Frame1 
         Height          =   9615
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   14775
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
            Left            =   2055
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   473
            Width           =   3525
         End
         Begin VB.ComboBox CmbStatus 
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
            ItemData        =   "FrmRpt_SecurityAttendance.frx":0000
            Left            =   120
            List            =   "FrmRpt_SecurityAttendance.frx":000D
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   1260
            Width           =   2745
         End
         Begin VB.TextBox TxtSecurityID 
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
            Left            =   2970
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1260
            Width           =   3000
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
            Left            =   13470
            TabIndex        =   8
            Top             =   1200
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
            Left            =   12255
            TabIndex        =   7
            Top             =   1200
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
            Height          =   375
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1200
            Width           =   1215
         End
         Begin VB.ComboBox CmbAgencyName 
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
            Left            =   9075
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   480
            Width           =   5640
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
            ItemData        =   "FrmRpt_SecurityAttendance.frx":0030
            Left            =   5610
            List            =   "FrmRpt_SecurityAttendance.frx":0032
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   480
            Width           =   3450
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   360
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
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
            CustomFormat    =   "MMMM-yyyy"
            Format          =   48300035
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7305
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2070
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12885
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   135
            Left            =   120
            TabIndex        =   27
            Top             =   9360
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label13 
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
            Left            =   3285
            TabIndex        =   31
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Status"
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
            Left            =   1215
            TabIndex        =   30
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Security ID"
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
            Left            =   3885
            TabIndex        =   29
            Top             =   960
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Agency Name"
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
            Left            =   11190
            TabIndex        =   28
            Top             =   225
            Width           =   1290
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   915
            TabIndex        =   26
            Top             =   225
            Width           =   1020
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
            Left            =   8640
            TabIndex        =   25
            Top             =   1320
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
            Left            =   6240
            TabIndex        =   24
            Top             =   1320
            Width           =   2325
         End
         Begin VB.Label lblstateregionlocation 
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
            Left            =   6953
            TabIndex        =   23
            Top             =   225
            Width           =   765
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0E0FF&
            Height          =   255
            Left            =   7800
            TabIndex        =   22
            Top             =   1700
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Absence"
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
            Left            =   8280
            TabIndex        =   21
            Top             =   1700
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Attendance"
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
            Left            =   6720
            TabIndex        =   20
            Top             =   1700
            Width           =   1020
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFC0&
            Height          =   255
            Left            =   6240
            TabIndex        =   19
            Top             =   1700
            Width           =   375
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Rejected"
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
            Left            =   9600
            TabIndex        =   18
            Top             =   1700
            Width           =   825
         End
         Begin VB.Label Label8 
            BackColor       =   &H008080FF&
            Height          =   255
            Left            =   9120
            TabIndex        =   17
            Top             =   1700
            Width           =   375
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   5760
            Width           =   90
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   10680
      Width           =   10935
   End
End
Attribute VB_Name = "FrmRpt_SecurityAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mAgencyID, mSecurityID, mStateID, mStatus As Variant
Dim mFromDate, mToDate As Date

Private Sub CmbAgencyName_GotFocus()
tmp = CmbAgencyName.Text
Call TableMst_SecurityAgency("")

CmbAgencyName.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Agency found"
   CmbAgencyName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbAgencyName.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbAgencyName.Text = tmp
End Sub

Private Sub CmbAgencyName_LostFocus()
If CmbAgencyName.Text = "" Then Exit Sub
   
RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbAgencyName.Text & "'"
If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Agency Selection "
   CmbAgencyName.SetFocus
   Exit Sub
End If

mAgencyID = RsMst_SecurityAgency!SecurityAgencyID
End Sub

Private Sub CmbLocation_GotFocus()
If CmbStateName.Text = "" Then
   MsgBox "Select State !!!"
   CmbLocation.Text = ""
   CmbStateName.SetFocus
   Exit Sub
End If

tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
    Call TableMst_Location(" Where StateID = " & mStateID & " ")
Else
    Call TableMst_Location(" WHERE StateID = " & mStateID & " And " & Gen_WcLocation)
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
    MsgBox "Invalid Selection!!!"
    CmbLocation.SetFocus
    Exit Sub
End If

mLocationID = RsMst_Location!LocationID
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
   CmbLocation.Text = ""
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
   CmbLocation.Text = ""
End If
mStateID = RsMst_State!StateID
End Sub


Private Sub CmbStatus_LostFocus()

If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) = "attendance" Then
   mStatus = "A"
ElseIf LCase(CmbStatus.Text) = "absence" Then
   mStatus = "L"
ElseIf LCase(CmbStatus.Text) = "rejected" Then
   mStatus = "R"
Else
   MsgBox "Invalid Selection!!!"
   CmbStatus.SetFocus
   Exit Sub
End If


End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbStateName.Text = ""
CmbLocation.Text = ""
CmbAgencyName.Text = ""
TxtFromDate.Value = Format(Date, "yyyy-MM-01")
CmbStatus.Text = ""
TxtSecurityID = ""

End Sub

Private Sub CmdExcel_Click()
Dim Path, strFilName, strFilLocation As Variant
Gen_mTimeStamp = GetTimeStamp
tmp = ""
ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If ans = vbYes Then
   Path = App.Path
   strFilName = "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityAttendance" & " .csv"
   strFilLocation = Path + strFilName
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
   Open strFilLocation For Output As #1
   With MSHFlexGrid1
   ProgressBar1.Value = 0
   ProgressBar1.Max = .Rows + 2
   For I = 0 To .Rows - 1
       tmp = ""
       For C = 0 To .Cols - 1
          If C = 0 Then
             tmp = tmp & Replace(Replace(.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          Else
             tmp = tmp & "," & Replace(Replace(.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          End If
       Next
       Print #1, tmp
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   Close #1
   End With
Else
    Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityAttendance.xls")
    Call Xls_Detail_Rpt
End If
MsgBox "Done !!!"



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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityAttendance.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = C
                
      If MSHFlexGrid1.CellBackColor = &H8080FF Then
         oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = &H8080FF
      ElseIf MSHFlexGrid1.CellBackColor = &HC0FFC0 Then
         oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = &HC0FFC0
      ElseIf MSHFlexGrid1.CellBackColor = &HC0E0FF Then
         oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = &HC0E0FF
      End If
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
oWB.Save
MsgBox ("Excel File Created!!!")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
'If TxtFromDate.Value > Date Then
'    MsgBox "Future Date not Allowed!!!"
'    TxtFromDate.SetFocus
'    Exit Sub
'End If
'If TxtToDate.Value > Date Then
'    MsgBox "Future Date not Allowed!!!"
'    TxtToDate.SetFocus
'    Exit Sub
'End If
'If TxtFromDate.Value > TxtToDate.Value Then
'    MsgBox "From Date should not greater than To Date!!!"
'    TxtToDate.SetFocus
'    Exit Sub
'End If

mFromDate = Format(TxtFromDate.Value, "yyyy-MM-01")
TxtFromDate.Value = mFromDate
mToDate = Format(TxtFromDate.Value, "yyyy-MM-" & Day(DateAdd("m", 1, TxtFromDate.Value - Day(TxtFromDate.Value) + 1) - 1) & "")

MsgBox "Wait till Next Message!!!"
Call Grid_Head

tmp = " Select TSA.AttendanceDate, MS.StateName,ML.LocationName, TSA.TxnDEtailID, TSA.GuardName, TSA.Status, TSA.Remarks, MSS.SecurityShifts SecurityType, MSD.SecurityDesignation, MSA.AgencyName,  "
tmp = tmp & " ME.EmployeeName+'~'+TSA.CreatedBy CreatedBy , TSA.CreatedDate, ME1.EmployeeName+'~'+TSA.UpdatedBy UpdatedBy, TSA.UpdatedDate"
tmp = tmp & " From Txn_securityAttendance TSA"
tmp = tmp & " Inner Join Mst_SecLocMappingDetail MSLMD on MSLMD.TxnDetailID = TSA.TxnDetailID"
tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID "
tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSLM.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLMD.SecurityDesignationID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MSLM.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Left Join Mst_Employee ME on ME.EmployeeNo = TSA.CreatedBy"
tmp = tmp & " Left Join Mst_Employee ME1 on ME1.EmployeeNo = TSA.UpdatedBy"
tmp = tmp & " Where TSA.AttendanceDAte between '" & Format(mFromDate, "yyyy-mm-dd") & "' and '" & Format(mToDate, "yyyy-mm-dd") & "'"

If CmbStateName.Text <> "" Then
   tmp = tmp & " and MS.StateName = '" & CmbStateName.Text & "'"
ElseIf Gen_WcState <> "" Then
   tmp = tmp & " and MS." & Gen_WcState
End If

If CmbLocation.Text <> "" Then
   tmp = tmp & " and MSLM.LocationID = " & mLocationID
ElseIf Gen_WcLocation <> "" Then
   tmp = tmp & " and MSLM." & Gen_WcLocation
End If

If CmbAgencyName.Text <> "" Then
   tmp = tmp & " and MSLM.SecurityAgencyID = " & mAgencyID
End If

If Trim(TxtSecurityID.Text) <> "" Then
   tmp = tmp & " and TSA.TxnDEtailID = " & Val(TxtSecurityID.Text)
End If

'mStatus
If CmbStatus.Text <> "" Then
   tmp = tmp & " and TSA.Status = '" & mStatus & "'"
End If

tmp = tmp & " Order by TSA.AttendanceDate Desc,MS.StateName,ML.LocationName"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   Label22.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      ProgressBar1.Max = .Rows
      For I = 1 To TmpRs.RecordCount
         .TextMatrix(I, 0) = IIf(IsNull(TmpRs!AttendanceDate), "", Format(TmpRs!AttendanceDate, "dd-MMM-yyyy"))
         .TextMatrix(I, 1) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         .TextMatrix(I, 3) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName) '"Name"
         .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GuardName), "", TmpRs!GuardName)
         .TextMatrix(I, 5) = IIf(IsNull(TmpRs!TxnDetailID), "", TmpRs!TxnDetailID) '"Security ID" '
         .TextMatrix(I, 6) = IIf(IsNull(TmpRs!SecurityType), "", TmpRs!SecurityType) '"Security Type"
         .TextMatrix(I, 7) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)  '"Security Designation"
         .TextMatrix(I, 8) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
         .TextMatrix(I, 9) = IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy & " :- " & TmpRs!CreatedDate)
         .TextMatrix(I, 10) = IIf(IsNull(TmpRs!UpdatedBy), "", TmpRs!UpdatedBy & " :- " & TmpRs!UpdatedDate)
         .TextMatrix(I, 11) = IIf(IsNull(TmpRs!Status), "", TmpRs!Status)
         
         If LCase(.TextMatrix(I, 11)) = "a" Then
            .TextMatrix(I, 11) = "Attendance"
         ElseIf LCase(.TextMatrix(I, 11)) = "l" Then
             .TextMatrix(I, 11) = "Absence"
         ElseIf LCase(.TextMatrix(I, 11)) = "r" Then
            .TextMatrix(I, 11) = "Rejected"
         End If
         
         Call GetAMRMName(I, .TextMatrix(I, 2))
         
         .row = I
         For k = 0 To .Cols - 1
            .Col = k
            If TmpRs!Status = "A" Then
               .CellBackColor = &HC0FFC0
            ElseIf TmpRs!Status = "R" Then
               .CellBackColor = &H8080FF
            ElseIf TmpRs!Status = "L" Then
               .CellBackColor = &HC0E0FF
            End If
         Next
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

Private Sub TxtFromDate_LostFocus()
'TxtToDate.Value = TxtFromDate.Value
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
    .Rows = 2
    .Cols = 14
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    
    .TextMatrix(0, 0) = "Date"
    .TextMatrix(0, 1) = "State Name"
    .TextMatrix(0, 2) = "Location Name"
    .TextMatrix(0, 3) = "Agency Name"
    .TextMatrix(0, 4) = "Guard Name"
    .TextMatrix(0, 5) = "Security ID"
    .TextMatrix(0, 6) = "Security Type"
    .TextMatrix(0, 7) = "Security Designation"
    .TextMatrix(0, 8) = "Remarks"
    .TextMatrix(0, 9) = "Created By"
    .TextMatrix(0, 10) = "Updated By"
    .TextMatrix(0, 11) = "Status"
    .TextMatrix(0, 12) = "AM"
    .TextMatrix(0, 13) = "RM"
        
    .ColWidth(0) = 1200
    .ColWidth(1) = 1800
    .ColWidth(2) = 2200
    .ColWidth(3) = 2200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 2500
    .ColWidth(8) = 2500
    .ColWidth(9) = 2500
End With
End Sub

Private Sub GetAMRMName(x_ As Variant, mLocName_ As Variant)

tmp = " Select ECM.LocationID, EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMName'"
tmp = tmp & " , EMP3.EmployeeName +'~'+EMP3.EmployeeNo as 'RMName'"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID And EMP.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP.DirectReportingTo = EMP2.EmployeeID And EMP2.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID And EMP3.Flag = 'Y'"
tmp = tmp & " Inner Join Mst_Location ML On ECM.LocationID = ML.LocationID "
tmp = tmp & " Where EMP2.EmployeeName Is Not Null And ML.LocationName = '" & mLocName_ & "' "
tmp = tmp & " Group By ECM.LocationID, EMP2.EmployeeName, EMP2.EmployeeNo"
tmp = tmp & " ,EMP3.EmployeeName ,EMP3.EmployeeNo"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x_, 12) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname)
   MSHFlexGrid1.TextMatrix(x_, 13) = IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName)
End If

End Sub

