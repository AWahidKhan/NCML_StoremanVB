VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_SecurityAttendance 
   Caption         =   "Security Attendance"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9960
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15120
      TabIndex        =   9
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9855
      Left            =   0
      TabIndex        =   10
      Top             =   -120
      Width           =   14895
      Begin VB.Frame Frame1 
         Height          =   9615
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   14655
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
            Left            =   12645
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CmdReject 
            Caption         =   "Reject Attendance"
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
            Left            =   10710
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton CmdLeave 
            Caption         =   "Mark Absence"
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
            Left            =   8775
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TxtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   360
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox ChkCheckAll 
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
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   135
            TabIndex        =   8
            Top             =   1065
            Width           =   255
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
            ItemData        =   "FrmTxn_SecurityAttendance.frx":0000
            Left            =   2970
            List            =   "FrmTxn_SecurityAttendance.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   510
            Width           =   2880
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
            Height          =   375
            Left            =   6105
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton CmdApprove 
            Caption         =   "Mark Attendance"
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
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   375
            Left            =   90
            TabIndex        =   1
            Top             =   510
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
            Format          =   111804417
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   8025
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1470
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   14155
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
            Top             =   510
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
            Format          =   111804417
            CurrentDate     =   36494
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3375
            Left            =   120
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   6120
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   5953
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
            TabIndex        =   29
            Top             =   5760
            Width           =   90
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
            TabIndex        =   28
            Top             =   5760
            Width           =   2325
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
            TabIndex        =   26
            Top             =   5760
            Width           =   9360
         End
         Begin VB.Label Label8 
            BackColor       =   &H008080FF&
            Height          =   255
            Left            =   9360
            TabIndex        =   25
            Top             =   840
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
            Left            =   9840
            TabIndex        =   24
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFC0&
            Height          =   255
            Left            =   6480
            TabIndex        =   23
            Top             =   840
            Width           =   375
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
            Left            =   6960
            TabIndex        =   22
            Top             =   840
            Width           =   1020
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
            Left            =   1860
            TabIndex        =   20
            Top             =   225
            Width           =   735
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
            Left            =   8520
            TabIndex        =   19
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0E0FF&
            Height          =   255
            Left            =   8040
            TabIndex        =   18
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Select All"
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
            Left            =   420
            TabIndex        =   17
            Top             =   1050
            Width           =   840
         End
         Begin VB.Label lblstateregionlocation 
            AutoSize        =   -1  'True
            Caption         =   "Location"
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
            Left            =   3960
            TabIndex        =   16
            Top             =   225
            Width           =   900
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
            Left            =   10800
            TabIndex        =   15
            Top             =   840
            Width           =   2325
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
            Left            =   13200
            TabIndex        =   14
            Top             =   840
            Width           =   90
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
            TabIndex        =   13
            Top             =   225
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_SecurityAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mSelRow, mSelCol, mReject As Variant
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub ChkCheckAll_Click()
If tmp = 0 Then Exit Sub
If MSHFlexGrid1.Rows = 2 Then
    ChkCheckAll.Value = vbUnchecked
    Exit Sub
End If
For I = 1 To MSHFlexGrid1.Rows - 2
    MSHFlexGrid1.row = I
    MSHFlexGrid1.Col = 0
    If MSHFlexGrid1.CellBackColor = -2147483643 Then   '<> &H8080FF Then '''&HC0FFC0 Then
        If ChkCheckAll.Value = vbChecked Then
           If Format(Date, "yyyy-mm-dd") > "2013-07-09" Then
              If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
                 strSplitString = Split(MSHFlexGrid1.TextMatrix(I, 1), ",")
                 AttDate = Trim(strSplitString(0))
                 'Commented on 13 Sept 2013 After mail from Mayur Sir and Sankar Sir
                 'If CDate(AttDate) < Date - 7 Then
                 If CDate(AttDate) < Date - Val(Gen_BlockingSecurityDay) Then
                    MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
                 Else
                    MSHFlexGrid1.TextMatrix(I, 0) = strChecked
                 End If
              Else
                 MSHFlexGrid1.TextMatrix(I, 0) = strChecked
              End If
           Else
              MSHFlexGrid1.TextMatrix(I, 0) = strChecked
           End If
        Else
           MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
        End If
    End If
Next
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

'Commeted on 09 Dec 2013 After mail from Preeti and approval from Mayur Sir
'If Gen_WcLocation = "" Then
'    Call TableMst_Location
'Else
'    Call TableMst_Location(" WHERE " & Gen_WcLocation)
'End If

mLastToDate = GetLastDateOfMonth(TxtToDate.Value)

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where LocationID not in (Select Distinct LocationID from Txn_MonthwiseSecLoc Where MonthYear >= '" & Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-" & "01" & "' And MonthYear <= '" & Format(mLastToDate, Gen_DatFormat) & "') ")
Else
   Call TableMst_Location(" WHERE " & Gen_WcLocation & " And LocationID not in (Select Distinct LocationID from Txn_MonthwiseSecLoc Where MonthYear >= '" & Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-" & "01" & "' And MonthYear <= '" & Format(mLastToDate, Gen_DatFormat) & "') ")
End If


CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   'CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!Locationname
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
If CmbLocation.Text <> tmp Then
    Call Grid_Head
    Call Grid_Head2
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "SecurityAttendance.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "SecurityAttendance.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        MSHFlexGrid1.row = I
        MSHFlexGrid1.Col = C
                
        If MSHFlexGrid1.CellBackColor = &H8080FF Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = &H8080FF
        ElseIf MSHFlexGrid1.CellBackColor = &HC0FFC0 Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = &HC0FFC0
        ElseIf MSHFlexGrid1.CellBackColor = &HC0E0FF Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = &HC0E0FF
        End If
              
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

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

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtFromDate.Value, "yyyy-mm") Then
   MsgBox "From Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtToDate.Value, "yyyy-mm") Then
   MsgBox "To Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!!"
   TxtToDate.SetFocus
   Exit Sub
End If

If CmbLocation.Text = "" Then
    MsgBox "Blank Location not Allowed!!!"
    CmbLocation.SetFocus
    Exit Sub
End If

MsgBox "Wait till Next Message!!!"

Call Grid_Head
Call Grid_Head2

tmp = " Select"
tmp = tmp & " distinct MSLM.SecurityAgencyID, MSS.SecurityShifts SecurityType, MSD.SecurityDesignation, MSA.AgencyName, MSLMD.TxnDetailID, MSLMD.StartDate, MSLMD.EndDate,MSLMD.RequestID"
tmp = tmp & " From Mst_SecLocMappingDetail MSLMD"
tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID"
tmp = tmp & " Inner Join Mst_SecLocGodMappingDetail MSLGMD on MSLGMD.TxnDetailID = MSLMD.TxnDetailID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSLM.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLMD.SecurityDesignationID"
'Excluding NCML Security Agency
tmp = tmp & " Where MSA.SecurityAgencyID <> 80 And MSLM.LocationID = " & mLocationID & " And ('" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' Between MSLMD.StartDate And MSLMD.SalaryPaidUpto Or '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "' Between MSLMD.StartDate And MSLMD.SalaryPaidUpto)"
tmp = tmp & " And MSLMD.TxnDetailID not in (Select a.SecLocMapDetailID from Txn_MonthwiseSecLocDetail a"
tmp = tmp & " Inner Join Txn_MonthwiseSecLoc b on a.MonthwiseTxnID = b.MonthwiseTxnID"
tmp = tmp & " Where b.MonthYear = '" & Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-" & "01" & "' )"

Call TmpTable


tmp = " Select"
tmp = tmp & " MSLM.SecurityAgencyID, MSS.SecurityShifts SecurityType, MSD.SecurityDesignation, MSA.AgencyName, MSLMD.TxnDetailID, MSLMD.StartDate, MSLMD.EndDate,MSLMD.RequestID"
tmp = tmp & " From Mst_SecLocMappingDetail MSLMD"
tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSLM.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLMD.SecurityDesignationID"
tmp = tmp & " Where MSA.SecurityAgencyID <> 80 And MSLM.LocationID = " & mLocationID & " And ('" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' Between MSLMD.StartDate And MSLMD.SalaryPaidUpto Or '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "' Between MSLMD.StartDate And MSLMD.SalaryPaidUpto)"
tmp = tmp & " And MSLMD.TxnDetailID not in (select isnull(TxnDetailID,0) from Mst_SecLocGodMappingDetail) "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   Label24.Caption = TmpRs1.RecordCount
   With MSHFlexGrid2
        .Rows = .Rows + TmpRs1.RecordCount
        For I = 1 To TmpRs1.RecordCount
            .TextMatrix(I, 0) = IIf(IsNull(TmpRs1!TxnDetailID), "", TmpRs1!TxnDetailID) '"Security ID"
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs1!SecurityType), "", TmpRs1!SecurityType) '"Security Type"
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs1!SecurityDesignation), "", TmpRs1!SecurityDesignation)  '"Security Designation"
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs1!AgencyName), "", TmpRs1!AgencyName) '"Name"
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs1!RequestID), "", TmpRs1!RequestID) '"Name"
            TmpRs1.MoveNext
        Next
    End With
    MSHFlexGrid1.Height = 4215
Else
    MSHFlexGrid1.Height = 8025
End If

If TmpRs.RecordCount > 0 Then
   tmp = " Select TSA.AttendanceID, TSA.AttendanceDate, TSA.TxnDEtailID, TSA.GuardName, TSA.Status, TSA.Remarks, "
   tmp = tmp & " ME.EmployeeName+'~'+TSA.CreatedBy CreatedBy , TSA.CreatedDate, ME1.EmployeeName+'~'+TSA.UpdatedBy UpdatedBy, TSA.UpdatedDate,MSLMD.RequestID"
   tmp = tmp & " From Txn_securityAttendance TSA"
   tmp = tmp & " Inner Join Mst_SecLocMappingDetail MSLMD on MSLMD.TxnDetailID = TSA.TxnDetailID"
   tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID "
   tmp = tmp & " Left Join Mst_Employee ME on ME.EmployeeNo = TSA.CreatedBy"
   tmp = tmp & " Left Join Mst_Employee ME1 on ME1.EmployeeNo = TSA.UpdatedBy"
   tmp = tmp & " Where MSLM.LocationID = " & mLocationID
   tmp = tmp & " Order by TSA.AttendanceDate Desc"
    ''''And TSA.AttendanceDate between '" & Format(TxtFromDate.Value, "yyyy-mm-dd") & "' and '" & Format(TxtToDate.Value, "yyyy-mm-dd") & "'"
   Call Tmp1Table
   mTotalDays = DateDiff("d", TxtFromDate.Value, TxtToDate.Value) + 1
   mNextDate = TxtFromDate.Value
   With MSHFlexGrid1
        For j = 1 To mTotalDays
            TmpRs.MoveFirst
            For l = 1 To TmpRs.RecordCount
                I = MSHFlexGrid1.Rows - 1
                'TmpRs.Filter = " StartDate <= '" & mNextDate & "' And EndDate >= '" & mNextDate & "' "
                If mNextDate >= TmpRs!StartDate And mNextDate <= TmpRs!EndDate Then
                    .TextMatrix(I, 1) = Format(CDate(mNextDate), "dd-MMM-yyyy") & ", " & Format(CDate(mNextDate), "dddd") 'Date
                    .TextMatrix(I, 3) = IIf(IsNull(TmpRs!TxnDetailID), "", TmpRs!TxnDetailID) '"Security ID" '
                    .TextMatrix(I, 4) = IIf(IsNull(TmpRs!SecurityType), "", TmpRs!SecurityType) '"Security Type"
                    .TextMatrix(I, 5) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)  '"Security Designation"
                    .TextMatrix(I, 6) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName) '"Name"
                    .TextMatrix(I, 11) = IIf(IsNull(TmpRs!RequestID), "", TmpRs!RequestID) '"Name"
                    If TmpRs1.RecordCount > 0 Then
                       TmpRs1.Find "TxnDetailID = " & TmpRs!TxnDetailID
                       If Not TmpRs1.EOF Then
                          .TextMatrix(I, 2) = IIf(IsNull(TmpRs1!GuardName), "", TmpRs1!GuardName)
                       End If
                       TmpRs1.MoveFirst
                    End If
                    
                    .row = I
                    .Col = 0
                    .CellFontName = "Wingdings"
                    .CellFontSize = 12
                    .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(I, 0) = strUnChecked
                    
                    If TmpRs1.RecordCount > 0 Then
                       TmpRs1.Filter = "TxnDetailID = " & TmpRs!TxnDetailID & " And AttendanceDate = '" & mNextDate & "' "
                       If TmpRs1.RecordCount > 0 Then
                          .TextMatrix(I, 7) = IIf(IsNull(TmpRs1!AttendanceID), "", TmpRs1!AttendanceID)
                          .TextMatrix(I, 2) = IIf(IsNull(TmpRs1!GuardName), "", TmpRs1!GuardName)
                          .TextMatrix(I, 8) = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
                          .TextMatrix(I, 9) = IIf(IsNull(TmpRs1!CreatedBy), "", TmpRs1!CreatedBy) & " :- " & IIf(IsNull(TmpRs1!CreatedDate), "", TmpRs1!CreatedDate)
                          .TextMatrix(I, 10) = IIf(IsNull(TmpRs1!UpdatedBy), "", TmpRs1!UpdatedBy) & " :- " & IIf(IsNull(TmpRs1!UpdatedDate), "", TmpRs1!UpdatedDate)
                          .row = I
                          .Col = 0
                          .CellFontName = "Wingdings"
                          .CellFontSize = 12
                          .CellAlignment = flexAlignCenterCenter
                          '.TextMatrix(I, 0) = strChecked
                          For k = 0 To .Cols - 1
                              .Col = k
                              If TmpRs1!Status = "A" Then
                                 .CellBackColor = &HC0FFC0
                              ElseIf TmpRs1!Status = "R" Then
                                 .CellBackColor = &H8080FF
                              ElseIf TmpRs1!Status = "L" Then
                                 .CellBackColor = &HC0E0FF
                              End If
                          Next
                        End If
                        TmpRs1.Filter = ""
                    End If
                MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                End If
                TmpRs.MoveNext
            Next
            mNextDate = DateAdd("d", 1, CDate(mNextDate))
        Next
    End With
    Label22.Caption = MSHFlexGrid1.Rows - 2
    CmdApprove.Enabled = True
    CmdLeave.Enabled = True
    CmdReject.Enabled = True
    MsgBox "Done!!!"
Else
    MsgBox "No Record Found!!!"
End If
End Sub

Private Function CheckGuardName() As Boolean
CheckGuardName = True
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
        If MSHFlexGrid1.TextMatrix(I, 2) = "" Then
            CheckGuardName = False
            Exit Function
        End If
        If mReject = "Y" And MSHFlexGrid1.TextMatrix(I, 8) = "" Then
            CheckGuardName = False
            Exit Function
        End If
    End If
Next
End Function

Private Sub CmdApprove_Click()
mReject = "N"
If CheckGuardName = False Then
    MsgBox "Blank Guard Name not Allowed!!!"
    Exit Sub
End If

For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
        If MSHFlexGrid1.TextMatrix(I, 7) = "" Then
            RsTxn_SecurityAttendance.AddNew
            RsTxn_SecurityAttendance!AttendanceID = GetMaxAttendanceID
            RsTxn_SecurityAttendance!G_UID = GetGUID
            strSplitString = Split(MSHFlexGrid1.TextMatrix(I, 1), ",")
            AttDate = Trim(strSplitString(0))
            RsTxn_SecurityAttendance!AttendanceDate = Format(CDate(AttDate), "yyyy-mm-dd")
            RsTxn_SecurityAttendance!TxnDetailID = Val(MSHFlexGrid1.TextMatrix(I, 3))
            RsTxn_SecurityAttendance!CreatedBy = Gen_UserLoginID
            RsTxn_SecurityAttendance!CreatedDate = Now
        Else
            RsTxn_SecurityAttendance.MoveFirst
            RsTxn_SecurityAttendance.Find "AttendanceID = " & Val(MSHFlexGrid1.TextMatrix(I, 7))
            If Not RsTxn_SecurityAttendance.EOF Then
                RsTxn_SecurityAttendance!UpdatedBy = Gen_UserLoginID
                RsTxn_SecurityAttendance!UpdatedDate = Now
            End If
        End If
        RsTxn_SecurityAttendance!GuardName = MSHFlexGrid1.TextMatrix(I, 2)
        RsTxn_SecurityAttendance!Remarks = MSHFlexGrid1.TextMatrix(I, 8)
        RsTxn_SecurityAttendance!Status = "A"
        RsTxn_SecurityAttendance.Update
    End If
Next
RsTxn_SecurityAttendance.Requery
MsgBox "Done!!!"
Call CmdGo_Click
'Call Grid_Head
'TxtFromDate.Value = Date
'CmbLocation.Text = ""
End Sub

Private Function GetMaxAttendanceID() As Double
Dim Retval As Double
tmp = "Select max(AttendanceID) from Txn_SecurityAttendance"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxAttendanceID = Retval
End Function

Private Sub Command1_Click()

End Sub

Private Sub CmdLeave_Click()
mReject = "N"
If CheckGuardName = False Then
    MsgBox "Blank Guard Name not Allowed!!!"
    Exit Sub
End If
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
        If MSHFlexGrid1.TextMatrix(I, 7) = "" Then
            RsTxn_SecurityAttendance.AddNew
            RsTxn_SecurityAttendance!AttendanceID = GetMaxAttendanceID
            RsTxn_SecurityAttendance!G_UID = GetGUID
            strSplitString = Split(MSHFlexGrid1.TextMatrix(I, 1), ",")
            AttDate = Trim(strSplitString(0))
            RsTxn_SecurityAttendance!AttendanceDate = Format(CDate(AttDate), "yyyy-mm-dd")
            RsTxn_SecurityAttendance!TxnDetailID = Val(MSHFlexGrid1.TextMatrix(I, 3))
            RsTxn_SecurityAttendance!CreatedBy = Gen_UserLoginID
            RsTxn_SecurityAttendance!CreatedDate = Now
        Else
            RsTxn_SecurityAttendance.MoveFirst
            RsTxn_SecurityAttendance.Find "AttendanceID = " & Val(MSHFlexGrid1.TextMatrix(I, 7))
            If Not RsTxn_SecurityAttendance.EOF Then
                RsTxn_SecurityAttendance!UpdatedBy = Gen_UserLoginID
                RsTxn_SecurityAttendance!UpdatedDate = Now
            End If
        End If
        RsTxn_SecurityAttendance!GuardName = MSHFlexGrid1.TextMatrix(I, 2)
        RsTxn_SecurityAttendance!Remarks = MSHFlexGrid1.TextMatrix(I, 8)
        RsTxn_SecurityAttendance!Status = "L"
        RsTxn_SecurityAttendance.Update
    End If
Next
RsTxn_SecurityAttendance.Requery
MsgBox "Done!!!"
'Call Grid_Head
Call CmdGo_Click
End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdReject_Click()
mReject = "Y"
If CheckGuardName = False Then
    MsgBox "Blank Guard Name Or Remarks not Allowed!!!"
    Exit Sub
End If
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And MSHFlexGrid1.TextMatrix(I, 2) <> "" Then
        If MSHFlexGrid1.TextMatrix(I, 7) = "" Then
            RsTxn_SecurityAttendance.AddNew
            RsTxn_SecurityAttendance!AttendanceID = GetMaxAttendanceID
            RsTxn_SecurityAttendance!G_UID = GetGUID
            strSplitString = Split(MSHFlexGrid1.TextMatrix(I, 1), ",")
            AttDate = Trim(strSplitString(0))
            RsTxn_SecurityAttendance!AttendanceDate = Format(CDate(AttDate), "yyyy-mm-dd")
            RsTxn_SecurityAttendance!TxnDetailID = Val(MSHFlexGrid1.TextMatrix(I, 3))
            RsTxn_SecurityAttendance!CreatedBy = Gen_UserLoginID
            RsTxn_SecurityAttendance!CreatedDate = Now
        Else
            RsTxn_SecurityAttendance.MoveFirst
            RsTxn_SecurityAttendance.Find "AttendanceID = " & Val(MSHFlexGrid1.TextMatrix(I, 7))
            If Not RsTxn_SecurityAttendance.EOF Then
                RsTxn_SecurityAttendance!UpdatedBy = Gen_UserLoginID
                RsTxn_SecurityAttendance!UpdatedDate = Now
            End If
        End If
        RsTxn_SecurityAttendance!GuardName = MSHFlexGrid1.TextMatrix(I, 2)
        RsTxn_SecurityAttendance!Remarks = MSHFlexGrid1.TextMatrix(I, 8)
        RsTxn_SecurityAttendance!Status = "R"
        RsTxn_SecurityAttendance.Update
    End If
Next
RsTxn_SecurityAttendance.Requery
MsgBox "Done!!!"
Call CmdGo_Click
'Call Grid_Head

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

Private Sub MSHFlexGrid1_Click()
'If MSHFlexGrid1.CellBackColor = &HC0FFC0 Then Exit Sub
mSelRow = MSHFlexGrid1.RowSel
mSelCol = MSHFlexGrid1.ColSel
If mSelCol = 0 And mSelRow <> 0 And mSelRow <> MSHFlexGrid1.Rows - 1 Then
    'If MSHFlexGrid1.CellBackColor = &HC0FFC0 Or MSHFlexGrid1.CellBackColor = &H8080FF Or MSHFlexGrid1.CellBackColor = &HC0E0FF Then Exit Sub
    If MSHFlexGrid1.CellBackColor = &H8080FF Then Exit Sub
    If Format(Date, "yyyy-mm-dd") > "2013-07-09" Then
       If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "admin" Then
          strSplitString = Split(MSHFlexGrid1.TextMatrix(mSelRow, 1), ",")
          AttDate = Trim(strSplitString(0))
          
          'Commented on 13 Sept 2013 After mail from Mayur Sir and Sankar Sir
          'If CDate(AttDate) < Date - 7 Then
          If CDate(AttDate) < Date - Val(Gen_BlockingSecurityDay) Then
             MsgBox "For marking attendance beyond the past " & Val(Gen_BlockingSecurityDay) & " day(s), you need to provide the details to CMG with reasons for such delay in marking the attendance. CMG will mark the attendance on your behalf, if they are satisfied with your explanation!!!"
             Exit Sub
          End If
       End If
    End If
    If MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked Then
        MSHFlexGrid1.TextMatrix(mSelRow, 0) = strUnChecked
        tmp = 0
        ChkCheckAll.Value = vbUnchecked
        tmp = 1
    Else
        MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked
    End If
ElseIf (mSelCol = 2 Or mSelCol = 8) And mSelRow <> 0 And mSelRow <> MSHFlexGrid1.Rows - 1 And MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked Then
    With TxtName
        .Visible = True
        .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
        .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
        .Width = MSHFlexGrid1.CellWidth - 10
        .Height = MSHFlexGrid1.CellHeight - 10
        .Text = MSHFlexGrid1.Text
        If mSelCol = 2 Then
            .MaxLength = 200
        Else
            .MaxLength = 2000
        End If
        .ZOrder
        .SetFocus
    End With
End If
End Sub

'Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'
'mSelRow = MSHFlexGrid1.RowSel
'mSelCol = MSHFlexGrid1.ColSel
'If mSelCol = 0 And mSelRow <> 0 And mSelRow <> MSHFlexGrid1.Rows - 1 Then
'    If MSHFlexGrid1.CellBackColor = &HC0FFC0 Then Exit Sub
'    If MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked Then
'        MSHFlexGrid1.TextMatrix(mSelRow, 0) = strUnChecked
'        tmp = 0
'        ChkCheckAll.Value = vbUnchecked
'        tmp = 1
'    Else
'        MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked
'    End If
'ElseIf mSelCol = 7 And mSelRow <> 0 And mSelRow <> MSHFlexGrid1.Rows - 1 And MSHFlexGrid1.TextMatrix(mSelRow, 0) = strChecked Then
'    With TxtName
'        .Visible = True
'        .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
'        .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
'        .Width = MSHFlexGrid1.CellWidth - 10
'        .Height = MSHFlexGrid1.CellHeight - 10
'        .Text = MSHFlexGrid1.Text
'        .ZOrder
'        .SetFocus
'    End With
'End If
'End Sub

Private Sub TxtFromDate_GotFocus()
tmp = TxtFromDate.Value
End Sub

Private Sub TxtFromDate_LostFocus()
TxtToDate.Value = TxtFromDate.Value
If TxtFromDate.Value <> tmp Then
    Call Grid_Head
    Call Grid_Head2
End If
End Sub

Private Sub TxtName_LostFocus()
MSHFlexGrid1.TextMatrix(mSelRow, mSelCol) = TxtName.Text
TxtName.Visible = False
MSHFlexGrid1.SetFocus
End Sub

Private Sub TxtToDate_GotFocus()
tmp = TxtToDate.Value
End Sub

Private Sub TxtToDate_LostFocus()
If TxtToDate.Value <> tmp Then
    Call Grid_Head
    Call Grid_Head2
End If
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
Call Grid_Head
Call Grid_Head2
TxtFromDate.Value = Date
TxtToDate.Value = Date
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call TableTxn_SecurityAttendance
'If LCase(Gen_UserRole) = "power" Then
'   ChkAll.Visible = True
'   lblAll.Visible = True
'Else
'   ChkAll.Visible = False
'   lblAll.Visible = False
'End If
'CmbMonth.Text = MonthName(Month(Date))
'CmbYear.Text = Year(Date)
End Sub

Private Sub Grid_Head()
ChkCheckAll.Value = vbUnchecked
Label22.Caption = ""
CmdApprove.Enabled = False
CmdLeave.Enabled = False
CmdReject.Enabled = False
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 12 '11
    .AllowUserResizing = flexResizeBoth
    '.WordWrap = True
    .TextMatrix(0, 0) = "Selection"
    .TextMatrix(0, 1) = "Date" '"12-June-2012, Friday" '
    .TextMatrix(0, 2) = "Guard Name"
    .TextMatrix(0, 3) = "Security ID"
    .TextMatrix(0, 4) = "Security Type"
    .TextMatrix(0, 5) = "Security Designation"
    .TextMatrix(0, 6) = "Agency Name"
    .TextMatrix(0, 7) = "AttendanceID"
    .TextMatrix(0, 8) = "Remarks"
    .TextMatrix(0, 9) = "Created By"
    .TextMatrix(0, 10) = "Updated By"
    .TextMatrix(0, 11) = "Request ID"
        
    .ColWidth(0) = 1000
    .ColWidth(1) = 2400
    .ColWidth(2) = 2500
    .ColWidth(3) = 1100
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    .ColWidth(6) = 2500
    .ColWidth(7) = 0
    .ColWidth(8) = 2500
    .ColWidth(9) = 3000
    .ColWidth(10) = 3000
            
    .RowHeight(0) = 400
End With
End Sub

Private Sub Grid_Head2()
Label24.Caption = ""
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 5 '4
    .AllowUserResizing = flexResizeBoth
    '.WordWrap = True
    .TextMatrix(0, 0) = "Security ID"
    .TextMatrix(0, 1) = "Security Type"
    .TextMatrix(0, 2) = "Security Designation"
    .TextMatrix(0, 3) = "Agency Name"
    .TextMatrix(0, 4) = "Request ID"
        
    .ColWidth(0) = 1100
    .ColWidth(1) = 3500
    .ColWidth(2) = 3500
    .ColWidth(3) = 3500
            
    .RowHeight(0) = 400
End With
End Sub

Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

