VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_AMwiseIncomeMIS 
   Caption         =   "AM wise Income MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "AM Wise Details"
      TabPicture(0)   =   "FrmRpt_AMwiseIncomeMIS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vendor Expense"
      TabPicture(1)   =   "FrmRpt_AMwiseIncomeMIS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Godown Rent"
      TabPicture(2)   =   "FrmRpt_AMwiseIncomeMIS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "NN-NR Income"
      TabPicture(3)   =   "FrmRpt_AMwiseIncomeMIS.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Employee Expense"
      TabPicture(4)   =   "FrmRpt_AMwiseIncomeMIS.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Employee Salary"
      TabPicture(5)   =   "FrmRpt_AMwiseIncomeMIS.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Not Agree. Godowns"
      TabPicture(6)   =   "FrmRpt_AMwiseIncomeMIS.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame8"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "DP Income"
      TabPicture(7)   =   "FrmRpt_AMwiseIncomeMIS.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame9"
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame9 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
            Height          =   7935
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
      Begin VB.Frame Frame8 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
            Height          =   7935
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
      Begin VB.Frame Frame7 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   14790
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
            Height          =   7935
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
      Begin VB.Frame Frame6 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   14790
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   7935
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
      Begin VB.Frame Frame5 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   14790
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   7935
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
      Begin VB.Frame Frame4 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   14790
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   7935
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
      Begin VB.Frame Frame3 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   14790
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   7935
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
         Height          =   8295
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   14790
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   7935
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   13996
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
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   15015
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
         Left            =   7695
         TabIndex        =   3
         Top             =   270
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
         Left            =   6480
         TabIndex        =   2
         Top             =   270
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
         Left            =   5265
         TabIndex        =   1
         Top             =   270
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Top             =   315
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   61079553
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   3675
         TabIndex        =   5
         Top             =   285
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   61079553
         CurrentDate     =   39850
         MinDate         =   38353
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
         Left            =   120
         TabIndex        =   7
         Top             =   345
         Width           =   990
      End
      Begin VB.Label Label1 
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
         Left            =   2805
         TabIndex        =   6
         Top             =   330
         Width           =   780
      End
   End
End
Attribute VB_Name = "FrmRpt_AMwiseIncomeMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTotDays As Variant
Dim mdate As Variant
Dim mStartDate As Variant
Dim mEndDate As Variant
Dim mYear As Variant
Dim mMonthNo As Variant
Dim mTotOtherEmpSal As Double
Dim mTotOtherEmpID As String
Dim mDoublePaidAmt As Double
Dim mInvStartDate As Variant
Dim mInvEndDate As Variant
Dim mRow As Double
Private Sub CmdClear_Click()
Call Grid_Head
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
TxtFromDate = Date - 30
TxtToDate = Date
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_AMWise Income MIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_AMWise Income MIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Report"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Vendor Expense Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid2.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
        End If


    Next
Next


Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "Godown Rent Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid3.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
        End If


    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "NN-NR Income"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid4.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid4.Cols - 1
        If InStr(MSHFlexGrid4.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid4.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid4.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(I, c)
        End If
    Next
Next


Set oWS = oWB.Worksheets("Sheet5")
oWS.Name = "Employee Expense"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid5.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid5.Cols - 1
        If InStr(MSHFlexGrid5.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid5.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid5.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid5.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid5.TextMatrix(I, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet6")
oWS.Name = "Salary Details"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid6.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid6.Cols - 1
        If InStr(MSHFlexGrid6.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid6.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid6.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid6.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid6.TextMatrix(I, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet7")
oWS.Name = "Not Agree. Godowns"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid7.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid7.Cols - 1
        If InStr(MSHFlexGrid7.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid7.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid7.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid7.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid7.TextMatrix(I, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet8")
oWS.Name = "DP Income"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid8.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid8.Cols - 1
        If InStr(MSHFlexGrid8.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid8.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid8.TextMatrix(I, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid8.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid8.TextMatrix(I, c)
        End If
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim mCCExp, mBDoExp As Double
Dim mLocationIDs As String
Dim aftTime, bfrTime, curTime As Variant
Dim mStateID, mRM As Variant
aftTime = "19.00"
bfrTime = "09.30"
curTime = Format(Now, "hh.mm")
Dim mRow1 As Double

If LCase(Gen_UserRole) <> "power" Then
   If curTime < aftTime And curTime > bfrTime Then
      MsgBox "Please run this Report before 9.30 am or After 7.00 pm."
      Exit Sub
   End If
End If
If TxtFromDate > TxtToDate Then
   MsgBox "FromDate Should be Less Than Todate."
   TxtFromDate.SetFocus
   Exit Sub
End If

'Call TempDataTableDelete
'Call TempDataTable

mStartDate = TxtFromDate.Value
mEndDate = TxtToDate.Value


mInvStartDate = Year(mStartDate) & "-" & Month(mStartDate) & "-" & 1
mTotDays = Day(DateSerial(Year(mEndDate), Month(mEndDate) + 1, 0))
mInvEndDate = Year(mEndDate) & "-" & Month(mEndDate) & "-" & mTotDays

Call Grid_Head

'--For Location Wise AM

tmp = "Select ML.LocationID,ML.LocationName,ML.StateID,EMP1.EmployeeID AS 'CCAMID',EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as 'CCAM',MD.Designation ,  "
tmp = tmp & " EMP2.EmployeeID as 'AMRMID',EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMRM'"
tmp = tmp & " ,EMP3.EmployeeID as 'RMManagmentID',EMP3.EmployeeName +'~'+EMP3.EmployeeNo +'~'+ MD2.Designation as 'RM/Managment'"
tmp = tmp & " ,EMP4.EmployeeID as 'ManagmentID',EMP4.EmployeeName +'~'+EMP4.EmployeeNo as Managment"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD1 On EMP2.DesignationID = MD1.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD2 On EMP3.DesignationID = MD2.DesignationID"
tmp = tmp & " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID"
tmp = tmp & " Group By ML.LocationID,ML.LocationName,ML.StateID,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,"
tmp = tmp & " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation"
tmp = tmp & " ,EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID"
tmp = tmp & " Order By EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID,ML.LocationName"

Call TmpTable


tmp = "Select ML.LocationID,ML.LocationName,ML.StateID,EMP1.EmployeeID AS 'CCAMID',EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as 'CCAM',MD.Designation ,  "
tmp = tmp & " EMP2.EmployeeID as 'AMRMID',EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMRM'"
tmp = tmp & " ,EMP3.EmployeeID as 'RMManagmentID',EMP3.EmployeeName +'~'+EMP3.EmployeeNo +'~'+ MD2.Designation as 'RM/Managment'"
tmp = tmp & " ,EMP4.EmployeeID as 'ManagmentID',EMP4.EmployeeName +'~'+EMP4.EmployeeNo as Managment"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD1 On EMP2.DesignationID = MD1.DesignationID"
tmp = tmp & " Inner Join Mst_Designation MD2 On EMP3.DesignationID = MD2.DesignationID"
tmp = tmp & " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID"
tmp = tmp & " Where EMP1.EmployeeID Is Null"
tmp = tmp & " Group By ML.LocationID,ML.LocationName,ML.StateID,EMP1.EmployeeNo,EMP1.EmployeeName,EMP2.EmployeeNo,EMP2.EmployeeName,"
tmp = tmp & " EMP3.EmployeeNo , EMP3.EmployeeName, EMP4.EmployeeName, EMP4.EmployeeNo, MD.Designation, MD1.Designation, MD2.Designation"
tmp = tmp & " ,EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID"
tmp = tmp & " Order By EMP1.EmployeeID,EMP2.EmployeeID,EMP3.EmployeeID,EMP4.EmployeeID,ML.LocationName"

Call Tmp4Table


'--For All AMs

Call TableMst_Employee(" Where DesignationID= 3")

For I = 1 To RsMst_Employee.RecordCount
    mLocationIDs = ""
    TmpRs.Filter = "CCAMID=" & RsMst_Employee!EmployeeID
    mCCExp = 0
    mBDoExp = 0
If RsMst_Employee!EmployeeID = 312 Then
   test = test
End If
    If TmpRs.RecordCount > 0 Then
        mRM = IIf(IsNull(TmpRs!AMRM), "", TmpRs!AMRM)
        For j = 1 To TmpRs.RecordCount
            If CheckLocation(TmpRs!LocationName) = False Then
                mStateID = TmpRs!StateID
                mRow = MSHFlexGrid1.Rows - 1
                MSHFlexGrid1.TextMatrix(mRow, 0) = IIf(IsNull(TmpRs!AMRM), "", TmpRs!AMRM)
                MSHFlexGrid1.TextMatrix(mRow, 1) = TmpRs!LocationName
                If mLocationIDs = "" Then
                    mLocationIDs = TmpRs!LocationID
                Else
                   mLocationIDs = mLocationIDs & "," & TmpRs!LocationID
                End If
                MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                If mStateID = 20 Then
                  tet = test
                End If
                TmpRs4.Filter = "StateID=" & mStateID
                For k = 1 To TmpRs4.RecordCount
                   If CheckLocation(TmpRs4!LocationName) = False Then
                   ''mStateID = TmpRs!StateID
                      mRow = MSHFlexGrid1.Rows - 1
                      MSHFlexGrid1.TextMatrix(mRow, 0) = mRM 'IIf(IsNull(TmpRs!AMRM), "", TmpRs!AMRM)
                      MSHFlexGrid1.TextMatrix(mRow, 1) = TmpRs4!LocationName
                      If mLocationIDs = "" Then
                         mLocationIDs = TmpRs4!LocationID
                      Else
                         mLocationIDs = mLocationIDs & "," & TmpRs4!LocationID
                      End If
                      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                   End If
    
                  TmpRs4.MoveNext
               Next
               TmpRs4.Filter = ""
            End If

            TmpRs.MoveNext
            
        Next
        
      
        
        mRow1 = MSHFlexGrid1.Rows - 1
        '-------for BDO & CC of Locations
        If mLocationIDs <> "" Then
           Call FillCCBDO(mLocationIDs, , , , mRM)
        '--For AM
           Call FillCCBDO(mLocationIDs, "AM", RsMst_Employee!EmployeeID, RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo, mRM)
         '--For CC,BDO And AM Expense
           Call FillCCBDOexpense(mLocationIDs, mRow1)
        End If
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
      End If
       
    
    TmpRs.Filter = ""
    RsMst_Employee.MoveNext
Next
''RsMst_Employee.MoveFirst
''Dim mFlag As Boolean
''mFlag = False
''For I = 1 To TmpRs.RecordCount
''   For r = 1 To MSHFlexGrid1.Rows - 1
''       If MSHFlexGrid1.TextMatrix(r, 1) = TmpRs!LocationName Then
''          mFlag = True
''          Exit For
''       End If
''   Next
''   If mFlag = False Then
''      mRow = MSHFlexGrid1.Rows - 1
''      MSHFlexGrid1.TextMatrix(mRow, 0) = IIf(IsNull(TmpRs!AMRM), "", TmpRs!AMRM)
''      MSHFlexGrid1.TextMatrix(mRow, 1) = TmpRs!LocationName
''
''      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
''      mRow1 = MSHFlexGrid1.Rows - 1
''
''      '-------for BDO & CC of Locations
''      Call FillCCBDO(TmpRs!LocationID)
''      '-------for AM of Locations
''      Call FillCCBDO(TmpRs!LocationID, "AM")
''      '-------for BDO & CC & AM Expense
''      Call FillCCBDOexpense(TmpRs!LocationID, mRow1)
''
''      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
''   End If
''   TmpRs.MoveNext
''Next
Call dispdata
Call FillCMPMExpense
Call FillEmployeeExpense
Call TotalSalaryLocWise
Call FillVendorExpense
Call FillGodownRent
Call FillNNNCIncome
Call FillDPIncome
Call FillNotAgreeGodowns
CmdExcel.Enabled = True
End Sub
Private Sub FillVendorExpense()
Dim mTotExpAmt, mExpAmt, mTotExpense As Double
Dim mSEndDate As Date
Dim mSStartDate As Date
Dim mRow As Double
tmp = "Select TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,TVC.ExpenseFrom,TVC.ExpenseTo,Sum(TVCD.Amount) as TotalExpenseAmount"
tmp = tmp & " From Txn_VendorClaim TVC"
tmp = tmp & " inner join Txn_VendorClaimDetail TVCD on TVC.ClaimID=TVCD.ClaimID"
tmp = tmp & " inner Join Mst_ExpenseType ME on TVCD.ExpenseTypeID=ME.ExpenseTypeID"
tmp = tmp & " Inner Join Mst_Vendor MV on TVC.VendorID=MV.VendorID"
tmp = tmp & " inner join Mst_Location ML on TVCD.LocationID=ML.LocationID"
tmp = tmp & " Where  TVC.ExpenseFrom <= '" & mEndDate & "' And TVC.ExpenseTo >= '" & mStartDate & "' And TVC.Flag <> 'R' And TVCD.ExpenseTypeID Not in (8,69)"
tmp = tmp & " Group by TVCD.LocationID,ML.LocationName,MV.VendorName,ME.ExpenseType,TVC.ExpenseFrom,TVC.ExpenseTo"
tmp = tmp & " Order By ML.LocationName"

Call Grid_Head1
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
    mRow = 1
    mTotMonths = DateDiff("m", mStartDate, mEndDate)
    For I = 0 To mTotMonths
        mCurrMonth = DateAdd("m", I, mStartDate)
        mYear = Year(mCurrMonth)
        mMonthNo = Month(mCurrMonth)
        mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
        mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
        mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
        
        TmpRs2.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
        
        For r = 1 To TmpRs2.RecordCount
            mExpAmt = 0
            mTotDays = DateDiff("d", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo)
            mTotDays = mTotDays + 1
            If mTotDays <> 0 Then
               mTotExpense = TmpRs2!TotalExpenseAmount / mTotDays
               If TmpRs2!ExpenseFrom < mSStartDate And TmpRs2!ExpenseTo > mSEndDate Then
                  mExpDays = DateDiff("d", mSStartDate, mSEndDate)
               ElseIf TmpRs2!ExpenseFrom < mSStartDate And TmpRs2!ExpenseTo <= mSEndDate Then
                  mExpDays = DateDiff("d", mSStartDate, TmpRs2!ExpenseTo)
               ElseIf TmpRs2!ExpenseFrom >= mSStartDate And TmpRs2!ExpenseTo > mSEndDate Then
                  mExpDays = DateDiff("d", TmpRs2!ExpenseFrom, mSEndDate)
               Else
                  mExpDays = DateDiff("d", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo)
               End If
               mExpDays = mExpDays + 1
               mExpAmt = mTotExpense * mExpDays
            End If
            Call FixVendorExpense(TmpRs2!LocationName, mExpAmt)
            MSHFlexGrid2.TextMatrix(mRow, 0) = TmpRs2!LocationName
            MSHFlexGrid2.TextMatrix(mRow, 1) = TmpRs2!VendorName
            MSHFlexGrid2.TextMatrix(mRow, 2) = TmpRs2!ExpenseType
            MSHFlexGrid2.TextMatrix(mRow, 3) = Format(TmpRs2!ExpenseFrom, "dd-MMM-yyyy")
            MSHFlexGrid2.TextMatrix(mRow, 4) = Format(TmpRs2!ExpenseTo, "dd-MMM-yyyy")
            MSHFlexGrid2.TextMatrix(mRow, 5) = TmpRs2!TotalExpenseAmount
            MSHFlexGrid2.TextMatrix(mRow, 6) = Format(mSStartDate, "dd-MMM-yyyy")
            MSHFlexGrid2.TextMatrix(mRow, 7) = Format(mSEndDate, "dd-MMM-yyyy")
            MSHFlexGrid2.TextMatrix(mRow, 8) = Format(mExpAmt, "#####0.00")
            MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
            mRow = mRow + 1
            TmpRs2.MoveNext
        Next r
        TmpRs2.Filter = ""
    Next
End If
End Sub
'
Private Sub FillGodownRent()
Dim mRentDays, mTotRentAmt, mRentAmt, mTotRent, mTotMonths As Double
Dim mRentStartDate, mRentEndDate As Variant
Dim mSStartDate, mSEndDate As Variant
Dim mRentStart, mRentEnd As Variant
Dim mPaidRent  As Double
Dim mRow As Integer

tmp = "Select ML.LocationID,ML.LocationName,MC.CMPName,TGL.GodownID,MG.GodownNo,TGL.LessorID,LA.LessorName,LA.NAVCode,TGL.RentAmount,TLA.StartDate,TLA.ExpiryDate,MG.CloseDate,MG.StartDate as GodownStartDate,TLA.Flag"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
tmp = tmp & " Inner join Mst_CMP  MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Location  ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Inner join Mst_Lessor  LA on TGL.LessorID=LA.LessorID"
tmp = tmp & " Where TLA.StartDate <= '" & mEndDate & "' And TLA.ExpiryDate>='" & mStartDate & "' And (MG.CloseDate >= '" & mStartDate & "' or MG.CloseDate is null) And TLA.Flag<>'R'"
tmp = tmp & " Order By ML.LocationID"

Call Tmp5Table

Call Grid_Head2
If TmpRs5.RecordCount > 0 Then
   mTotRentAmt = 0
   mRow = 1
   TmpRs5.MoveFirst
   For r = 1 To TmpRs5.RecordCount
        mRentAmt = 0
        mTotMonths = DateDiff("m", mStartDate, mEndDate)
        For I = 0 To mTotMonths
            mCurrMonth = DateAdd("m", I, mStartDate)
            mYear = Year(mCurrMonth)
            mMonthNo = Month(mCurrMonth)
            mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
            mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
            mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
            mRentStartDate = IIf(TmpRs5!StartDate < TmpRs5!GodownStartDate, TmpRs5!GodownStartDate, TmpRs5!StartDate)
            
            If IsNull(TmpRs5!CloseDate) = False Then
               mRentEndDate = IIf(TmpRs5!ExpiryDate > TmpRs5!CloseDate, TmpRs5!CloseDate, TmpRs5!ExpiryDate)
            Else
               mRentEndDate = TmpRs5!ExpiryDate
            End If
            If mTotDays <> 0 Then
               
               mTotRent = TmpRs5!RentAmount / mTotDays
               
               If mRentStartDate < CDate(mSStartDate) And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = DateDiff("d", CDate(mSStartDate), CDate(mSEndDate))
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
               
               ElseIf mRentStartDate < CDate(mSStartDate) And mRentEndDate <= CDate(mSEndDate) And mRentEndDate >= CDate(mSStartDate) Then
                  mRentDays = DateDiff("d", CDate(mSStartDate), mRentEndDate)
                  mRentStart = mSStartDate
                  mRentEnd = mRentEndDate
'                    x = x
               ElseIf mRentStartDate < CDate(mSStartDate) And mRentEndDate <= CDate(mSEndDate) Then
                  mRentDays = 0
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
                  
               ElseIf mRentStartDate >= CDate(mSEndDate) And mRentEndDate >= CDate(mSStartDate) Then  ' And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = 0
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
               ElseIf mRentStartDate <= CDate(mSEndDate) And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = DateDiff("d", mRentStartDate, CDate(mSEndDate))
                  mRentStart = mRentStartDate
                  mRentEnd = mSEndDate
               Else
                  mRentDays = mTotDays 'DateDiff("d", mRentStartDate, mRentEndDate)
               End If
               If mRentDays <> 0 Then
                  mRentDays = mRentDays + 1
                  mRentAmt = mRentAmt + mTotRent * mRentDays
               End If
            End If
            mDoublePaidAmt = 0
            mPaidRent = GetPaidRent(TmpRs5!GodownID, TmpRs5!LessorID, mYear, mMonthNo)
            Call FixGodownRent(TmpRs5!LocationName, Format(mTotRent * mRentDays, "#####0.00"))
            MSHFlexGrid3.TextMatrix(mRow, 0) = TmpRs5!LocationID
            MSHFlexGrid3.TextMatrix(mRow, 1) = TmpRs5!LocationName
            MSHFlexGrid3.TextMatrix(mRow, 2) = TmpRs5!CMPName
            MSHFlexGrid3.TextMatrix(mRow, 3) = TmpRs5!GodownNo
            MSHFlexGrid3.TextMatrix(mRow, 4) = TmpRs5!LessorName
            MSHFlexGrid3.TextMatrix(mRow, 5) = IIf(IsNull(TmpRs5!NAVCode), "", TmpRs5!NAVCode)
            MSHFlexGrid3.TextMatrix(mRow, 6) = Format(mRentStartDate, "dd-MMM-yyyy")
            MSHFlexGrid3.TextMatrix(mRow, 7) = Format(mRentEndDate, "dd-MMM-yyyy")
            MSHFlexGrid3.TextMatrix(mRow, 8) = TmpRs5!RentAmount
            MSHFlexGrid3.TextMatrix(mRow, 9) = Format(mRentStart, "dd-MMM-yyyy")
            MSHFlexGrid3.TextMatrix(mRow, 10) = Format(mRentEnd, "dd-MMM-yyyy")
            MSHFlexGrid3.TextMatrix(mRow, 11) = mRentDays
            MSHFlexGrid3.TextMatrix(mRow, 12) = Format(mTotRent * mRentDays, "#####0.00")
            MSHFlexGrid3.TextMatrix(mRow, 13) = Format(mPaidRent, "#####0.00")
            MSHFlexGrid3.TextMatrix(mRow, 14) = Format(mDoublePaidAmt, "#####0.00")
            MSHFlexGrid3.TextMatrix(mRow, 15) = TmpRs5!Flag
            
            mRow = mRow + 1
            MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
            
        Next
        TmpRs5.MoveNext
    Next
End If
End Sub
Private Function GetPaidRent(mgodownID_ As Double, mLessorID_ As Double, myear_ As Variant, mMonth_ As Variant) As Double
Dim Retval As Variant
Retval = 0
tmp = " Select TG.GodownID,TGr.LessorID,TLA.StartDate,TGr.ExpenseFrom,TGr.ExpenseTo,TG.Amount as RentAmount from  Txn_GodownRentDetail TG"
tmp = tmp & " Inner join Txn_GodownRent TGr on TG.ClaimID=TGr.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TG.GodownID"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TG.LAID=TLA.LAID"
tmp = tmp & " Where  TGr.Flag <> 'R' And TG.GodownID=" & mgodownID_ & " And TGr.LessorID=" & mLessorID_ & " And Month(ExpenseFrom)=" & mMonth_ & " And year(ExpenseFrom)=" & myear_

Call Tmp2Table
For I = 1 To TmpRs2.RecordCount
    If I = 1 Then
       Retval = TmpRs2!RentAmount
    Else
       mDoublePaidAmt = mDoublePaidAmt + TmpRs2!RentAmount
    End If
Next
GetPaidRent = Retval
End Function
Private Sub dispdata()

'--For Location Wise DP Income

tmp = "Select MDP.LocationID,ML.LocationName,sum(TotalAmt)As IncomeAmt from Mst_DPInvoiceUpload MDP"
tmp = tmp & " inner Join Mst_location ML on MDp.LocationID=ML.LocationID"
tmp = tmp & " Where MDP.MonthYear>='" & mStartDate & "' And MDP.MonthYear<='" & mEndDate & "'"
tmp = tmp & " Group By MDP.LocationID,ML.LocationName"

Call Tmp3Table


'--For Location wise NN-NC Income

tmp = "Select ML.LocationID,ML.LocationName,Sum(TIG.InvoiceAmount) IncomeAmt from  Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_CMP MC on TIG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join MSt_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And TIG.InvoiceDate >='" & mStartDate & "' And TIG.InvoiceDate <='" & mEndDate & "'"
tmp = tmp & " Group By ML.LocationID,ML.LocationName"


''tmp = "Select ML.LocationID,ML.LocationName,Sum(TIG.InvoiceAmount) IncomeAmt from  Txn_InvMonGSLDetail TIG"
''tmp = tmp & " Inner Join Mst_CMP MC on TIG.CMPID=MC.CMPID"
''tmp = tmp & " Inner Join Mst_Godown MG on TIG.GodownID=MG.GodownID"
''tmp = tmp & " Inner Join MSt_Location ML on MC.LocationID=ML.LocationID"
''tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null) And MG.UnitTypeID in (1,2)"
''tmp = tmp & " And TIG.InvoiceDate >='" & mStartDate & "' And TIG.InvoiceDate <='" & mEndDate & "'"
''tmp = tmp & " Group By ML.LocationID,ML.LocationName"


Call Tmp6Table

'--For Location wise Lease Capacity

'tmp = "Select ML.LocationName As Location,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML"
'tmp = tmp & " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID"
'tmp = tmp & " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID"
'tmp = tmp & " Where MG.UnitTypeID in (1,2,3) And (MG.CloseDate is Null or MG.CloseDate <='" & mEndDate & "') And MG.StartDate <='" & mEndDate & "'"
'tmp = tmp & " Group By ML.LocationName"

'-- After discus with Sankar Sir, Added on 19-07-2010

tmp = "Select ML.LocationName As Location,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML"
tmp = tmp & " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID"
tmp = tmp & " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID"
tmp = tmp & " Where MG.UnitTypeID =1 And (MG.CloseDate is Null or MG.CloseDate >'" & mEndDate & "') And MG.StartDate <='" & mEndDate & "'"
tmp = tmp & " Group By ML.LocationName"

'tmp = "Select ML.LocationName As Location,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML"
'tmp = tmp & " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID"
'tmp = tmp & " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID"
'tmp = tmp & " Where MG.UnitTypeID =1 And (MG.CloseDate is Null or MG.CloseDate >'" & mStartDate & "') And MG.StartDate <='" & mEndDate & "'"
'tmp = tmp & " Group By ML.LocationName"


Call Tmp1Table

'--For Location wise Franchisee Stock
tmp = "Select ML.LocationName,Sum(GSL.BalanceWeight) as Stock from Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG on GSL.GodownID=MG.GodownID"
tmp = tmp & " Inner join Mst_CMP  MC on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner join Mst_Location  ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where (GSL.Flag <> 'E') AND (GSL.Flag <> 'Z') And (MG.CloseDate is Null or MG.CloseDate >'" & mEndDate & "') And MG.StartDate <='" & mEndDate & "' And MG.UnitTypeID=2"
tmp = tmp & " Group by ML.LocationName"

Call Tmp7Table

For r = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(r, 1) <> "" Then
    '--For Lease Capacity
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
       End If
       TmpRs1.Find "Location='" & MSHFlexGrid1.TextMatrix(r, 1) & "'"
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(r, 2) = IIf(IsNull(TmpRs1!TotalCapacity), 0, TmpRs1!TotalCapacity)
       End If
       
    '--For Franchisee Stock
       If TmpRs7.RecordCount > 0 Then
          TmpRs7.MoveFirst
       End If
       TmpRs7.Find "LocationName='" & MSHFlexGrid1.TextMatrix(r, 1) & "'"
       If Not TmpRs7.EOF Then
          MSHFlexGrid1.TextMatrix(r, 2) = Val(MSHFlexGrid1.TextMatrix(r, 2)) + IIf(IsNull(TmpRs7!Stock), 0, TmpRs7!Stock)
       End If
       
     '--For DP Income
       If TmpRs3.RecordCount > 0 Then
          TmpRs3.MoveFirst
       End If
       TmpRs3.Find "LocationName='" & MSHFlexGrid1.TextMatrix(r, 1) & "'"
       If Not TmpRs3.EOF Then
          MSHFlexGrid1.TextMatrix(r, 3) = Format(IIf(IsNull(TmpRs3!IncomeAmt), 0, TmpRs3!IncomeAmt), "######0.00")
       End If
       
     '--For NN-NC Income
       If TmpRs6.RecordCount > 0 Then
          TmpRs6.MoveFirst
       End If
       TmpRs6.Find "LocationName='" & MSHFlexGrid1.TextMatrix(r, 1) & "'"
       If Not TmpRs6.EOF Then
          MSHFlexGrid1.TextMatrix(r, 3) = Val(MSHFlexGrid1.TextMatrix(r, 3)) + Format(IIf(IsNull(TmpRs6!IncomeAmt), 0, TmpRs6!IncomeAmt), "######0.00")
       End If
       
    End If
Next
End Sub
'Private Sub FillNotAgreeGodowns()
'
'tmp = "Select ML.LocationName,MG.GodownID,MG.GodownNo,MG.Address From Mst_Godown MG"
'tmp = tmp & " inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " Where MG.UnitTypeID=1 And (MG.CloseDate is Null or MG.CloseDate <='" & mEndDate & "') And MG.StartDate <='" & mEndDate & "' And MG.GodownID not in (Select distinct TGL.GodownID"
'tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
'tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
'tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
'tmp = tmp & " Where TLA.StartDate <= '" & mEndDate & "' And TLA.ExpiryDate>='" & mStartDate & "' And (MG.CloseDate >= '" & mStartDate & "' or MG.CloseDate is null) And TLA.Flag<>'R')"
'
'Call Tmp2Table
'Call Grid_Head6
'If TmpRs2.RecordCount > 0 Then
'   For I = 1 To TmpRs2.RecordCount
'       r = MSHFlexGrid7.Rows - 1
'       MSHFlexGrid7.TextMatrix(r, 0) = TmpRs2!LocationName
'       MSHFlexGrid7.TextMatrix(r, 1) = TmpRs2!GodownID
'       MSHFlexGrid7.TextMatrix(r, 2) = TmpRs2!GodownNo
'       MSHFlexGrid7.TextMatrix(r, 3) = TmpRs2!Address
'
'       MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + 1
'       TmpRs2.MoveNext
'   Next
'End If
'End Sub
Private Sub FillCCBDOexpense(mLocationIDs_ As String, mRow_ As Double)
Dim mFlag As Boolean
Dim mMonths, mExpAmt, mTotMonths As Variant

tmp = "Select TCS.EmployeeID,ME.EmployeeName + '~'+ ME.EmployeeNo As EmpName , ME.DesignationID,Sum(TCS.Amount) As TotalExpenseAmount,TCS.ExpenseFrom,TCS.ExpenseTo"
tmp = tmp & " From Txn_ClaimSubmission TCS"
tmp = tmp & " Inner Join Mst_Employee ME On TCS.EmployeeID = ME.EmployeeID"
tmp = tmp & " Where TCS.ExpenseFrom <= '" & mEndDate & "' And TCS.ExpenseTo >= '" & mStartDate & "' And TCS.Flag <> 'R' And ME.DesignationID in (4,9,3)"
tmp = tmp & " And TCS.LocationID in ( " & mLocationIDs_ & " )"
tmp = tmp & " Group By TCS.EmployeeID,ME.EmployeeName, ME.EmployeeNo , ME.DesignationID,TCS.ExpenseFrom,TCS.ExpenseTo Order by ME.DesignationID"

Call Tmp2Table
mFlag = False
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       mMonths = DateDiff("m", mStartDate, mEndDate) + 1

        mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo) + 1
       mExpAmt = IIf(IsNull(TmpRs2!TotalExpenseAmount), 0, TmpRs2!TotalExpenseAmount) / mTotMonths
       
       If TmpRs2!ExpenseFrom < mStartDate Then
          mTotMonths = DateDiff("m", mStartDate, TmpRs2!ExpenseTo) + 1
       ElseIf TmpRs2!ExpenseTo > mEndDate Then
           mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, mEndDate) + 1
       Else
           mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo) + 1
       End If
       If mTotMonths > mMonths Then
          mTotExpAmt = mExpAmt * mMonths
       Else
          mTotExpAmt = mExpAmt * mTotMonths
       End If
       For r = mRow_ To MSHFlexGrid1.Rows - 1
           If TmpRs2!DesignationID = 3 Then
              If MSHFlexGrid1.TextMatrix(r, 1) = "AM(Current)-" & TmpRs2!EmpName Or MSHFlexGrid1.TextMatrix(r, 1) = "AM-" & TmpRs2!EmpName Then
                 MSHFlexGrid1.TextMatrix(r, 4) = Val(MSHFlexGrid1.TextMatrix(r, 4)) + Format(mTotExpAmt, "######0.00")
                 Exit For
              End If
           ElseIf TmpRs2!DesignationID = 4 Then
              If MSHFlexGrid1.TextMatrix(r, 1) = "CC-" & TmpRs2!EmpName Then
                 MSHFlexGrid1.TextMatrix(r, 4) = Val(MSHFlexGrid1.TextMatrix(r, 4)) + Format(mTotExpAmt, "######0.00")
                 Exit For
              End If
           ElseIf TmpRs2!DesignationID = 9 Then
              If MSHFlexGrid1.TextMatrix(r, 1) = "BDO-" & TmpRs2!EmpName Then
                 MSHFlexGrid1.TextMatrix(r, 4) = Val(MSHFlexGrid1.TextMatrix(r, 4)) + Format(mTotExpAmt, "######0.00")
                 Exit For
              End If
           End If
       Next
       TmpRs2.MoveNext
   Next
End If

End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 17
MSHFlexGrid1.FixedCols = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.TextMatrix(0, 0) = "RM"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Capacity"
MSHFlexGrid1.TextMatrix(0, 3) = "Income(NNNR+DP)"
MSHFlexGrid1.TextMatrix(0, 4) = "Employee Expense"
MSHFlexGrid1.TextMatrix(0, 5) = "Vendor Expense "
MSHFlexGrid1.TextMatrix(0, 6) = "Franchisee Fees "
MSHFlexGrid1.TextMatrix(0, 7) = "Godown Rent "
MSHFlexGrid1.TextMatrix(0, 8) = "Employee Salary"
MSHFlexGrid1.TextMatrix(0, 9) = "Surplus(from Adjustment Sales option)"
MSHFlexGrid1.TextMatrix(0, 10) = "Loss(from adjustment Theft/Fire/burglary)"
MSHFlexGrid1.TextMatrix(0, 11) = "AM/RM Allocation"
MSHFlexGrid1.TextMatrix(0, 12) = "CMG Allocation"
MSHFlexGrid1.TextMatrix(0, 13) = "Total Expense"
MSHFlexGrid1.TextMatrix(0, 14) = "Gross Margin"
MSHFlexGrid1.TextMatrix(0, 15) = "CO Allocation"
MSHFlexGrid1.TextMatrix(0, 16) = "Net Margin"

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200

MSHFlexGrid1.RowHeight(0) = 600

End Sub
Private Sub TotalSalaryLocWise()

tmp = " Select ME.EmployeeID,ME.EmployeeName + '~' + TC.EmployeeCode as EmployeeName,MD.Designation,ME.AllocateExpense,TC.CTC_Yearly " & _
      " From Txn_EmpCTC TC " & _
      " Inner Join Mst_Employee ME On TC.EmployeeCode = ME.EmployeeNo " & _
      " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID " & _
      " Where EmployeeCode In (Select EmployeeNo From Mst_Employee ) " & _
      " Order By EmployeeCode"

Call Tmp1Table

mTotOtherEmpSal = 0
Call Grid_Head5
mRow = 1
For i1 = 1 To TmpRs1.RecordCount
    Call CalDirectIncomeLocWise(8, TmpRs1!EmployeeID, TmpRs1!Designation, TmpRs1!CTC_YEARLY)
    TmpRs1.MoveNext
Next

mRow = MSHFlexGrid1.Rows - 1
MSHFlexGrid1.TextMatrix(mRow, 1) = "Other Employee Salary"
MSHFlexGrid1.TextMatrix(mRow, 8) = mTotOtherEmpSal
MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
tmp = mTotOtherEmpID
End Sub
Private Sub CalDirectIncomeLocWise(mRow_ As Variant, mEmpID_ As Variant, mEMPDesignation_ As Variant, mExpenseAmount_ As Variant)
Dim mAmtAfterDeduction, mAmtAfterDeductionI, mCorporateExpense, mCorporateExpenseAfterDed, mCorporateExpenseAfterDedI As Variant
Dim mEMPIDDummy_ As Double
Dim mTotCap As Double
Dim mTotIncome As Double


mEMPIDDummy_ = mEmpID_

If LCase(mEMPDesignation_) = "managment" Or LCase(mEMPDesignation_) = "rag" Then
   Exit Sub
Else
   If LCase(mEMPDesignation_) = "qqs" Or LCase(mEMPDesignation_) = "drc" Or LCase(mEMPDesignation_) = "bdo" Then
      tmp = "Select DirectReportingTo From Mst_Employee Where EmployeeID = " & mEMPIDDummy_ & ""
      Call Tmp2Table
      If TmpRs2.RecordCount <> 0 Then
         mEMPIDDummy_ = IIf(IsNull(TmpRs2!DirectReportingTo), 0, TmpRs2!DirectReportingTo)
      End If
   End If
   tmp = "Select distinct ECM.LocationID,ML.LocationName " & _
         " From Mst_EMPCMPMapping ECM " & _
         " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID " & _
         " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID " & _
         " Left Outer Join Mst_Employee EMP6 On EMP.ConnectedToRAG = EMP6.EmployeeID " & _
         " Left Outer Join Mst_Location ML On ECM.LocationID = ML.LocationID " & _
         " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID " & _
         " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID " & _
         " Where (EMP.EmployeeID = " & mEMPIDDummy_ & " Or EMP1.EmployeeID = " & mEMPIDDummy_ & " Or EMP2.EmployeeID = " & mEMPIDDummy_ & " Or EMP3.EmployeeID = " & mEMPIDDummy_ & " Or EMP4.EmployeeID = " & mEMPIDDummy_ & " Or EMP6.EmployeeID = " & mEMPIDDummy_ & ") And " & _
         " MG.CloseDate is Null " & _
         " Group By ECM.LocationID,ML.LocationName " & _
         " Order By ECM.LocationID "
End If

Call Tmp2Table

mTotExpense = mExpenseAmount_ / 12
''SSTab1.Tab = 0
mExpDays = DateDiff("m", mStartDate, mEndDate)
mExpDays = mExpDays + 1
mExpenseAmount_ = Format(mTotExpense * mExpDays, "#####0.00")
   
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       For ix = 1 To MSHFlexGrid1.Rows - 1
           If MSHFlexGrid1.TextMatrix(ix, 1) = TmpRs2!LocationName Then
               If TmpRs.RecordCount > 0 Then
                  TmpRs.MoveFirst
               End If
               TmpRs.Filter = ""
               TmpRs.Filter = "LocationName='" & TmpRs2!LocationName & "'"
               If TmpRs.RecordCount > 0 Then
                  If FixAMSalary("AM(Current)-" & TmpRs!CCAM, mExpenseAmount_) = False Then
                     mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
                  End If
                  For r = 0 To mExpDays - 1
                      mCurrMonth = DateAdd("m", r, mStartDate)
                      mYear = Year(mCurrMonth)
                      mMonthNo = Month(mCurrMonth)
                      mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
                      mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
                      mRow = MSHFlexGrid6.Rows - 1
                      MSHFlexGrid6.TextMatrix(mRow, 0) = Format(mSStartDate, "dd-MMM-yyyy")
                      MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs1!EmployeeName
                      MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs1!CTC_YEARLY
                      MSHFlexGrid6.TextMatrix(mRow, 3) = mTotExpense
                      MSHFlexGrid6.TextMatrix(mRow, 4) = TmpRs!CCAM
                        
                      MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
                  Next r
                  Exit Sub
               Else
                  mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
               End If
           End If
       Next
       TmpRs2.MoveNext
   Next

Else
  tmp = "Select LocationName From Mst_Employee Inner Join Mst_Location on Mst_Employee.BaseLocationID=Mst_Location.LocationID Where EmployeeID=" & mEmpID_

  Call Tmp2Table

  If TmpRs2.RecordCount > 0 Then
     For I = 1 To TmpRs2.RecordCount
         For ix = 1 To MSHFlexGrid1.Rows - 1
             If MSHFlexGrid1.TextMatrix(ix, 1) = TmpRs2!LocationName Then
                If TmpRs.RecordCount > 0 Then
                   TmpRs.MoveFirst
                End If
                TmpRs.Filter = ""
                TmpRs.Filter = "LocationName='" & TmpRs2!LocationName & "'"
                If TmpRs.RecordCount > 0 Then
                   If FixAMSalary("AM(Current)-" & TmpRs!CCAM, mExpenseAmount_) = False Then
                      mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
                   End If

                   For r = 0 To mExpDays - 1
                       mCurrMonth = DateAdd("m", r, mStartDate)
                       mYear = Year(mCurrMonth)
                       mMonthNo = Month(mCurrMonth)
                       mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
                       mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
                       mRow = MSHFlexGrid6.Rows - 1
                       MSHFlexGrid6.TextMatrix(mRow, 0) = Format(mSStartDate, "dd-MMM-yyyy")
                       MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs1!EmployeeName
                       MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs1!CTC_YEARLY
                       MSHFlexGrid6.TextMatrix(mRow, 3) = mTotExpense
                       MSHFlexGrid6.TextMatrix(mRow, 4) = IIf(IsNull(TmpRs!CCAM), "", TmpRs!CCAM)

                       MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
                   Next r
                   Exit Sub
                Else
                   mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
                End If
             End If
         Next
         TmpRs2.MoveNext
     Next
  Else
    For r = 0 To mExpDays - 1
        mCurrMonth = DateAdd("m", r, mStartDate)
        mYear = Year(mCurrMonth)
        mMonthNo = Month(mCurrMonth)
        mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
        mSStartDate = mYear & "-" & mMonthNo & "-" & "01"

        mRow = MSHFlexGrid6.Rows - 1
        MSHFlexGrid6.TextMatrix(mRow, 0) = Format(mSStartDate, "dd-MMM-yyyy")
        MSHFlexGrid6.TextMatrix(mRow, 1) = TmpRs1!EmployeeName
        MSHFlexGrid6.TextMatrix(mRow, 2) = TmpRs1!CTC_YEARLY
        MSHFlexGrid6.TextMatrix(mRow, 3) = mTotExpense
        MSHFlexGrid6.TextMatrix(mRow, 4) = "" 'IIf(IsNull(TmpRs!CCAM), "", TmpRs!CCAM)
        
        MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
    Next
    mTotOtherEmpSal = mTotOtherEmpSal + mExpenseAmount_
  End If
  
  If mTotOtherEmpID = "" Then
     mTotOtherEmpID = mEmpID_
  Else
     mTotOtherEmpID = mTotOtherEmpID & "," & mEmpID_
  End If
End If

End Sub
Private Function FixAMSalary(mAm_ As Variant, mSalary_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False
For ix = 1 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 1)) = LCase(mAm_) Then
       MSHFlexGrid1.TextMatrix(ix, 8) = Val(MSHFlexGrid1.TextMatrix(ix, 8)) + Val(mSalary_)
       Retval = True
    End If
Next
FixAMSalary = Retval
End Function
Private Sub FixVendorExpense(mLoc_ As Variant, mexp_ As Variant)
For ix = 1 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 1)) = LCase(mLoc_) Then
       MSHFlexGrid1.TextMatrix(ix, 5) = Val(MSHFlexGrid1.TextMatrix(ix, 5)) + Val(mexp_)
       Exit For
    End If
Next
End Sub

Private Sub FixGodownRent(mLoc_ As Variant, mexp_ As Variant)
For ix = 1 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(ix, 1)) = LCase(mLoc_) Then
       MSHFlexGrid1.TextMatrix(ix, 7) = Val(MSHFlexGrid1.TextMatrix(ix, 7)) + Val(mexp_)
       Exit For
    End If
Next
End Sub
Private Sub Form_Load()
'Dim FrmLoadAccess() As Boolean
'FrmLoadAccess = GetFrmLoadAccess(Me.Name)
'If FrmLoadAccess(0) = False Then
'   Frame1.Enabled = False
'   Frame2.Enabled = False
'   MsgBox "Access Denied !!!!!!!!!"
'   Exit Sub
'End If
Frame1.Enabled = True
Frame2.Enabled = True
Call CmdClear_Click
End Sub

Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 9
MSHFlexGrid2.FixedCols = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.TextMatrix(0, 0) = "Location"
MSHFlexGrid2.TextMatrix(0, 1) = "Vendor Name"
MSHFlexGrid2.TextMatrix(0, 2) = "Expense Type"
MSHFlexGrid2.TextMatrix(0, 3) = "Expense From"
MSHFlexGrid2.TextMatrix(0, 4) = "Expense To"
MSHFlexGrid2.TextMatrix(0, 5) = "Expense Amount"
MSHFlexGrid2.TextMatrix(0, 6) = "From Date"
MSHFlexGrid2.TextMatrix(0, 7) = "To Date"
MSHFlexGrid2.TextMatrix(0, 8) = "Expense Amount"

MSHFlexGrid2.ColWidth(0) = 0
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 1400
MSHFlexGrid2.ColWidth(3) = 1400
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 1400
MSHFlexGrid2.ColWidth(6) = 1400
MSHFlexGrid2.ColWidth(7) = 1400
MSHFlexGrid2.ColWidth(8) = 1200
MSHFlexGrid2.RowHeight(0) = 600

End Sub

Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 16
MSHFlexGrid3.FixedCols = 2
MSHFlexGrid3.FixedRows = 1
MSHFlexGrid3.TextMatrix(0, 0) = "LocationID"
MSHFlexGrid3.TextMatrix(0, 1) = "Location"
MSHFlexGrid3.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid3.TextMatrix(0, 3) = "Godown No"
MSHFlexGrid3.TextMatrix(0, 4) = "Lessor Name"
MSHFlexGrid3.TextMatrix(0, 5) = "Lessor NavCode"
MSHFlexGrid3.TextMatrix(0, 6) = "Rent Start Date"
MSHFlexGrid3.TextMatrix(0, 7) = "Rent End Date"
MSHFlexGrid3.TextMatrix(0, 8) = "Rent Amount"
MSHFlexGrid3.TextMatrix(0, 9) = "Rent From Date"
MSHFlexGrid3.TextMatrix(0, 10) = "Rent To Date"
MSHFlexGrid3.TextMatrix(0, 11) = "Rent Days"
MSHFlexGrid3.TextMatrix(0, 12) = "Rent Amount"
MSHFlexGrid3.TextMatrix(0, 13) = "Paid Rent Amount"
MSHFlexGrid3.TextMatrix(0, 14) = "Duplicate Rent Amount"
MSHFlexGrid3.TextMatrix(0, 15) = "Flag"

MSHFlexGrid3.ColWidth(0) = 0
MSHFlexGrid3.ColWidth(1) = 2000
MSHFlexGrid3.ColWidth(2) = 1400
MSHFlexGrid3.ColWidth(3) = 1400
MSHFlexGrid3.ColWidth(4) = 1400
MSHFlexGrid3.ColWidth(5) = 1400
MSHFlexGrid3.ColWidth(6) = 1400
MSHFlexGrid3.ColWidth(7) = 1400
MSHFlexGrid3.ColWidth(8) = 1200
MSHFlexGrid3.RowHeight(0) = 600

End Sub

Private Sub Grid_Head3()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 2
MSHFlexGrid4.Cols = 5
MSHFlexGrid4.FixedCols = 2
MSHFlexGrid4.FixedRows = 1
MSHFlexGrid4.TextMatrix(0, 0) = "LocationID"
MSHFlexGrid4.TextMatrix(0, 1) = "Location"
MSHFlexGrid4.TextMatrix(0, 2) = "Unit Type"
MSHFlexGrid4.TextMatrix(0, 3) = "Invoice Date"
MSHFlexGrid4.TextMatrix(0, 4) = "Invoice Amount"

MSHFlexGrid4.ColWidth(0) = 0
MSHFlexGrid4.ColWidth(1) = 2000
MSHFlexGrid4.ColWidth(2) = 1400
MSHFlexGrid4.ColWidth(3) = 1400
MSHFlexGrid4.ColWidth(4) = 1400
MSHFlexGrid4.RowHeight(0) = 600

End Sub
Private Sub Grid_Head7()

MSHFlexGrid8.Clear
MSHFlexGrid8.Rows = 2
MSHFlexGrid8.Cols = 5
MSHFlexGrid8.FixedCols = 2
MSHFlexGrid8.FixedRows = 1
MSHFlexGrid8.TextMatrix(0, 0) = "LocationID"
MSHFlexGrid8.TextMatrix(0, 1) = "Location"
MSHFlexGrid8.TextMatrix(0, 2) = "Unit Type"
MSHFlexGrid8.TextMatrix(0, 3) = "Invoice Date"
MSHFlexGrid8.TextMatrix(0, 4) = "Invoice Amount"
MSHFlexGrid8.ColWidth(0) = 0
MSHFlexGrid8.ColWidth(1) = 2000
MSHFlexGrid8.ColWidth(2) = 1400
MSHFlexGrid8.ColWidth(3) = 1400
MSHFlexGrid8.ColWidth(4) = 1400
MSHFlexGrid8.RowHeight(0) = 600

End Sub
Private Sub FillNNNCIncome()

tmp = "Select ML.LocationID,ML.LocationName,MU.UnitType,TIG.InvoiceDate,Sum(TIG.InvoiceAmount)InvoiceAmount"
tmp = tmp & " from  Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on TIG.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
tmp = tmp & " Inner Join Mst_CMP MC on TIG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join MSt_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TIG.ExchangeTypeID=1 And (TIG.ContractID=0 Or TIG.ContractID is Null)"  ' And MG.UnitTypeID in (1,2)"
tmp = tmp & " And TIG.InvoiceDate >='" & mStartDate & "' And TIG.InvoiceDate <='" & mEndDate & "'"
tmp = tmp & " Group By ML.LocationID,ML.LocationName,MU.UnitType,TIG.InvoiceDate"
tmp = tmp & " Order By ML.LocationName,MU.UnitType,TIG.InvoiceDate"

Call Grid_Head3
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid4.Rows = TmpRs2.RecordCount + 2
   For r = 1 To TmpRs2.RecordCount
       MSHFlexGrid4.TextMatrix(r, 0) = TmpRs2!LocationID
       MSHFlexGrid4.TextMatrix(r, 1) = TmpRs2!LocationName
       MSHFlexGrid4.TextMatrix(r, 2) = TmpRs2!UnitType
       MSHFlexGrid4.TextMatrix(r, 3) = Format(TmpRs2!InvoiceDate, "dd-MMM-yyyy")
       MSHFlexGrid4.TextMatrix(r, 4) = Format(TmpRs2!InvoiceAmount, "#####0.00")
       TmpRs2.MoveNext
   Next
End If
End Sub

Private Sub FillDPIncome()

tmp = "Select MDP.LocationID,ML.LocationName,MDP.MonthYear,sum(TotalAmt)As IncomeAmt from Mst_DPInvoiceUpload MDP"
tmp = tmp & " inner Join Mst_location ML on MDp.LocationID=ML.LocationID"
tmp = tmp & " Where MDP.MonthYear>='" & mStartDate & "' And MDP.MonthYear<='" & mEndDate & "'"
tmp = tmp & " Group By MDP.LocationID,ML.LocationName,MDP.MonthYear"
tmp = tmp & " Order By MDP.LocationID,ML.LocationName,MDP.MonthYear"

Call Grid_Head7
Call Tmp2Table
If TmpRs2.RecordCount > 0 Then
   MSHFlexGrid8.Rows = TmpRs2.RecordCount + 2
   For r = 1 To TmpRs2.RecordCount
       MSHFlexGrid8.TextMatrix(r, 0) = TmpRs2!LocationID
       MSHFlexGrid8.TextMatrix(r, 1) = TmpRs2!LocationName
       MSHFlexGrid8.TextMatrix(r, 2) = "" 'TmpRs2!UnitType
       MSHFlexGrid8.TextMatrix(r, 3) = Format(TmpRs2!MonthYear, "dd-MMM-yyyy")
       MSHFlexGrid8.TextMatrix(r, 4) = Format(TmpRs2!IncomeAmt, "#####0.00")
       TmpRs2.MoveNext
   Next
End If
End Sub

Private Sub FillEmployeeExpense()
Dim mRentDays, mTotExpAmt, mExpAmt As Double
Dim mRentStartDate, mRentEndDate As Variant
Dim mSStartDate, mSEndDate As Variant
Dim mRentStart, mRentEnd As Variant
Dim mPaidRent  As Double
Dim mRow As Integer
Dim mTotMonths As Double

tmp = "Select TC.ClaimRowID,ME.EmployeeID, ME.EmployeeName + '~' + ME.EmployeeNo  as Employee,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType, Sum(TCS.Amount) As TotalExpenseAmount,TC.Remarks,Me.AllocateExpense"
tmp = tmp & " From Txn_ClaimSubmissionDetail TCS"
tmp = tmp & " inner Join Txn_ClaimSubmission TC on TCS.ClaimRowID=TC.ClaimRowID"
tmp = tmp & " inner Join Mst_ExpenseType MEx on TCS.ExpenseTypeID=MEx.ExpenseTypeID"
tmp = tmp & " Inner Join Mst_Employee ME On TC.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
tmp = tmp & " Where TC.ExpenseFrom <= '" & mEndDate & "' And TC.ExpenseTo >= '" & mStartDate & "' And TC.Flag <> 'R' And ME.DesignationID in (5,6,8,4,9,3)"
tmp = tmp & " Group By TC.ClaimRowID,ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,ME.AllocateExpense,MD.Designation,TC.ExpenseFrom,TC.ExpenseTo,MEx.ExpenseType ,TC.Remarks"

Call Tmp5Table

Call Grid_Head4
If TmpRs5.RecordCount > 0 Then
   
    mRow = 1
    mTotMonths = DateDiff("m", mStartDate, mEndDate)
    For I = 0 To mTotMonths
        mCurrMonth = DateAdd("m", I, mStartDate)
        mYear = Year(mCurrMonth)
        mMonthNo = Month(mCurrMonth)
        mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
        mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
        mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
        
        TmpRs5.Filter = "ExpenseFrom <= '" & mSEndDate & "' And ExpenseTo >= '" & mSStartDate & "'"
        
        For r = 1 To TmpRs5.RecordCount
            mTotMonths = DateDiff("m", TmpRs5!ExpenseFrom, TmpRs5!ExpenseTo) + 1
            mExpAmt = TmpRs5!TotalExpenseAmount / mTotMonths
            MSHFlexGrid5.TextMatrix(mRow, 0) = TmpRs5!ClaimRowID
            MSHFlexGrid5.TextMatrix(mRow, 1) = TmpRs5!Employee
            MSHFlexGrid5.TextMatrix(mRow, 2) = TmpRs5!Designation
            MSHFlexGrid5.TextMatrix(mRow, 3) = Format(TmpRs5!ExpenseFrom, "dd-MMM-yyyy")
            MSHFlexGrid5.TextMatrix(mRow, 4) = Format(TmpRs5!ExpenseTo, "dd-MMM-yyyy")
            MSHFlexGrid5.TextMatrix(mRow, 5) = TmpRs5!ExpenseType
            MSHFlexGrid5.TextMatrix(mRow, 6) = TmpRs5!TotalExpenseAmount
            MSHFlexGrid5.TextMatrix(mRow, 7) = Format(mSStartDate, "dd-MMM-yyyy")
            MSHFlexGrid5.TextMatrix(mRow, 8) = Format(mExpAmt, "#####0.00")
            
            mRow = mRow + 1
            MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
            TmpRs5.MoveNext
        Next r
        TmpRs5.Filter = ""
    Next

End If
End Sub

Private Sub Grid_Head4()

MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 2
MSHFlexGrid5.Cols = 9
MSHFlexGrid5.FixedCols = 2
MSHFlexGrid5.FixedRows = 1
MSHFlexGrid5.TextMatrix(0, 0) = "ClaimID"
MSHFlexGrid5.TextMatrix(0, 1) = "EmployeeName"
MSHFlexGrid5.TextMatrix(0, 2) = "Designation"
MSHFlexGrid5.TextMatrix(0, 3) = "Expesne From"
MSHFlexGrid5.TextMatrix(0, 4) = "Expense To"
MSHFlexGrid5.TextMatrix(0, 5) = "Expense Type"
MSHFlexGrid5.TextMatrix(0, 6) = "Total Expense Amount"
MSHFlexGrid5.TextMatrix(0, 7) = "For Month"
MSHFlexGrid5.TextMatrix(0, 8) = "Expense Amount"

MSHFlexGrid5.ColWidth(0) = 0
MSHFlexGrid5.ColWidth(1) = 2000
MSHFlexGrid5.ColWidth(2) = 1400
MSHFlexGrid5.ColWidth(3) = 1400
MSHFlexGrid5.ColWidth(4) = 1400
MSHFlexGrid5.ColWidth(5) = 1400
MSHFlexGrid5.ColWidth(6) = 1400
MSHFlexGrid5.ColWidth(7) = 1400
MSHFlexGrid5.ColWidth(8) = 1200
MSHFlexGrid5.RowHeight(0) = 600

End Sub

Private Sub Grid_Head5()

MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 2
MSHFlexGrid6.Cols = 5
MSHFlexGrid6.FixedCols = 2
MSHFlexGrid6.FixedRows = 1
MSHFlexGrid6.TextMatrix(0, 0) = "Month"
MSHFlexGrid6.TextMatrix(0, 1) = "Employee Name"
MSHFlexGrid6.TextMatrix(0, 2) = "CTC Yearly"
MSHFlexGrid6.TextMatrix(0, 3) = "Monthly"
MSHFlexGrid6.TextMatrix(0, 4) = "AM"


MSHFlexGrid6.ColWidth(0) = 0
MSHFlexGrid6.ColWidth(1) = 2000
MSHFlexGrid6.ColWidth(2) = 1400
MSHFlexGrid6.ColWidth(3) = 1400
MSHFlexGrid6.ColWidth(4) = 1400
MSHFlexGrid6.RowHeight(0) = 600

End Sub

'Private Sub Grid_Head6()
'
'MSHFlexGrid7.Clear
'MSHFlexGrid7.Rows = 2
'MSHFlexGrid7.Cols = 4
'MSHFlexGrid7.FixedCols = 1
'MSHFlexGrid7.FixedRows = 1
'MSHFlexGrid7.TextMatrix(0, 0) = "Location Name"
'MSHFlexGrid7.TextMatrix(0, 1) = "GodownID"
'MSHFlexGrid7.TextMatrix(0, 2) = "Godown No"
'MSHFlexGrid7.TextMatrix(0, 3) = "Address"
'
'MSHFlexGrid7.ColWidth(0) = 2200
'MSHFlexGrid7.ColWidth(1) = 1000
'MSHFlexGrid7.ColWidth(2) = 1400
'MSHFlexGrid7.ColWidth(3) = 3000
'
'MSHFlexGrid7.RowHeight(0) = 600
'
'End Sub
Private Sub FillCCBDO(mLocationIDs_ As String, Optional mDesgn_ As String, Optional mAmID_ As Double, Optional mAm_ As String, Optional mRM_ As Variant)
Dim mFlag As Boolean
Dim mMonths, mExpAmt, mTotMonths As Variant

tmp = "Select TCS.EmployeeID,ME.EmployeeName + '~'+ ME.EmployeeNo As EmpName , ME.DesignationID,Sum(TCS.Amount) As TotalExpenseAmount"
tmp = tmp & " From Txn_ClaimSubmission TCS"
tmp = tmp & " Inner Join Mst_Employee ME On TCS.EmployeeID = ME.EmployeeID"
tmp = tmp & " Where TCS.ExpenseFrom <= '" & mEndDate & "' And TCS.ExpenseTo >= '" & mStartDate & "' And TCS.Flag <> 'R' And ME.DesignationID in (4,9,3)"
tmp = tmp & " And TCS.LocationID in ( " & mLocationIDs_ & " )"
tmp = tmp & " Group By TCS.EmployeeID,ME.EmployeeName, ME.EmployeeNo , ME.DesignationID Order by ME.DesignationID"

Call Tmp2Table

mFlag = False
If TmpRs2.RecordCount > 0 Then
   For r = 1 To TmpRs2.RecordCount
       If mDesgn_ = "AM" Then
          
          If TmpRs2!DesignationID = 3 Then
             mRow = MSHFlexGrid1.Rows - 1
             MSHFlexGrid1.TextMatrix(mRow, 0) = mRM_
             If mAmID_ = TmpRs2!EmployeeID Then
                MSHFlexGrid1.TextMatrix(mRow, 1) = "AM(Current)-" & TmpRs2!EmpName
                mFlag = True
             Else
                MSHFlexGrid1.TextMatrix(mRow, 1) = "AM-" & TmpRs2!EmpName
             End If
             MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          End If
       Else
          If TmpRs2!DesignationID = 4 Then
             mRow = MSHFlexGrid1.Rows - 1
             MSHFlexGrid1.TextMatrix(mRow, 0) = mRM_
             MSHFlexGrid1.TextMatrix(mRow, 1) = "CC-" & TmpRs2!EmpName
            
             MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          ElseIf TmpRs2!DesignationID = 9 Then
             mRow = MSHFlexGrid1.Rows - 1
             MSHFlexGrid1.TextMatrix(mRow, 0) = mRM_
             MSHFlexGrid1.TextMatrix(mRow, 1) = "BDO-" & TmpRs2!EmpName
             MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          End If
       End If
       TmpRs2.MoveNext
   Next
End If
If mFlag = False Then
   If mDesgn_ = "AM" And mAm_ <> "" Then
      mRow = MSHFlexGrid1.Rows - 1
      MSHFlexGrid1.TextMatrix(mRow, 0) = mRM_
      MSHFlexGrid1.TextMatrix(mRow, 1) = "AM(Current)-" & mAm_
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   End If
End If

End Sub


Private Sub FillCMPMExpense()
Dim mFlag As Boolean
Dim mMonths, mExpAmt, mTotMonths As Variant


tmp = "Select TCS.LocationID,ML.LocationName,Sum(TCS.Amount) As TotalExpenseAmount,TCS.ExpenseFrom,TCS.ExpenseTo"
tmp = tmp & " From Txn_ClaimSubmission TCS"
tmp = tmp & " Inner Join Mst_Employee ME On TCS.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Location ML on TCS.LocationID=ML.LocationID"
tmp = tmp & " Where TCS.ExpenseFrom <= '" & mEndDate & "' And TCS.ExpenseTo >= '" & mStartDate & "' And TCS.Flag <> 'R' And ME.DesignationID in (5,6,8)"
tmp = tmp & " Group By TCS.LocationID,ML.LocationName,TCS.ExpenseFrom,TCS.ExpenseTo"

Call Tmp2Table

mFlag = False
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       mMonths = DateDiff("m", mStartDate, mEndDate) + 1

       mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo) + 1
       mExpAmt = IIf(IsNull(TmpRs2!TotalExpenseAmount), 0, TmpRs2!TotalExpenseAmount) / mTotMonths
       If TmpRs2!ExpenseFrom < mStartDate Then
          mTotMonths = DateDiff("m", mStartDate, TmpRs2!ExpenseTo) + 1
       ElseIf TmpRs2!ExpenseTo > mEndDate Then
           mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, mEndDate) + 1
       Else
           mTotMonths = DateDiff("m", TmpRs2!ExpenseFrom, TmpRs2!ExpenseTo) + 1
       End If
       
       If mTotMonths > mMonths Then
          mTotExpAmt = mExpAmt * mMonths
       Else
          mTotExpAmt = mExpAmt * mTotMonths
       End If
       For r = 1 To MSHFlexGrid1.Rows - 1
           If MSHFlexGrid1.TextMatrix(r, 1) = TmpRs2!LocationName Then
              MSHFlexGrid1.TextMatrix(r, 4) = Val(MSHFlexGrid1.TextMatrix(r, 4)) + Format(mTotExpAmt, "######0.00")
           End If
       Next r
       TmpRs2.MoveNext
   Next
End If

End Sub

Private Function CheckLocation(mLocation_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False
For r = 1 To MSHFlexGrid1.Rows - 1
    If LCase(MSHFlexGrid1.TextMatrix(r, 1)) = LCase(mLocation_) Then
       Retval = True
    End If
Next
CheckLocation = Retval
End Function

Private Sub Grid_Head6()

MSHFlexGrid7.Clear
MSHFlexGrid7.Rows = 2
MSHFlexGrid7.Cols = 6
MSHFlexGrid7.FixedCols = 1
MSHFlexGrid7.FixedRows = 1
MSHFlexGrid7.TextMatrix(0, 0) = "Location Name"
MSHFlexGrid7.TextMatrix(0, 1) = "GodownID"
MSHFlexGrid7.TextMatrix(0, 2) = "Godown No"
MSHFlexGrid7.TextMatrix(0, 3) = "Address"
MSHFlexGrid7.TextMatrix(0, 4) = "Start Date"
MSHFlexGrid7.TextMatrix(0, 5) = "Close Date"

MSHFlexGrid7.ColWidth(0) = 2200
MSHFlexGrid7.ColWidth(1) = 1000
MSHFlexGrid7.ColWidth(2) = 1400
MSHFlexGrid7.ColWidth(3) = 3000
MSHFlexGrid7.ColWidth(4) = 1500
MSHFlexGrid7.ColWidth(5) = 1500


MSHFlexGrid7.RowHeight(0) = 600

End Sub


Private Sub FillNotAgreeGodowns()

'Commented On 04 Jan 2010
'tmp = "Select ML.LocationName,MG.GodownID,MG.GodownNo,MG.Address From Mst_Godown MG"
'tmp = tmp & " inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " Where MG.UnitTypeID=1 And (MG.CloseDate is Null or MG.CloseDate <='" & mEndDate & "') And MG.StartDate <='" & mEndDate & "' And MG.GodownID not in (Select distinct TGL.GodownID"
'tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
'tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
'tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
'tmp = tmp & " Where TLA.StartDate <= '" & mEndDate & "' And TLA.ExpiryDate>='2010-07-01' And (MG.CloseDate >= '" & mStartDate & "' or MG.CloseDate is null) And TLA.Flag<>'R')"


tmp = "Select ML.LocationName,MG.GodownID,MG.GodownNo,MG.Address,MG.StartDate,MG.CloseDate From Mst_Godown MG"
tmp = tmp & " inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where MG.UnitTypeID=1 And (MG.CloseDate is Null or (MG.CloseDate <='" & mEndDate & "' And MG.CloseDate >='" & mStartDate & "')) And MG.StartDate <='" & mEndDate & "' And MG.GodownID not in (Select distinct TGL.GodownID"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
tmp = tmp & " Where TLA.StartDate <= '" & mEndDate & "' And TLA.ExpiryDate>='" & mStartDate & "' And (MG.CloseDate >= '" & mStartDate & "' or MG.CloseDate is null) And TLA.Flag<>'R')"


Call Tmp2Table
Call Grid_Head6
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       r = MSHFlexGrid7.Rows - 1
       MSHFlexGrid7.TextMatrix(r, 0) = TmpRs2!LocationName
       MSHFlexGrid7.TextMatrix(r, 1) = TmpRs2!GodownID
       MSHFlexGrid7.TextMatrix(r, 2) = TmpRs2!GodownNo
       MSHFlexGrid7.TextMatrix(r, 3) = TmpRs2!Address
       MSHFlexGrid7.TextMatrix(r, 4) = IIf(IsNull(TmpRs2!StartDate), "", TmpRs2!StartDate)
       MSHFlexGrid7.TextMatrix(r, 5) = IIf(IsNull(TmpRs2!CloseDate), "", TmpRs2!CloseDate)
       MSHFlexGrid7.TextMatrix(r, 4) = Format(MSHFlexGrid7.TextMatrix(r, 4), "dd-mmm-yyyy")
       MSHFlexGrid7.TextMatrix(r, 5) = Format(MSHFlexGrid7.TextMatrix(r, 5), "dd-mmm-yyyy")
       MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + 1
       TmpRs2.MoveNext
   Next
End If
End Sub

