VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_AvgMonthStockMIS 
   Caption         =   "Average Month Stock MIS"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   9615
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   15030
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   7560
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   8040
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4215
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7435
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2895
         Left            =   120
         TabIndex        =   8
         Top             =   4560
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5106
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   8040
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   8040
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   8040
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
         Height          =   975
         Left            =   0
         TabIndex        =   18
         Top             =   8160
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
         Height          =   975
         Left            =   0
         TabIndex        =   19
         Top             =   8040
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid9 
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   7800
         Visible         =   0   'False
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   1720
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox ChkException 
         Appearance      =   0  'Flat
         Caption         =   "With Exception Details"
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
         Height          =   255
         Left            =   210
         TabIndex        =   21
         Top             =   750
         Width           =   2865
      End
      Begin VB.ComboBox CmbYear 
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
         ItemData        =   "FrmRpt_AvgMonthStockMIS.frx":0000
         Left            =   3330
         List            =   "FrmRpt_AvgMonthStockMIS.frx":0002
         TabIndex        =   11
         Top             =   322
         Width           =   1560
      End
      Begin VB.ComboBox CmbMonth 
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
         ItemData        =   "FrmRpt_AvgMonthStockMIS.frx":0004
         Left            =   855
         List            =   "FrmRpt_AvgMonthStockMIS.frx":0006
         TabIndex        =   10
         Top             =   322
         Width           =   1680
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
         Height          =   495
         Left            =   9060
         TabIndex        =   3
         Top             =   255
         Width           =   1335
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
         Height          =   495
         Left            =   6390
         TabIndex        =   2
         Top             =   255
         Width           =   1335
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
         Height          =   495
         Left            =   7725
         TabIndex        =   1
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         Left            =   2790
         TabIndex        =   13
         Top             =   382
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   195
         TabIndex        =   12
         Top             =   382
         Width           =   540
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
         Left            =   11280
         TabIndex        =   5
         Top             =   345
         Width           =   2415
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
         Left            =   13695
         TabIndex        =   4
         Top             =   345
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmRpt_AvgMonthStockMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTotDays As Variant
Dim mDate As Variant
Dim mYear As Variant
Dim mMonthNo As Variant
Dim mNoOfBags, mGrossWt, mGainLoss As Double


Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 10

MSHFlexGrid1.TextMatrix(0, 0) = "State"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "Stock Bags"
MSHFlexGrid1.TextMatrix(0, 5) = "Stock Weight"
MSHFlexGrid1.TextMatrix(0, 6) = "Date"
MSHFlexGrid1.TextMatrix(0, 7) = "Price"
MSHFlexGrid1.TextMatrix(0, 8) = "Amount"
MSHFlexGrid1.TextMatrix(0, 9) = "Gain / Loss"

MSHFlexGrid1.ColWidth(0) = 2200
MSHFlexGrid1.ColWidth(1) = 2200
MSHFlexGrid1.ColWidth(2) = 2200
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1400

End Sub
Private Sub Grid_Head2()

MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 2
MSHFlexGrid3.Cols = 12

MSHFlexGrid3.TextMatrix(0, 0) = "State"
MSHFlexGrid3.TextMatrix(0, 1) = "Location"
MSHFlexGrid3.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid3.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid3.TextMatrix(0, 4) = "Stock Bags"
MSHFlexGrid3.TextMatrix(0, 5) = "Stock Weight"
MSHFlexGrid3.TextMatrix(0, 6) = "Date"
MSHFlexGrid3.TextMatrix(0, 7) = "DepBags"
MSHFlexGrid3.TextMatrix(0, 8) = "DepWt"
MSHFlexGrid3.TextMatrix(0, 9) = "WitBags"
MSHFlexGrid3.TextMatrix(0, 10) = "WitWt"
MSHFlexGrid3.TextMatrix(0, 11) = "Weight Gain/Loss"

MSHFlexGrid3.ColWidth(0) = 2200
MSHFlexGrid3.ColWidth(1) = 2200
MSHFlexGrid3.ColWidth(2) = 2200
MSHFlexGrid3.ColWidth(3) = 2200
MSHFlexGrid3.ColWidth(4) = 1400
MSHFlexGrid3.ColWidth(5) = 1400
MSHFlexGrid3.ColWidth(6) = 1400

End Sub
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 10

MSHFlexGrid2.TextMatrix(0, 0) = "State"
MSHFlexGrid2.TextMatrix(0, 1) = "Location"
MSHFlexGrid2.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid2.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid2.TextMatrix(0, 4) = "Balance Bags"
MSHFlexGrid2.TextMatrix(0, 5) = "Balance Weight"
MSHFlexGrid2.TextMatrix(0, 6) = "Average Bags"
MSHFlexGrid2.TextMatrix(0, 7) = "Average Weight"
MSHFlexGrid2.TextMatrix(0, 8) = "Total Amount"
MSHFlexGrid2.TextMatrix(0, 9) = "Average Amount"

MSHFlexGrid2.ColWidth(0) = 2200
MSHFlexGrid2.ColWidth(1) = 2200
MSHFlexGrid2.ColWidth(2) = 2200
MSHFlexGrid2.ColWidth(3) = 2200
MSHFlexGrid2.ColWidth(4) = 1400
MSHFlexGrid2.ColWidth(5) = 1400
MSHFlexGrid2.ColWidth(6) = 1400
MSHFlexGrid2.ColWidth(7) = 1400
MSHFlexGrid2.ColWidth(8) = 1400
MSHFlexGrid2.ColWidth(9) = 1400

End Sub
Private Sub Grid_Head3()

MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 2
MSHFlexGrid4.Cols = 10

MSHFlexGrid4.TextMatrix(0, 0) = "State"
MSHFlexGrid4.TextMatrix(0, 1) = "Location"
MSHFlexGrid4.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid4.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid4.TextMatrix(0, 4) = "Stock Bags"
MSHFlexGrid4.TextMatrix(0, 5) = "Stock Weight"
MSHFlexGrid4.TextMatrix(0, 6) = "Date"
MSHFlexGrid4.TextMatrix(0, 7) = "Price"
MSHFlexGrid4.TextMatrix(0, 8) = "Amount"
MSHFlexGrid4.TextMatrix(0, 9) = "Gain / Loss"

MSHFlexGrid4.ColWidth(0) = 2200
MSHFlexGrid4.ColWidth(1) = 2200
MSHFlexGrid4.ColWidth(2) = 2200
MSHFlexGrid4.ColWidth(3) = 2200
MSHFlexGrid4.ColWidth(4) = 1400
MSHFlexGrid4.ColWidth(5) = 1400
MSHFlexGrid4.ColWidth(6) = 1400

End Sub
Private Sub Grid_Head4()

MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 2
MSHFlexGrid5.Cols = 10

MSHFlexGrid5.TextMatrix(0, 0) = "State"
MSHFlexGrid5.TextMatrix(0, 1) = "Location"
MSHFlexGrid5.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid5.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid5.TextMatrix(0, 4) = "Balance Bags"
MSHFlexGrid5.TextMatrix(0, 5) = "Balance Weight"
MSHFlexGrid5.TextMatrix(0, 6) = "Average Bags"
MSHFlexGrid5.TextMatrix(0, 7) = "Average Weight"
MSHFlexGrid5.TextMatrix(0, 8) = "Total Amount"
MSHFlexGrid5.TextMatrix(0, 9) = "Average Amount"

MSHFlexGrid5.ColWidth(0) = 2200
MSHFlexGrid5.ColWidth(1) = 2200
MSHFlexGrid5.ColWidth(2) = 2200
MSHFlexGrid5.ColWidth(3) = 2200
MSHFlexGrid5.ColWidth(4) = 1400
MSHFlexGrid5.ColWidth(5) = 1400
MSHFlexGrid5.ColWidth(6) = 1400
MSHFlexGrid5.ColWidth(7) = 1400
MSHFlexGrid5.ColWidth(8) = 1400
MSHFlexGrid5.ColWidth(9) = 1400

End Sub
Private Sub Grid_Head5()

MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 2
MSHFlexGrid6.Cols = 12

MSHFlexGrid6.TextMatrix(0, 0) = "State"
MSHFlexGrid6.TextMatrix(0, 1) = "Location"
MSHFlexGrid6.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid6.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid6.TextMatrix(0, 4) = "Stock Bags"
MSHFlexGrid6.TextMatrix(0, 5) = "Stock Weight"
MSHFlexGrid6.TextMatrix(0, 6) = "Date"
MSHFlexGrid6.TextMatrix(0, 7) = "DepBags"
MSHFlexGrid6.TextMatrix(0, 8) = "DepWt"
MSHFlexGrid6.TextMatrix(0, 9) = "WitBags"
MSHFlexGrid6.TextMatrix(0, 10) = "WitWt"
MSHFlexGrid6.TextMatrix(0, 11) = "Weight Gain/Loss"

MSHFlexGrid6.ColWidth(0) = 2200
MSHFlexGrid6.ColWidth(1) = 2200
MSHFlexGrid6.ColWidth(2) = 2200
MSHFlexGrid6.ColWidth(3) = 2200
MSHFlexGrid6.ColWidth(4) = 1400
MSHFlexGrid6.ColWidth(5) = 1400
MSHFlexGrid6.ColWidth(6) = 1400

End Sub
Private Sub Grid_Head6()

MSHFlexGrid7.Clear
MSHFlexGrid7.Rows = 2
MSHFlexGrid7.Cols = 10

MSHFlexGrid7.TextMatrix(0, 0) = "State"
MSHFlexGrid7.TextMatrix(0, 1) = "Location"
MSHFlexGrid7.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid7.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid7.TextMatrix(0, 4) = "Stock Bags"
MSHFlexGrid7.TextMatrix(0, 5) = "Stock Weight"
MSHFlexGrid7.TextMatrix(0, 6) = "Date"
MSHFlexGrid7.TextMatrix(0, 7) = "Price"
MSHFlexGrid7.TextMatrix(0, 8) = "Amount"
MSHFlexGrid7.TextMatrix(0, 9) = "Gain / Loss"

MSHFlexGrid7.ColWidth(0) = 2200
MSHFlexGrid7.ColWidth(1) = 2200
MSHFlexGrid7.ColWidth(2) = 2200
MSHFlexGrid7.ColWidth(3) = 2200
MSHFlexGrid7.ColWidth(4) = 1400
MSHFlexGrid7.ColWidth(5) = 1400
MSHFlexGrid7.ColWidth(6) = 1400

End Sub
Private Sub Grid_Head7()

MSHFlexGrid8.Clear
MSHFlexGrid8.Rows = 2
MSHFlexGrid8.Cols = 10

MSHFlexGrid8.TextMatrix(0, 0) = "State"
MSHFlexGrid8.TextMatrix(0, 1) = "Location"
MSHFlexGrid8.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid8.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid8.TextMatrix(0, 4) = "Balance Bags"
MSHFlexGrid8.TextMatrix(0, 5) = "Balance Weight"
MSHFlexGrid8.TextMatrix(0, 6) = "Average Bags"
MSHFlexGrid8.TextMatrix(0, 7) = "Average Weight"
MSHFlexGrid8.TextMatrix(0, 8) = "Total Amount"
MSHFlexGrid8.TextMatrix(0, 9) = "Average Amount"

MSHFlexGrid8.ColWidth(0) = 2200
MSHFlexGrid8.ColWidth(1) = 2200
MSHFlexGrid8.ColWidth(2) = 2200
MSHFlexGrid8.ColWidth(3) = 2200
MSHFlexGrid8.ColWidth(4) = 1400
MSHFlexGrid8.ColWidth(5) = 1400
MSHFlexGrid8.ColWidth(6) = 1400
MSHFlexGrid8.ColWidth(7) = 1400
MSHFlexGrid8.ColWidth(8) = 1400
MSHFlexGrid8.ColWidth(9) = 1400

End Sub
Private Sub Grid_Head8()

MSHFlexGrid9.Clear
MSHFlexGrid9.Rows = 2
MSHFlexGrid9.Cols = 12

MSHFlexGrid9.TextMatrix(0, 0) = "State"
MSHFlexGrid9.TextMatrix(0, 1) = "Location"
MSHFlexGrid9.TextMatrix(0, 2) = "Earthquake Zone"
MSHFlexGrid9.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid9.TextMatrix(0, 4) = "Stock Bags"
MSHFlexGrid9.TextMatrix(0, 5) = "Stock Weight"
MSHFlexGrid9.TextMatrix(0, 6) = "Date"
MSHFlexGrid9.TextMatrix(0, 7) = "DepBags"
MSHFlexGrid9.TextMatrix(0, 8) = "DepWt"
MSHFlexGrid9.TextMatrix(0, 9) = "WitBags"
MSHFlexGrid9.TextMatrix(0, 10) = "WitWt"
MSHFlexGrid9.TextMatrix(0, 11) = "Weight Gain/Loss"

MSHFlexGrid9.ColWidth(0) = 2200
MSHFlexGrid9.ColWidth(1) = 2200
MSHFlexGrid9.ColWidth(2) = 2200
MSHFlexGrid9.ColWidth(3) = 2200
MSHFlexGrid9.ColWidth(4) = 1400
MSHFlexGrid9.ColWidth(5) = 1400
MSHFlexGrid9.ColWidth(6) = 1400

End Sub

Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For i = 1 To 12
    CmbMonth.AddItem MonthName(i)
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
If CmbMonth.Text = "" Then
   Exit Sub
End If
End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For i = 2009 To Year(Now)
    CmbYear.AddItem i
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If CmbYear.Text = "" Then
   Exit Sub
End If
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head8
Label2.Caption = ""
CmdExcel.Enabled = False
ChkException.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_AvgMonthStockMIS_" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()

Dim SheetNo As Double
Dim SheetName As Variant

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String

Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_AvgMonthStockMIS_" & CmbMonth.Text & "-" & CmbYear.Text & ".xls")

'below code added by prajakta on 9 Mar 15
SheetNo = oWB.Worksheets.Count
If SheetNo < 15 Then
   For iCtr = SheetNo To 15
       SheetName = "Sheet" & iCtr
       oWB.Worksheets.Add , oWB.Worksheets(SheetName), 1
   Next
End If

Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Totalstock"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next


Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "TotalStockSummary"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid2.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid2.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(i, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "Stockofenddateofmonth"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid3.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid3.Cols - 1
        If InStr(MSHFlexGrid3.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid3.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid3.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(i, c)
        End If
    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "UninsuredStockDayWise"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid4.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid4.Cols - 1
        If InStr(MSHFlexGrid4.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid4.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid4.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(i, c)
        End If
    Next
Next
Set oWS = oWB.Worksheets("Sheet5")
oWS.Name = "UninsuredStockSummary"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid5.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid5.Cols - 1
        If InStr(MSHFlexGrid5.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid5.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid5.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid5.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid5.TextMatrix(i, c)
        End If
    Next
Next
Set oWS = oWB.Worksheets("Sheet6")
oWS.Name = "Uninsured stock end of month"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid6.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid6.Cols - 1
        If InStr(MSHFlexGrid6.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid6.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid6.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid6.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid6.TextMatrix(i, c)
        End If
    Next
Next
Set oWS = oWB.Worksheets("Sheet7")
oWS.Name = "CottonbalesDayWise"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid7.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid7.Cols - 1
        If InStr(MSHFlexGrid7.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid7.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid7.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid7.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid7.TextMatrix(i, c)
        End If
    Next
Next
Set oWS = oWB.Worksheets("Sheet8")
oWS.Name = "CottonBalesSummary"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid8.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid8.Cols - 1
        If InStr(MSHFlexGrid8.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid8.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid8.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid8.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid8.TextMatrix(i, c)
        End If
    Next
Next
Set oWS = oWB.Worksheets("Sheet9")
oWS.Name = "CottonBales end of the month"
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid9.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid9.Cols - 1
        If InStr(MSHFlexGrid9.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid9.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid9.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid9.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid9.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub CmdGetReport_Click()
Dim mDayNo As Variant


If CmbMonth.Text = "" Then
   MsgBox "Please Select Month."
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox "Please Select Year."
   Exit Sub
End If

Call Grid_Head
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head8

MsgBox "Wait till next message !!!"
Label2.Caption = ""
mYear = CmbYear.Text
mMonthNo = Month(mYear & "-" & CmbMonth.Text & "-" & 1)
mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))

'tmp = "Select ML.LocationName,MCm.Commodity"
'tmp = tmp & " From Txn_CxTF_Details TC"
'tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity"
'tmp = tmp & " order by ML.LocationName,MCm.Commodity"

'Commented on 11 Oct 2012 After mail from Sandeep Mohite regarding -ve Balance Bags
'tmp = "Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
'tmp = tmp & " From Txn_CxTF_Details TC"
'tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'tmp = tmp & " order by MS.StateName,ML.LocationName,MCm.Commodity"
''



tmp = " Select StateName,LocationName,EQZone,Commodity From"
tmp = tmp & " (("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_CxTF_Details TC"
tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
'==========================added transaction details on 19 apr 2013 ==========================
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_AdjTransaction TAD"
tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
'====================================================
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_Spillage TC"
tmp = tmp & " inner Join Mst_GSL MG On TC.GSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_Spillage TC"
tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )) TempTable"
tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity"


Call TmpTable

MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 1
For n = 1 To TmpRs.RecordCount
    MSHFlexGrid2.TextMatrix(n, 0) = TmpRs!StateName
    MSHFlexGrid2.TextMatrix(n, 1) = TmpRs!LocationName
    MSHFlexGrid2.TextMatrix(n, 2) = IIf(IsNull(TmpRs!EQZone), "", TmpRs!EQZone)
    MSHFlexGrid2.TextMatrix(n, 3) = TmpRs!Commodity
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
TmpRs.MoveFirst
    
ProgressBar1.Value = 0
ProgressBar1.Max = mTotDays + 1
For r1 = 1 To mTotDays
    
    Call Grid_Head2
    mDate = mYear & "-" & mMonthNo & "-" & r1

'    tmp = "Select ML.LocationName,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' "
'    tmp = tmp & " group by ML.LocationName,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by ML.LocationName,MCm.Commodity,TC.TxnType"
    
    
'Commented on 11 Oct 2012 After mail from Sandeep Mohite regarding -ve Balance Bags
'    tmp = "Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'"
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by MS.StateName,ML.LocationName,MCm.Commodity,TC.TxnType"
    
    

    tmp = " Select StateName,LocationName,EQZone,Commodity,TxnType,Sum(NoOfBags)'NoOfBags',Sum(GrossWt)'GrossWt',Sum(GainLoss) 'GainLoss' From"
    tmp = tmp & " (("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight) + Sum(isnull(TG.WeightGainLossGSL,0)) 'GrossWt' , Sum(isnull(TG.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_CxTF_Details TC"
    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
    tmp = tmp & " )"
    
    '==========================added transaction details on 19 apr 2013 ==========================
    
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'D'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType = 'AR' "
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
    
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'W'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType <> 'AR' "
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
'====================================================
    
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
    tmp = tmp & " From Txn_Spillage TC"
    tmp = tmp & " inner Join Mst_GSL MG On TC.GSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix  And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
    tmp = tmp & " From Txn_Spillage TC"
    tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )) TempTable"
    tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity,TxnType"
    
        
    
    Call Tmp2Table
    
    If TmpRs2.RecordCount > 0 Then
       For r = 1 To TmpRs2.RecordCount
           Call IndataGrid1
           Call IndataGrid2
           TmpRs2.MoveNext
       Next
    End If
    Call CopyGrid3ToGrid1
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max


tmp = "select MS.StateName,MC.Commodity ,MCP.* from Mst_CommodityPrice MCP"
tmp = tmp & " inner join Mst_State MS on MCP.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MC on MCP.CommodityID=MC.CommodityID"
'tmp = tmp & " Where MCP.MonthYear='" & Left(CmbMonth.Text, 3) & "-" & mYear & "'"
tmp = tmp & " Where MCP.YearMonth='" & Format(CDate(mDate), "yyyy-mm") & "'"

Call Tmp3Table
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For i = 1 To MSHFlexGrid1.Rows - 2
  
    TmpRs3.Filter = "StateName='" & MSHFlexGrid1.TextMatrix(i, 0) & "' And Commodity='" & MSHFlexGrid1.TextMatrix(i, 3) & "'"
    mDayNo = Day(MSHFlexGrid1.TextMatrix(i, 6))
    If TmpRs3.RecordCount > 0 Then
       'MSHFlexGrid1.TextMatrix(I, 7) = TmpRs3.Fields(mDayNo + 4)
       
       MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(TmpRs3.Fields(mDayNo + 4)), 0, TmpRs3.Fields(mDayNo + 4))
       
    Else
       MSHFlexGrid1.TextMatrix(i, 7) = 0
    End If
    MSHFlexGrid1.TextMatrix(i, 8) = Round(Val(MSHFlexGrid1.TextMatrix(i, 7)) * Val(MSHFlexGrid1.TextMatrix(i, 5)), 0)
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Value + 1
Call CalAvgPrice
Label2.Caption = MSHFlexGrid1.Rows - 2
ProgressBar1.Value = ProgressBar1.Max


If ChkException.Value = 1 Then
   
   tmp = "Select Distinct LocationID, ClientIDRow from Txn_InsuranceExceptionMapping"
   tmp = tmp & " Where ClientIDRow Is Not Null"
   Call Tmp5Table

   If Not TmpRs5.EOF Then
      Call DisplayDataLocationandclientwise
   End If
   
   tmp = "Select Distinct LocationID, CommodityID from Txn_InsuranceExceptionMapping"
   tmp = tmp & " Where CommodityID Is Not Null"
   Call Tmp5Table

   If Not TmpRs5.EOF Then
      Call DisplayDataLocationandcommoditywise
   End If
End If

CmdExcel.Enabled = True
MsgBox "Done!!"
End Sub
Private Sub DisplayDataLocationandcommoditywise()
Dim mLocationID, mCommodityID As Variant
Dim wc As String
'Dim mGen_mTimeStamp As Variant
'Dim mTable As String
'mGen_mTimeStamp = Gen_mTimeStamp
'mTable = "" & "tmp" & mGen_mTimeStamp & Gen_UserLoginID

'Call Grid_Head3
'Call Grid_Head4

Call Grid_Head6
Call Grid_Head7

wc = ""
tmp = "Select Distinct LocationID, COmmodityID from Txn_InsuranceExceptionMapping"
tmp = tmp & " Where CommodityID Is Not Null"
Call Tmp5Table
If Not TmpRs5.EOF Then
   'TmpRs5.MoveFirst
   For i = 1 To TmpRs5.RecordCount
       If i = TmpRs5.RecordCount Then
          wc = wc & "ML.LocationID = " & TmpRs5!LocationID & " And " & "MCm.CommodityID =" & TmpRs5!CommodityID
          Exit For
       End If
       wc = wc & "ML.LocationID = " & TmpRs5!LocationID & " And " & "MCm.CommodityID =" & TmpRs5!CommodityID & " or "
       TmpRs5.MoveNext
   Next
End If



tmp = " Select StateName,LocationName,EQZone,Commodity From"
tmp = tmp & " (("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_CxTF_Details TC"
tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
If wc <> "" Then
   tmp = tmp & " And (" & wc & ")"
End If
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
'==========================added transaction details on 19 apr 2013 ==========================
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_AdjTransaction TAD"
tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
If wc <> "" Then
   tmp = tmp & " And (" & wc & ")"
End If
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
'====================================================
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_Spillage TC"
tmp = tmp & " inner Join Mst_GSL MG On TC.GSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
If wc <> "" Then
   tmp = tmp & " And (" & wc & ")"
End If
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_Spillage TC"
tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
If wc <> "" Then
   tmp = tmp & " And (" & wc & ")"
End If
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )) TempTable"
tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity"


Call TmpTable

MSHFlexGrid8.Rows = TmpRs.RecordCount + 2
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 1
For n = 1 To TmpRs.RecordCount
    MSHFlexGrid8.TextMatrix(n, 0) = TmpRs!StateName
    MSHFlexGrid8.TextMatrix(n, 1) = TmpRs!LocationName
    MSHFlexGrid8.TextMatrix(n, 2) = IIf(IsNull(TmpRs!EQZone), "", TmpRs!EQZone)
    MSHFlexGrid8.TextMatrix(n, 3) = TmpRs!Commodity
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
TmpRs.MoveFirst
    
ProgressBar1.Value = 0
ProgressBar1.Max = mTotDays + 1
For r1 = 1 To mTotDays
    
'   Call Grid_Head2
'   Call Grid_Head5
    Call Grid_Head8
    If r1 = 12 Then
       a = a
    End If
    mDate = mYear & "-" & mMonthNo & "-" & r1

'    tmp = "Select ML.LocationName,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' "
'    tmp = tmp & " group by ML.LocationName,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by ML.LocationName,MCm.Commodity,TC.TxnType"
    
    
'Commented on 11 Oct 2012 After mail from Sandeep Mohite regarding -ve Balance Bags
'    tmp = "Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'"
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by MS.StateName,ML.LocationName,MCm.Commodity,TC.TxnType"
    tmp = " Select StateName,LocationName,EQZone,Commodity,TxnType,StateName+LocationName+EQZone+Commodity+TxnType as KeyName,Sum(NoOfBags)'NoOfBags',Sum(GrossWt)'GrossWt',Sum(GainLoss) 'GainLoss'  From"
    tmp = tmp & " (("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight) + Sum(isnull(TG.WeightGainLossGSL,0)) 'GrossWt' , Sum(isnull(TG.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_CxTF_Details TC"
    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = TG.ClientIDRow"
    tmp = tmp & " Inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
    tmp = tmp & " )"
    
    '==========================added transaction details on 19 apr 2013 ==========================
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'D'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_GSL MG on MG.CMPID = MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
    tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType = 'AR' "
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
    
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'W'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_GSL MG on MG.CMPID = MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
    tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType <> 'AR' "
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
'====================================================
    
'    tmp = tmp & " Union All"
'    tmp = tmp & " ("
'    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
'    tmp = tmp & " From Txn_Spillage TC"
'    tmp = tmp & " inner Join Mst_GSL MG On TC.GSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix  And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
'    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'    tmp = tmp & " )"
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
    tmp = tmp & " From Txn_Spillage TC"
    tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
    tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )) TempTable"
    tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity,TxnType"

    Call Tmp6Table


    
    tmp = " Select StateName,LocationName,EQZone,Commodity,TxnType,StateName+LocationName+EQZone+Commodity+TxnType as KeyName,Sum(NoOfBags)'NoOfBags',Sum(GrossWt)'GrossWt',Sum(GainLoss) 'GainLoss' From"
    tmp = tmp & " (("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight) + Sum(isnull(TG.WeightGainLossGSL,0)) 'GrossWt' , Sum(isnull(TG.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_CxTF_Details TC"
    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    If wc <> "" Then
       tmp = tmp & " And (" & wc & ")"
    End If
    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
    tmp = tmp & " )"
    
    '==========================added transaction details on 19 apr 2013 ==========================
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'D'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType = 'AR' "
    If wc <> "" Then
       tmp = tmp & " And (" & wc & ")"
    End If
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
    
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'W'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType <> 'AR' "
    If wc <> "" Then
       tmp = tmp & " And (" & wc & ")"
    End If
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
'====================================================
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
    tmp = tmp & " From Txn_Spillage TC"
    tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    If wc <> "" Then
       tmp = tmp & " And (" & wc & ")"
    End If
    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )) TempTable"
    tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity,TxnType"
    
    
    Call Tmp2Table
    
'    If Not TmpRs2.EOF Then
'       TmpRs2.MoveFirst
'       For i = 1 To TmpRs2.RecordCount
'           TmpRs6.Filter = "StateName='" & TmpRs2!StateName & "' And LocationName='" & TmpRs2!LocationName & "'And EQZone='" & TmpRs2!EQZone & "'And Commodity='" & TmpRs2!Commodity & "'And TxnType='" & TmpRs2!TxnType & "'"
'           If TmpRs6.RecordCount > 0 Then
'
'              TmpRs2!NoofBags = TmpRs2!NoofBags - TmpRs6!NoofBags
'              TmpRs2!GrossWt = TmpRs2!GrossWt - TmpRs6!GrossWt
'              TmpRs2!GainLoss = TmpRs2!GainLoss - TmpRs6!GainLoss
'              TmpRs2.Update
'           End If
'           TmpRs6.Filter = ""
'           TmpRs2.MoveNext
'       Next
'    End If
    
    If TmpRs2.RecordCount > 0 Then
       For r = 1 To TmpRs2.RecordCount
           'TmpRs6.Filter = "StateName='" & TmpRs2!stateName & "' And LocationName='" & TmpRs2!LocationName & "'And EQZone='" & TmpRs2!EQZone & "'And Commodity='" & TmpRs2!Commodity & "'And TxnType='" & TmpRs2!TxnType & "'"
           
           TmpRs6.Find "KeyName = '" & TmpRs2!KeyName & "'"
           'If TmpRs6.RecordCount > 0 Then
           If Not TmpRs6.EOF Then
              mNoOfBags = 0
              mGrossWt = 0
              mGainLoss = 0
              mNoOfBags = TmpRs6!NoofBags
              mGrossWt = TmpRs6!GrossWt
              mGainLoss = TmpRs6!GainLoss
           Else
              mNoOfBags = 0
              mGrossWt = 0
              mGainLoss = 0
           End If
           'Call IndataGrid1
           'Call IndataGrid4
           Call IndataGrid6
           'Call IndataGrid2
           'Call IndataGrid5
           Call IndataGrid7
           TmpRs6.Filter = ""
           TmpRs2.MoveNext
       Next
    End If
    'Call CopyGrid3ToGrid1
    'Call CopyGrid6ToGrid4
     Call CopyGrid9ToGrid7
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max


tmp = "select MS.StateName,MC.Commodity ,MCP.* from Mst_CommodityPrice MCP"
tmp = tmp & " inner join Mst_State MS on MCP.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MC on MCP.CommodityID=MC.CommodityID"
'tmp = tmp & " Where MCP.MonthYear='" & Left(CmbMonth.Text, 3) & "-" & mYear & "'"
tmp = tmp & " Where MCP.YearMonth='" & Format(CDate(mDate), "yyyy-mm") & "'"

Call Tmp3Table
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid7.Rows + 1
For i = 1 To MSHFlexGrid7.Rows - 2
  
    TmpRs3.Filter = "StateName='" & MSHFlexGrid7.TextMatrix(i, 0) & "' And Commodity='" & MSHFlexGrid7.TextMatrix(i, 3) & "'"
    mDayNo = Day(MSHFlexGrid7.TextMatrix(i, 6))
    If TmpRs3.RecordCount > 0 Then
       'MSHFlexGrid7.TextMatrix(I, 7) = TmpRs3.Fields(mDayNo + 4)
       
       MSHFlexGrid7.TextMatrix(i, 7) = IIf(IsNull(TmpRs3.Fields(mDayNo + 4)), 0, TmpRs3.Fields(mDayNo + 4))
       
    Else
       MSHFlexGrid7.TextMatrix(i, 7) = 0
    End If
    MSHFlexGrid7.TextMatrix(i, 8) = Round(Val(MSHFlexGrid7.TextMatrix(i, 7)) * Val(MSHFlexGrid7.TextMatrix(i, 5)), 0)
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Value + 1
'Call CalAvgPriceLocandClienwise
CalAvgPriceLocandCommoditywise
'Label2.Caption = MSHFlexGrid7.Rows - 2
ProgressBar1.Value = ProgressBar1.Max

End Sub
Private Sub DisplayDataLocationandclientwise()

Call Grid_Head3
Call Grid_Head4

tmp = " Select StateName,LocationName,EQZone,Commodity From"
tmp = tmp & " (("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_CxTF_Details TC"
tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = TG.ClientIDRow"
tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
'==========================added transaction details on 19 apr 2013 ==========================
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_AdjTransaction TAD"
tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
tmp = tmp & " inner Join Mst_GSL MG on MG.CMPID = MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )"
'====================================================
'tmp = tmp & " Union All"
'tmp = tmp & " ("
'tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
'tmp = tmp & " From Txn_Spillage TC"
'tmp = tmp & " inner Join Mst_GSL MG On TC.GSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
'tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
'tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'tmp = tmp & " )"
tmp = tmp & " Union All"
tmp = tmp & " ("
tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity"
tmp = tmp & " From Txn_Spillage TC"
tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & mTotDays & "' "
tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
tmp = tmp & " )) TempTable"
tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity"


Call TmpTable

MSHFlexGrid5.Rows = TmpRs.RecordCount + 2
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 1
For n = 1 To TmpRs.RecordCount
    MSHFlexGrid5.TextMatrix(n, 0) = TmpRs!StateName
    MSHFlexGrid5.TextMatrix(n, 1) = TmpRs!LocationName
    MSHFlexGrid5.TextMatrix(n, 2) = IIf(IsNull(TmpRs!EQZone), "", TmpRs!EQZone)
    MSHFlexGrid5.TextMatrix(n, 3) = TmpRs!Commodity
    TmpRs.MoveNext
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
TmpRs.MoveFirst
    
ProgressBar1.Value = 0
ProgressBar1.Max = mTotDays + 1
For r1 = 1 To mTotDays
    
    'Call Grid_Head2
    Call Grid_Head5
    If r1 = 12 Then
       a = a
    End If
    mDate = mYear & "-" & mMonthNo & "-" & r1

'    tmp = "Select ML.LocationName,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' "
'    tmp = tmp & " group by ML.LocationName,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by ML.LocationName,MCm.Commodity,TC.TxnType"
    
    
'Commented on 11 Oct 2012 After mail from Sandeep Mohite regarding -ve Balance Bags
'    tmp = "Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
'    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight)'GrossWt'"
'    tmp = tmp & " From Txn_CxTF_Details TC"
'    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
'    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'"
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
'    tmp = tmp & " order by MS.StateName,ML.LocationName,MCm.Commodity,TC.TxnType"
    
    

    tmp = " Select StateName,LocationName,EQZone,Commodity,TxnType,Sum(NoOfBags)'NoOfBags',Sum(GrossWt)'GrossWt',Sum(GainLoss) 'GainLoss' From"
    tmp = tmp & " (("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,TC.TxnType"
    tmp = tmp & " ,sum(TG.NoOfBags) 'NoOfBags',sum(TG.GrossWeight) + Sum(isnull(TG.WeightGainLossGSL,0)) 'GrossWt' , Sum(isnull(TG.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_CxTF_Details TC"
    tmp = tmp & " inner join Txn_CxTF_GSL TG on TC.CxTFNo=TG.CxTFNo and TC.SM_Prefix=TG.SM_Prefix  And TC.TxnType = TG.TxnType"
    tmp = tmp & " inner join Mst_CMP MC on TC.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = TG.ClientIDRow"
    tmp = tmp & " Inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.CxTFDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    tmp = tmp & " group by MS.StateName,ML.LocationName,MEQ.EQZone,TC.TxnType,MCm.Commodity"
    tmp = tmp & " )"
    
    '==========================added transaction details on 19 apr 2013 ==========================
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'D'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_GSL MG on MG.CMPID = MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
    tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType = 'AR' "
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
    
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity, 'W'"
    tmp = tmp & " ,sum(TGAD.NoOfBags) 'NoOfBags',sum(TGAD.Quantity) + Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GrossWt', Sum(isnull(TGAD.WeightGainLossGSL,0)) 'GainLoss'"
    tmp = tmp & " From Txn_AdjTransaction TAD"
    tmp = tmp & " inner join Txn_GSLAdjTransaction TGAD on TGAD.txnid = TAD.txnid"
    tmp = tmp & " inner join Mst_CMP MC on TAD.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Mst_GSL MG on MG.CMPID = MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
    tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TAD.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TAD.TxnDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "' And TAD.TxnType <> 'AR' "
    tmp = tmp & " group by ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )"
'====================================================
    
'    tmp = tmp & " Union All"
'    tmp = tmp & " ("
'    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
'    tmp = tmp & " From Txn_Spillage TC"
'    tmp = tmp & " inner Join Mst_GSL MG On TC.GSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix  And TC.Flag = 'S' " 'And MG.SpillageFLag = 1"
'    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
'    tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
'    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
'    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
'    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
'    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
'    tmp = tmp & " )"
    tmp = tmp & " Union All"
    tmp = tmp & " ("
    tmp = tmp & " Select MS.StateName,ML.LocationName,MEQ.EQZone,MCm.Commodity,'D',sum(TC.SpillageNoOfBags) 'NoOfBags',sum(TC.SpillageQuantity)'GrossWt',0"
    tmp = tmp & " From Txn_Spillage TC"
    tmp = tmp & " inner Join Mst_GSL MG On TC.NewGSLID = MG.GSLID And TC.SM_Prefix = MG.SM_Prefix And MG.MadeUpFlag = 1"
    tmp = tmp & " inner join Mst_CMP MC on MG.CMPID=MC.CMPID"
    tmp = tmp & " inner Join Txn_InsuranceExceptionMapping as TEM on TEM.LocationID = MC.LocationID And TEM.ClientIDRow = MG.ClientIDRow"
    tmp = tmp & " inner join Mst_Location ML on TEM.LocationID = ML.LocationID"
    'tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
    tmp = tmp & " left join Mst_EarthquakeZone MEQ on MEQ.EQZoneID=ML.EQZoneID"
    tmp = tmp & " inner join Mst_State MS on ML.StateID=MS.StateID"
    tmp = tmp & " inner join Mst_Commodity MCm on TC.CommodityID=MCm.CommodityID"
    tmp = tmp & " where TC.SpillageDate<='" & mYear & "-" & mMonthNo & "-" & r1 & "'  "
    tmp = tmp & " group by MS.StateName,ML.LocationName,MCm.Commodity,MS.StateName,MEQ.EQZone"
    tmp = tmp & " )) TempTable"
    tmp = tmp & " Group By StateName,LocationName,EQZone,Commodity,TxnType"
    Call Tmp2Table
    If TmpRs2.RecordCount > 0 Then
       For r = 1 To TmpRs2.RecordCount
           'Call IndataGrid1
           Call IndataGrid4
           'Call IndataGrid2
           Call IndataGrid5
           TmpRs2.MoveNext
       Next
    End If
    'Call CopyGrid3ToGrid1
     Call CopyGrid6ToGrid4
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max


tmp = "select MS.StateName,MC.Commodity ,MCP.* from Mst_CommodityPrice MCP"
tmp = tmp & " inner join Mst_State MS on MCP.StateID=MS.StateID"
tmp = tmp & " inner join Mst_Commodity MC on MCP.CommodityID=MC.CommodityID"
'tmp = tmp & " Where MCP.MonthYear='" & Left(CmbMonth.Text, 3) & "-" & mYear & "'"
tmp = tmp & " Where MCP.YearMonth='" & Format(CDate(mDate), "yyyy-mm") & "'"

Call Tmp3Table
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid4.Rows + 1
For i = 1 To MSHFlexGrid4.Rows - 2
  
    TmpRs3.Filter = "StateName='" & MSHFlexGrid4.TextMatrix(i, 0) & "' And Commodity='" & MSHFlexGrid4.TextMatrix(i, 3) & "'"
    mDayNo = Day(MSHFlexGrid4.TextMatrix(i, 6))
    If TmpRs3.RecordCount > 0 Then
       'MSHFlexGrid4.TextMatrix(I, 7) = TmpRs3.Fields(mDayNo + 4)
       
       MSHFlexGrid4.TextMatrix(i, 7) = IIf(IsNull(TmpRs3.Fields(mDayNo + 4)), 0, TmpRs3.Fields(mDayNo + 4))
       
    Else
       MSHFlexGrid4.TextMatrix(i, 7) = 0
    End If
    MSHFlexGrid4.TextMatrix(i, 8) = Round(Val(MSHFlexGrid4.TextMatrix(i, 7)) * Val(MSHFlexGrid4.TextMatrix(i, 5)), 0)
    ProgressBar1.Value = ProgressBar1.Value + 1
Next

ProgressBar1.Value = ProgressBar1.Value + 1
Call CalAvgPriceLocandClienwise
'Label2.Caption = MSHFlexGrid4.Rows - 2
ProgressBar1.Value = ProgressBar1.Max
End Sub
Private Sub CalAvgPriceLocandCommoditywise()
Dim mTotAmt As Double
For i = 1 To MSHFlexGrid8.Rows - 2
    mTotAmt = 0
    For j = 0 To MSHFlexGrid7.Rows - 2
        If MSHFlexGrid7.TextMatrix(j, 1) = MSHFlexGrid8.TextMatrix(i, 1) And MSHFlexGrid7.TextMatrix(j, 3) = MSHFlexGrid8.TextMatrix(i, 3) Then
           mTotAmt = mTotAmt + Val(MSHFlexGrid7.TextMatrix(j, 8))
        End If
    Next
    MSHFlexGrid8.TextMatrix(i, 8) = Round(mTotAmt)
    MSHFlexGrid8.TextMatrix(i, 9) = Round(Val(MSHFlexGrid8.TextMatrix(i, 8)) / mTotDays, 0)
Next
End Sub
Private Sub CalAvgPriceLocandClienwise()
Dim mTotAmt As Double
For i = 1 To MSHFlexGrid5.Rows - 2
    mTotAmt = 0
    For j = 0 To MSHFlexGrid4.Rows - 2
        If MSHFlexGrid4.TextMatrix(j, 1) = MSHFlexGrid5.TextMatrix(i, 1) And MSHFlexGrid4.TextMatrix(j, 3) = MSHFlexGrid5.TextMatrix(i, 3) Then
           mTotAmt = mTotAmt + Val(MSHFlexGrid4.TextMatrix(j, 8))
        End If
    Next
    MSHFlexGrid5.TextMatrix(i, 8) = Round(mTotAmt)
    MSHFlexGrid5.TextMatrix(i, 9) = Round(Val(MSHFlexGrid5.TextMatrix(i, 8)) / mTotDays, 0)
Next
End Sub
Private Sub CopyGrid9ToGrid7()
For i = 1 To MSHFlexGrid9.Rows - 2
    mr = MSHFlexGrid7.Rows - 1
    For c = 0 To 6
        MSHFlexGrid7.TextMatrix(mr, c) = MSHFlexGrid9.TextMatrix(i, c)
    Next
    MSHFlexGrid7.TextMatrix(mr, 9) = MSHFlexGrid9.TextMatrix(i, 11)
    MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + 1
Next
End Sub
Private Sub CopyGrid6ToGrid4()
For i = 1 To MSHFlexGrid6.Rows - 2
    mr = MSHFlexGrid4.Rows - 1
    For c = 0 To 6
        MSHFlexGrid4.TextMatrix(mr, c) = MSHFlexGrid6.TextMatrix(i, c)
    Next
    MSHFlexGrid4.TextMatrix(mr, 9) = MSHFlexGrid6.TextMatrix(i, 11)
    MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
Next
End Sub
Private Sub IndataGrid5()
For m = 1 To MSHFlexGrid5.Rows - 1
    If TmpRs2!LocationName = MSHFlexGrid5.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid5.TextMatrix(m, 3) Then
       If TmpRs2!TxnType = "D" Then
          MSHFlexGrid5.TextMatrix(m, 4) = Val(MSHFlexGrid5.TextMatrix(m, 4)) + (TmpRs2!NoofBags - mNoOfBags)
          MSHFlexGrid5.TextMatrix(m, 5) = Val(MSHFlexGrid5.TextMatrix(m, 5)) + (TmpRs2!GrossWt - mGrossWt)
       Else
          MSHFlexGrid5.TextMatrix(m, 4) = Val(MSHFlexGrid5.TextMatrix(m, 4)) - (TmpRs2!NoofBags - mNoOfBags)
          MSHFlexGrid5.TextMatrix(m, 5) = Val(MSHFlexGrid5.TextMatrix(m, 5)) - (TmpRs2!GrossWt - mGrossWt)
       End If
       If Val(MSHFlexGrid5.TextMatrix(m, 5)) < 0 Then
          MSHFlexGrid5.TextMatrix(m, 5) = 0
       End If
       If Val(MSHFlexGrid5.TextMatrix(m, 4)) < 0 Then
          MSHFlexGrid5.TextMatrix(m, 4) = 0
       End If
       MSHFlexGrid5.TextMatrix(m, 6) = Round(Val(MSHFlexGrid5.TextMatrix(m, 4)) / mTotDays, 0)
       MSHFlexGrid5.TextMatrix(m, 7) = Round(Val(MSHFlexGrid5.TextMatrix(m, 5)) / mTotDays, 3)
       If Val(MSHFlexGrid5.TextMatrix(m, 6)) < 0 Then
          MSHFlexGrid5.TextMatrix(m, 6) = 0
       End If
       If Val(MSHFlexGrid5.TextMatrix(m, 7)) < 0 Then
          MSHFlexGrid5.TextMatrix(m, 7) = 0
       End If
   End If
Next
End Sub
Private Sub IndataGrid7()
For m = 1 To MSHFlexGrid8.Rows - 1
    If TmpRs2!LocationName = MSHFlexGrid8.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid8.TextMatrix(m, 3) Then
       If TmpRs2!TxnType = "D" Then
          MSHFlexGrid8.TextMatrix(m, 4) = Val(MSHFlexGrid8.TextMatrix(m, 4)) + (TmpRs2!NoofBags - mNoOfBags)
          MSHFlexGrid8.TextMatrix(m, 5) = Val(MSHFlexGrid8.TextMatrix(m, 5)) + (TmpRs2!GrossWt - mGrossWt)
       Else
          MSHFlexGrid8.TextMatrix(m, 4) = Val(MSHFlexGrid8.TextMatrix(m, 4)) - (TmpRs2!NoofBags - mNoOfBags)
          MSHFlexGrid8.TextMatrix(m, 5) = Val(MSHFlexGrid8.TextMatrix(m, 5)) - (TmpRs2!GrossWt - mGrossWt)
       End If
       If Val(MSHFlexGrid8.TextMatrix(m, 5)) < 0 Then
          MSHFlexGrid8.TextMatrix(m, 5) = 0
       End If
       If Val(MSHFlexGrid8.TextMatrix(m, 4)) < 0 Then
          MSHFlexGrid8.TextMatrix(m, 4) = 0
       End If
       MSHFlexGrid8.TextMatrix(m, 6) = Round(Val(MSHFlexGrid8.TextMatrix(m, 4)) / mTotDays, 0)
       MSHFlexGrid8.TextMatrix(m, 7) = Round(Val(MSHFlexGrid8.TextMatrix(m, 5)) / mTotDays, 3)
       If Val(MSHFlexGrid8.TextMatrix(m, 6)) < 0 Then
          MSHFlexGrid8.TextMatrix(m, 6) = 0
       End If
       If Val(MSHFlexGrid8.TextMatrix(m, 7)) < 0 Then
          MSHFlexGrid8.TextMatrix(m, 7) = 0
       End If
   End If
Next
End Sub

Private Sub IndataGrid6()
Dim flg, mRowNo As Variant
flg = "N"
For m = MSHFlexGrid9.Rows - 1 To 1 Step -1
    If TmpRs2!LocationName = MSHFlexGrid9.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid9.TextMatrix(m, 3) And mDate = MSHFlexGrid9.TextMatrix(m, 6) Then
       flg = "Y"
       mRowNo = m
       Exit For
    End If
Next

If flg = "N" Then
   mRowNo = MSHFlexGrid9.Rows - 1
   MSHFlexGrid9.TextMatrix(mRowNo, 0) = TmpRs2!StateName
   MSHFlexGrid9.TextMatrix(mRowNo, 1) = TmpRs2!LocationName
   MSHFlexGrid9.TextMatrix(mRowNo, 2) = IIf(IsNull(TmpRs2!EQZone), "", TmpRs2!EQZone)
   MSHFlexGrid9.TextMatrix(mRowNo, 3) = TmpRs2!Commodity
   MSHFlexGrid9.TextMatrix(mRowNo, 6) = mDate
   MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + 1
End If
If TmpRs2!TxnType = "D" Then
   MSHFlexGrid9.TextMatrix(mRowNo, 7) = TmpRs2!NoofBags - mNoOfBags
   MSHFlexGrid9.TextMatrix(mRowNo, 8) = TmpRs2!GrossWt - mGrossWt
Else
   MSHFlexGrid9.TextMatrix(mRowNo, 9) = TmpRs2!NoofBags - mNoOfBags
   MSHFlexGrid9.TextMatrix(mRowNo, 10) = TmpRs2!GrossWt - mGrossWt
End If
MSHFlexGrid9.TextMatrix(mRowNo, 11) = TmpRs2!GainLoss
MSHFlexGrid9.TextMatrix(mRowNo, 4) = Round(Val(MSHFlexGrid9.TextMatrix(mRowNo, 7)) - Val(MSHFlexGrid9.TextMatrix(mRowNo, 9)), 0)
MSHFlexGrid9.TextMatrix(mRowNo, 5) = Round(Val(MSHFlexGrid9.TextMatrix(mRowNo, 8)) - Val(MSHFlexGrid9.TextMatrix(mRowNo, 10)), 3)


'Commented on 12 Sept 2013 while testing of data mismatch for transaction and GSL wise
If Val(MSHFlexGrid9.TextMatrix(mRowNo, 4)) <= 0 Then
   MSHFlexGrid9.TextMatrix(mRowNo, 4) = 0
   MSHFlexGrid9.TextMatrix(mRowNo, 5) = 0
ElseIf Val(MSHFlexGrid9.TextMatrix(mRowNo, 4)) > 0 Then
   If Val(MSHFlexGrid9.TextMatrix(mRowNo, 5)) <= 0 Then
      MSHFlexGrid9.TextMatrix(mRowNo, 5) = 0
   End If
End If

If Val(MSHFlexGrid9.TextMatrix(mRowNo, 5)) < 0 Then
   If Val(MSHFlexGrid9.TextMatrix(mRowNo, 4)) <= 0 Then
      MSHFlexGrid9.TextMatrix(mRowNo, 4) = 0
   End If
   MSHFlexGrid9.TextMatrix(mRowNo, 5) = 0
End If
End Sub
Private Sub IndataGrid4()
Dim flg, mRowNo As Variant
flg = "N"
For m = MSHFlexGrid6.Rows - 1 To 1 Step -1
    If TmpRs2!LocationName = MSHFlexGrid6.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid6.TextMatrix(m, 3) And mDate = MSHFlexGrid6.TextMatrix(m, 6) Then
       flg = "Y"
       mRowNo = m
       Exit For
    End If
Next

If flg = "N" Then
   mRowNo = MSHFlexGrid6.Rows - 1
   MSHFlexGrid6.TextMatrix(mRowNo, 0) = TmpRs2!StateName
   MSHFlexGrid6.TextMatrix(mRowNo, 1) = TmpRs2!LocationName
   MSHFlexGrid6.TextMatrix(mRowNo, 2) = IIf(IsNull(TmpRs2!EQZone), "", TmpRs2!EQZone)
   MSHFlexGrid6.TextMatrix(mRowNo, 3) = TmpRs2!Commodity
   MSHFlexGrid6.TextMatrix(mRowNo, 6) = mDate
   MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + 1
End If
If TmpRs2!TxnType = "D" Then
   MSHFlexGrid6.TextMatrix(mRowNo, 7) = TmpRs2!NoofBags - mNoOfBags
   MSHFlexGrid6.TextMatrix(mRowNo, 8) = TmpRs2!GrossWt - mGrossWt
Else
   MSHFlexGrid6.TextMatrix(mRowNo, 9) = TmpRs2!NoofBags - mNoOfBags
   MSHFlexGrid6.TextMatrix(mRowNo, 10) = TmpRs2!GrossWt - mGrossWt
End If
MSHFlexGrid6.TextMatrix(mRowNo, 11) = TmpRs2!GainLoss
MSHFlexGrid6.TextMatrix(mRowNo, 4) = Round(Val(MSHFlexGrid6.TextMatrix(mRowNo, 7)) - Val(MSHFlexGrid6.TextMatrix(mRowNo, 9)), 0)
MSHFlexGrid6.TextMatrix(mRowNo, 5) = Round(Val(MSHFlexGrid6.TextMatrix(mRowNo, 8)) - Val(MSHFlexGrid6.TextMatrix(mRowNo, 10)), 3)


'Commented on 12 Sept 2013 while testing of data mismatch for transaction and GSL wise
If Val(MSHFlexGrid6.TextMatrix(mRowNo, 4)) <= 0 Then
   MSHFlexGrid6.TextMatrix(mRowNo, 4) = 0
   MSHFlexGrid6.TextMatrix(mRowNo, 5) = 0
ElseIf Val(MSHFlexGrid6.TextMatrix(mRowNo, 4)) > 0 Then
   If Val(MSHFlexGrid6.TextMatrix(mRowNo, 5)) <= 0 Then
      MSHFlexGrid6.TextMatrix(mRowNo, 5) = 0
   End If
End If

If Val(MSHFlexGrid6.TextMatrix(mRowNo, 5)) < 0 Then
   If Val(MSHFlexGrid6.TextMatrix(mRowNo, 4)) <= 0 Then
      MSHFlexGrid6.TextMatrix(mRowNo, 4) = 0
   End If
   MSHFlexGrid6.TextMatrix(mRowNo, 5) = 0
End If

End Sub


Private Sub CalAvgPrice()
Dim mTotAmt As Double
For i = 1 To MSHFlexGrid2.Rows - 2
    mTotAmt = 0
    For j = 0 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(j, 1) = MSHFlexGrid2.TextMatrix(i, 1) And MSHFlexGrid1.TextMatrix(j, 3) = MSHFlexGrid2.TextMatrix(i, 3) Then
           mTotAmt = mTotAmt + Val(MSHFlexGrid1.TextMatrix(j, 8))
        End If
    Next
    MSHFlexGrid2.TextMatrix(i, 8) = Round(mTotAmt)
    MSHFlexGrid2.TextMatrix(i, 9) = Round(Val(MSHFlexGrid2.TextMatrix(i, 8)) / mTotDays, 0)
Next

End Sub

Private Sub CopyGrid3ToGrid1()
For i = 1 To MSHFlexGrid3.Rows - 2
    mr = MSHFlexGrid1.Rows - 1
    For c = 0 To 6
        MSHFlexGrid1.TextMatrix(mr, c) = MSHFlexGrid3.TextMatrix(i, c)
    Next
    MSHFlexGrid1.TextMatrix(mr, 9) = MSHFlexGrid3.TextMatrix(i, 11)
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
End Sub
Private Sub IndataGrid2()
For m = 1 To MSHFlexGrid2.Rows - 1
    If TmpRs2!LocationName = MSHFlexGrid2.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid2.TextMatrix(m, 3) Then
       If TmpRs2!TxnType = "D" Then
          MSHFlexGrid2.TextMatrix(m, 4) = Val(MSHFlexGrid2.TextMatrix(m, 4)) + TmpRs2!NoofBags
          MSHFlexGrid2.TextMatrix(m, 5) = Val(MSHFlexGrid2.TextMatrix(m, 5)) + TmpRs2!GrossWt
       Else
          MSHFlexGrid2.TextMatrix(m, 4) = Val(MSHFlexGrid2.TextMatrix(m, 4)) - TmpRs2!NoofBags
          MSHFlexGrid2.TextMatrix(m, 5) = Val(MSHFlexGrid2.TextMatrix(m, 5)) - TmpRs2!GrossWt
       End If
       If Val(MSHFlexGrid2.TextMatrix(m, 5)) < 0 Then
          MSHFlexGrid2.TextMatrix(m, 5) = 0
       End If
       If Val(MSHFlexGrid2.TextMatrix(m, 4)) < 0 Then
          MSHFlexGrid2.TextMatrix(m, 4) = 0
       End If
       MSHFlexGrid2.TextMatrix(m, 6) = Round(Val(MSHFlexGrid2.TextMatrix(m, 4)) / mTotDays, 0)
       MSHFlexGrid2.TextMatrix(m, 7) = Round(Val(MSHFlexGrid2.TextMatrix(m, 5)) / mTotDays, 3)
       If Val(MSHFlexGrid2.TextMatrix(m, 6)) < 0 Then
          MSHFlexGrid2.TextMatrix(m, 6) = 0
       End If
       If Val(MSHFlexGrid2.TextMatrix(m, 7)) < 0 Then
          MSHFlexGrid2.TextMatrix(m, 7) = 0
       End If
    End If
Next

End Sub
Private Sub IndataGrid1()
Dim flg, mRowNo As Variant
flg = "N"
For m = MSHFlexGrid3.Rows - 1 To 1 Step -1
    If TmpRs2!LocationName = MSHFlexGrid3.TextMatrix(m, 1) And TmpRs2!Commodity = MSHFlexGrid3.TextMatrix(m, 3) And mDate = MSHFlexGrid3.TextMatrix(m, 6) Then
       flg = "Y"
       mRowNo = m
       Exit For
    End If
Next

If flg = "N" Then
   mRowNo = MSHFlexGrid3.Rows - 1
   MSHFlexGrid3.TextMatrix(mRowNo, 0) = TmpRs2!StateName
   MSHFlexGrid3.TextMatrix(mRowNo, 1) = TmpRs2!LocationName
   MSHFlexGrid3.TextMatrix(mRowNo, 2) = IIf(IsNull(TmpRs2!EQZone), "", TmpRs2!EQZone)
   MSHFlexGrid3.TextMatrix(mRowNo, 3) = TmpRs2!Commodity
   MSHFlexGrid3.TextMatrix(mRowNo, 6) = mDate
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + 1
End If
If TmpRs2!TxnType = "D" Then
   MSHFlexGrid3.TextMatrix(mRowNo, 7) = TmpRs2!NoofBags
   MSHFlexGrid3.TextMatrix(mRowNo, 8) = TmpRs2!GrossWt
Else
   MSHFlexGrid3.TextMatrix(mRowNo, 9) = TmpRs2!NoofBags
   MSHFlexGrid3.TextMatrix(mRowNo, 10) = TmpRs2!GrossWt
End If
MSHFlexGrid3.TextMatrix(mRowNo, 11) = TmpRs2!GainLoss
MSHFlexGrid3.TextMatrix(mRowNo, 4) = Round(Val(MSHFlexGrid3.TextMatrix(mRowNo, 7)) - Val(MSHFlexGrid3.TextMatrix(mRowNo, 9)), 0)
MSHFlexGrid3.TextMatrix(mRowNo, 5) = Round(Val(MSHFlexGrid3.TextMatrix(mRowNo, 8)) - Val(MSHFlexGrid3.TextMatrix(mRowNo, 10)), 3)


'Commented on 12 Sept 2013 while testing of data mismatch for transaction and GSL wise
If Val(MSHFlexGrid3.TextMatrix(mRowNo, 4)) <= 0 Then
   MSHFlexGrid3.TextMatrix(mRowNo, 4) = 0
   MSHFlexGrid3.TextMatrix(mRowNo, 5) = 0
ElseIf Val(MSHFlexGrid3.TextMatrix(mRowNo, 4)) > 0 Then
   If Val(MSHFlexGrid3.TextMatrix(mRowNo, 5)) <= 0 Then
      MSHFlexGrid3.TextMatrix(mRowNo, 5) = 0
   End If
End If

If Val(MSHFlexGrid3.TextMatrix(mRowNo, 5)) < 0 Then
   If Val(MSHFlexGrid3.TextMatrix(mRowNo, 4)) <= 0 Then
      MSHFlexGrid3.TextMatrix(mRowNo, 4) = 0
   End If
   MSHFlexGrid3.TextMatrix(mRowNo, 5) = 0
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
Call Grid_Head
Call Grid_Head1
CmdExcel.Enabled = False
End Sub

