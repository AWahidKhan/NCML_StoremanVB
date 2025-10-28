VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_CommodityPrice 
   Caption         =   "Commodity Price"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9255
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   570
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   14775
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
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
         TabIndex        =   13
         Top             =   200
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   14775
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
         Height          =   350
         Left            =   2880
         TabIndex        =   1
         Top             =   240
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
         Height          =   350
         Left            =   10020
         TabIndex        =   7
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
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
         Left            =   4125
         TabIndex        =   2
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Left            =   7650
         TabIndex        =   5
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Left            =   6465
         TabIndex        =   4
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Height          =   350
         Left            =   8835
         TabIndex        =   6
         Top             =   240
         Width           =   1170
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
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
         Height          =   350
         Left            =   5295
         TabIndex        =   3
         Top             =   240
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker TxtMonthYear 
         Height          =   360
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   109707267
         CurrentDate     =   43263
         MinDate         =   38353
      End
      Begin VB.Label Label6 
         Caption         =   "Month-Year"
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
         TabIndex        =   18
         Top             =   288
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
         Left            =   11445
         TabIndex        =   17
         Top             =   285
         Width           =   2295
      End
      Begin VB.Label lblTotalRecords 
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
         Left            =   13845
         TabIndex        =   16
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   14775
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   240
         Width           =   14535
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6015
      Left            =   120
      TabIndex        =   9
      Top             =   3650
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10610
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   375
      Left            =   12360
      TabIndex        =   14
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmMst_CommodityPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim WrkSheet As String
Dim RowCount As Double
Dim FName() As Variant, vTemp As Variant
Dim mMinRate, mMaxRate As Variant
Dim Namef As String
Dim mPriceCell As Boolean

Dim mAllDayPrice As Integer, mPreviousMonthYearDays As Integer

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CommodityPriceUpload.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CommodityPriceUpload.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
        MSHFlexGrid1.row = I
        MSHFlexGrid1.Col = C
                
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
        End If
              
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
On Error Resume Next
oXL.Quit
Unload Me
End Sub

Private Sub CmdGetReport_Click()
Dim mR As Double, mC As Double
Call Grid_Head
If DateDiff("m", CDate(Format(Date, "yyyy-mm-01")), CDate(Format(TxtMonthYear.Value, "yyyy-mm-01"))) > 0 Then
    MsgBox "Future Month-Year data not present."
    Exit Sub
Else
    
''    tmp = " Select A.StateID, B.StateName, A.CommodityID, C.Commodity ,A.MonthYear, A.YearMonth, A.Day1,A.Day2,A.Day3,A.Day4,A.Day5,A.Day6,A.Day7,A.Day8,A.Day9,A.Day10,"
''    tmp = tmp & " A.Day11,A.Day12,A.Day13,A.Day14,A.Day15,A.Day16,A.Day17,A.Day18,A.Day19,A.Day20,"
''    tmp = tmp & " a.Day21 , a.Day22, a.Day23, a.Day24, a.Day25, a.Day26, a.Day27, a.Day28, a.Day29, a.Day30, a.Day31"
''    tmp = tmp & " from Mst_CommodityPrice A"
''    tmp = tmp & " Inner Join Mst_State B On A.StateID = B.StateID"
''    tmp = tmp & " Inner Join Mst_Commodity C On A.CommodityID = C.CommodityID"
''    tmp = tmp & " Where A.YearMonth = '" & Format(TxtMonthYear.Value, "yyyy-mm") & "'"
    
    tmp = " Select A.StateID, B.StateName, A.CommodityID, C.Commodity ,A.MonthYear, A.YearMonth, A.Day1,A.Day2,A.Day3,A.Day4,A.Day5,A.Day6,A.Day7,A.Day8,A.Day9,A.Day10, A.Day11,A.Day12,A.Day13,A.Day14,A.Day15,A.Day16,A.Day17,A.Day18,A.Day19,A.Day20, a.Day21 , a.Day22, a.Day23, a.Day24, a.Day25, a.Day26, a.Day27, a.Day28, a.Day29, a.Day30, a.Day31"
    tmp = tmp & " from Mst_CommodityPrice A"
    tmp = tmp & " Inner Join Mst_State B On A.StateID = B.StateID"
    tmp = tmp & " Inner Join Mst_Commodity C On A.CommodityID = C.CommodityID"
    tmp = tmp & " Where A.YearMonth = '" & Format(TxtMonthYear.Value, "yyyy-mm") & "'"
    tmp = tmp & " Union"
    tmp = tmp & " Select ML.StateID,MS.StateName,GSL.CommodityID,MC.Commodity, Null MonthYear, Null YearMonth,Null Day1, Null Day2,Null Day3,Null Day4,Null Day5,Null Day6,Null Day7,Null Day8,Null Day9,Null Day10,Null Day11,Null Day12,Null Day13,Null Day14,Null Day15,Null Day16,Null Day17,Null Day18,Null Day19,Null Day20,Null Day21,Null Day22,Null Day23,Null Day24,Null Day25,Null Day26,Null Day27,Null Day28,Null Day29,Null Day30,Null Day31"
    tmp = tmp & " From Mst_GSL GSL"
    tmp = tmp & " Inner Join Mst_Location ML On GSL.SM_Prefix = ML.SM_Prefix"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Commodity MC On GSL.CommodityID = MC.CommodityID"
    tmp = tmp & " Where GSL.Status Not In ('B','Z','E') And Convert(Varchar,ML.StateID)+'~'+Convert(Varchar,GSL.CommodityID) NOT IN (Select Convert(Varchar,A.StateID)+'~'+Convert(Varchar,A.CommodityID) from Mst_CommodityPrice A Inner Join Mst_State B On A.StateID = B.StateID Inner Join Mst_Commodity C On A.CommodityID = C.CommodityID Where A.YearMonth = '" & Format(TxtMonthYear.Value, "yyyy-mm") & "')"
    tmp = tmp & " Group By ML.StateID,MS.StateName,GSL.CommodityID,MC.Commodity"
    tmp = tmp & " Order By B.StateName,C.Commodity"
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        ProgressBar1.Max = TmpRs.RecordCount
        With MSHFlexGrid1
        For mR = 1 To TmpRs.RecordCount
                .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
                .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
                .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!MonthYear), "", TmpRs!MonthYear)
                .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!Day1), "", Format(TmpRs!Day1, "#0.00"))
                .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!Day2), "", Format(TmpRs!Day2, "#0.00"))
                .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!Day3), "", Format(TmpRs!Day3, "#0.00"))
                .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!Day4), "", Format(TmpRs!Day4, "#0.00"))
                .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!Day5), "", Format(TmpRs!Day5, "#0.00"))
                .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!Day6), "", Format(TmpRs!Day6, "#0.00"))
                .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!Day7), "", Format(TmpRs!Day7, "#0.00"))
                .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!Day8), "", Format(TmpRs!Day8, "#0.00"))
                .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!Day9), "", Format(TmpRs!Day9, "#0.00"))
                .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!Day10), "", Format(TmpRs!Day10, "#0.00"))
                .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!Day11), "", Format(TmpRs!Day11, "#0.00"))
                .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!Day12), "", Format(TmpRs!Day12, "#0.00"))
                .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!Day13), "", Format(TmpRs!Day13, "#0.00"))
                .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!Day14), "", Format(TmpRs!Day14, "#0.00"))
                .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!Day15), "", Format(TmpRs!Day15, "#0.00"))
                .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!Day16), "", Format(TmpRs!Day16, "#0.00"))
                .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!Day17), "", Format(TmpRs!Day17, "#0.00"))
                .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!Day18), "", Format(TmpRs!Day18, "#0.00"))
                .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!Day19), "", Format(TmpRs!Day19, "#0.00"))
                .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!Day20), "", Format(TmpRs!Day20, "#0.00"))
                .TextMatrix(mR, 23) = IIf(IsNull(TmpRs!Day21), "", Format(TmpRs!Day21, "#0.00"))
                .TextMatrix(mR, 24) = IIf(IsNull(TmpRs!Day22), "", Format(TmpRs!Day22, "#0.00"))
                .TextMatrix(mR, 25) = IIf(IsNull(TmpRs!Day23), "", Format(TmpRs!Day23, "#0.00"))
                .TextMatrix(mR, 26) = IIf(IsNull(TmpRs!Day24), "", Format(TmpRs!Day24, "#0.00"))
                .TextMatrix(mR, 27) = IIf(IsNull(TmpRs!Day25), "", Format(TmpRs!Day25, "#0.00"))
                .TextMatrix(mR, 28) = IIf(IsNull(TmpRs!Day26), "", Format(TmpRs!Day26, "#0.00"))
                .TextMatrix(mR, 29) = IIf(IsNull(TmpRs!Day27), "", Format(TmpRs!Day27, "#0.00"))
                .TextMatrix(mR, 30) = IIf(IsNull(TmpRs!Day28), "", Format(TmpRs!Day28, "#0.00"))
                .TextMatrix(mR, 31) = IIf(IsNull(TmpRs!Day29), "", Format(TmpRs!Day29, "#0.00"))
                .TextMatrix(mR, 32) = IIf(IsNull(TmpRs!Day30), "", Format(TmpRs!Day30, "#0.00"))
                .TextMatrix(mR, 33) = IIf(IsNull(TmpRs!Day31), "", Format(TmpRs!Day31, "#0.00"))
                .Rows = .Rows + 1
                TmpRs.MoveNext
                ProgressBar1.Value = mR
        Next mR
        End With
    End If
End If
LblTotalRecords.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdReset_Click()
ListView1.ListItems.Clear
MSHFlexGrid1.Clear
CmdGetFile.Enabled = True
CmdExit.Enabled = True
CmdImport.Enabled = False
CmdSave.Enabled = False
TxtErrorMsg.Text = ""
ProgressBar1.Value = 0
LblTotalRecords.Caption = ""
Call Grid_Head
TxtMonthYear.Value = Date
End Sub
Private Sub CmdGetFile_Click()
ListView1.ListItems.Clear
TxtErrorMsg.Text = ""

On Error GoTo errortrap
With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx"
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
LblFile.Caption = Path
Call UploadFile
MsgBox "Click Import To import File!!!"
CmdImport.SetFocus
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub UploadFile()

Status = 0
TxtErrorMsg.Text = ""
If LblFile.Caption <> "" Then
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   
   ' Code For Ceating List of Sheet in Excel Sheet
   Dim I As Integer
       For I = 1 To oWB.Worksheets.Count
           Set oWS = oWB.Worksheets(I)
           If I = 1 Then
              ListView1.ListItems.Add I, , oWS.Name
           End If
       Next
   ListView1.ListItems.Item(1).Checked = True
   If Status = 1 Then
      ListView1.ListItems.Clear
      MsgBox ErrMsg, vbExclamation
   End If
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   CmdExit.Enabled = True
   CmdSave.Enabled = False
   oXL.Quit
Else
   MsgBox "Select File", vbInformation, "Information"
End If
End Sub
Private Sub CmdImport_Click()
'Dim mulCheck As Integer
'Dim isSelected As Boolean

LblTotalRecords.Caption = ""

TxtErrorMsg.Text = ""

Call Grid_Head
'For j = 1 To ListView1.ListItems.Count
'    ErrMsg = ListView1.ListItems.Item(j).Text
'    If ListView1.ListItems.Item(j).Checked Then
'       isSelected = True
'       mulCheck = mulCheck + 1
'    End If
'Next j
'If mulCheck > 1 Then
'   MsgBox "You Can Select Only One WorkSheet At A Time!!!"
'   Exit Sub
'End If
'If isSelected = True Then
ErrMsg = ""
Status = 0
RowCount = 0
Set oXL = New Excel.Application
Set oWB = oXL.Workbooks.Open(LblFile.Caption)
Call Validation
If Status = 1 Then
   TxtErrorMsg.Text = ErrMsg
Else
   Call Grid_Head
   Call FillData
   CmdGetFile.Enabled = False
   CmdImport.Enabled = True
   CmdSave.Enabled = True
   CmdReset.Enabled = True
   CmdSave.SetFocus
End If
CmdGetFile.Enabled = True
    
'Else
'    MsgBox "Please Select WorkSheet To Import !...", vbInformation
'End If
End Sub
Private Sub Validation()
For j = 1 To ListView1.ListItems.Count
' Check Column Name In Excel Sheet
    If ListView1.ListItems.Item(j).Checked Then
          Set oWS = oWB.Worksheets(ListView1.ListItems.Item(j).Text)
      
          If LCase(Trim(oWS.Cells(1, 1))) <> "state" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of A " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 2))) <> "commodity" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of B " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 3))) <> "monthyear" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of C " & vbCrLf
             Status = 1
          End If
       
          If LCase(Trim(oWS.Cells(1, 4))) <> "day 1" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of D " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 5))) <> "day 2" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of E " & vbCrLf
             Status = 1
          End If
         
          If LCase(Trim(oWS.Cells(1, 6))) <> "day 3" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of F " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 7))) <> "day 4" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of G " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 8))) <> "day 5" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of H " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 9))) <> "day 6" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of I " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 10))) <> "day 7" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of J " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 11))) <> "day 8" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of K " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 12))) <> "day 9" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of L " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 13))) <> "day 10" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of M " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 14))) <> "day 11" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of N " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 15))) <> "day 12" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of O " & vbCrLf
             Status = 1
          End If
          
          If LCase(Trim(oWS.Cells(1, 16))) <> "day 13" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of p " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 17))) <> "day 14" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of Q " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 18))) <> "day 15" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of R " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 19))) <> "day 16" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of S " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 20))) <> "day 17" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of T " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 21))) <> "day 18" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of U " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 22))) <> "day 19" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of V " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 23))) <> "day 20" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of W " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 24))) <> "day 21" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of X " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 25))) <> "day 22" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of Y " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 26))) <> "day 23" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of Z " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 27))) <> "day 24" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AA " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 28))) <> "day 25" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AB " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 29))) <> "day 26" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AC " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 30))) <> "day 27" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AD " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 31))) <> "day 28" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AE " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 32))) <> "day 29" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AF " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 33))) <> "day 30" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AG " & vbCrLf
             Status = 1
          End If
          If LCase(Trim(oWS.Cells(1, 34))) <> "day 31" Then
             ErrMsg = ErrMsg + "Please Check Sheet " & oWS.Name & " Column Name of AH " & vbCrLf
             Status = 1
          End If
    End If
Next
End Sub
Private Sub FillData()
Dim intNumberRows As Double
Call Grid_Head
For j = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(j).Checked Then
       Set oWS = oWB.Worksheets(ListView1.ListItems.Item(j).Text)
       oWS.ClearArrows
       mRow = 2
       RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
       'RowCount = oWS.Rows.Count
       ProgressBar1.Value = 0
       ProgressBar1.Max = RowCount
       'For L = 1 To RowCount
       Do
         If oWS.Cells(mRow, 1) = "" Then
            Exit Do
         End If
                    
         If oWS.Cells(mRow, 1).Value <> "" Then
            For C = 0 To MSHFlexGrid1.Cols - 1
                If C = 2 Then
                    'MSHFlexGrid1.TextMatrix(mRow - 1, C) = Trim(oWS.Cells(mRow, C + 1).Value)
                    MSHFlexGrid1.TextMatrix(mRow - 1, C) = IIf(IsDate(oWS.Cells(mRow, C + 1).Value) = True, Format(CDate(oWS.Cells(mRow, C + 1).Value), "mmm-yyyy"), "")
                Else
                    MSHFlexGrid1.TextMatrix(mRow - 1, C) = Trim(oWS.Cells(mRow, C + 1).Value)
                End If
                
            Next
            MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
         End If
         ProgressBar1.Value = ProgressBar1.Value + 1
         mRow = mRow + 1
       Loop Until False
       oXL.Quit
       Set oXL = Nothing
    End If
Next
LblTotalRecords.Caption = MSHFlexGrid1.Rows - 2
ProgressBar1.Value = 0
End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 34
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "State"
    .TextMatrix(0, 1) = "Commodity"
    .TextMatrix(0, 2) = "MonthYear"
    .TextMatrix(0, 3) = "Day 1"
    .TextMatrix(0, 4) = "Day 2"
    .TextMatrix(0, 5) = "Day 3"
    .TextMatrix(0, 6) = "Day 4"
    .TextMatrix(0, 7) = "Day 5"
    .TextMatrix(0, 8) = "Day 6"
    .TextMatrix(0, 9) = "Day 7"
    .TextMatrix(0, 10) = "Day 8"
    .TextMatrix(0, 11) = "Day 9"
    .TextMatrix(0, 12) = "Day 10"
    .TextMatrix(0, 13) = "Day 11"
    .TextMatrix(0, 14) = "Day 12"
    .TextMatrix(0, 15) = "Day 13"
    .TextMatrix(0, 16) = "Day 14"
    .TextMatrix(0, 17) = "Day 15"
    .TextMatrix(0, 18) = "Day 16"
    .TextMatrix(0, 19) = "Day 17"
    .TextMatrix(0, 20) = "Day 18"
    .TextMatrix(0, 21) = "Day 19"
    .TextMatrix(0, 22) = "Day 20"
    .TextMatrix(0, 23) = "Day 21"
    .TextMatrix(0, 24) = "Day 22"
    .TextMatrix(0, 25) = "Day 23"
    .TextMatrix(0, 26) = "Day 24"
    .TextMatrix(0, 27) = "Day 25"
    .TextMatrix(0, 28) = "Day 26"
    .TextMatrix(0, 29) = "Day 27"
    .TextMatrix(0, 30) = "Day 28"
    .TextMatrix(0, 31) = "Day 29"
    .TextMatrix(0, 32) = "Day 30"
    .TextMatrix(0, 33) = "Day 31"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 1800
    .ColWidth(2) = 1500
    
    For I = 3 To 33
        .ColWidth(I) = 800
    Next
End With
End Sub


Private Sub CmdSave_Click()
Dim ErrMsg, mCol As Variant
Dim mStateID, mCommodityID, mMYear As Variant
Dim mRow As Double

Call TableMst_CommodityPrice
Call TableMst_State
Call TableMst_Commodity

ErrMsg = ""
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows - 1
mRow = 1

For mRow = 1 To MSHFlexGrid1.Rows - 1
With MSHFlexGrid1
   If .TextMatrix(mRow, 1) = "" Then Exit For
   Status = 0
   wc = ""
   RsMst_State.MoveFirst
   RsMst_Commodity.MoveFirst
   
   ''-- State Name
   If .TextMatrix(mRow, 0) <> "" Then
      RsMst_State.Find ("StateName='" & .TextMatrix(mRow, 0) & "'")
      If RsMst_State.EOF = False Then
         mStateID = RsMst_State!StateID
      Else
        ErrMsg = ErrMsg + "State ID is Not Found of State  " & .TextMatrix(mRow, 0) & "  at row " & mRow + 1 & vbCrLf
        Status = 1
        mCol = 0
        mStateID = ""
        Call ColorRow(mRow, vbRed)
      End If
   Else
      ErrMsg = ErrMsg + "State is blank at row " & mRow + 1 & vbCrLf
      Call ColorRow(mRow, vbRed)
      Status = 1
   End If
   
   ''-- Commodity Name
   If .TextMatrix(mRow, 1) <> "" Then
      RsMst_Commodity.Find ("Commodity='" & .TextMatrix(mRow, 1) & "'")
      If RsMst_Commodity.EOF = False Then
         mMinRate = IIf(IsNull(RsMst_Commodity!MinRate), 0, RsMst_Commodity!MinRate)
         mMaxRate = IIf(IsNull(RsMst_Commodity!MaxRate), 0, RsMst_Commodity!MaxRate)
         mCommodityID = RsMst_Commodity!CommodityID
      Else
         ErrMsg = ErrMsg + "Commodity ID is Not Found of Commodity   " & MSHFlexGrid1.TextMatrix(mRow, 1) & "  at row " & mRow + 1 & vbCrLf
         Status = 1
         mCol = 1
         mMinRate = ""
         mMaxRate = ""
         mCommodityID = ""
         Call ColorRow(mRow, vbRed)
      End If
   Else
      ErrMsg = ErrMsg + "Commodity is blank at row " & mRow + 1 & vbCrLf
      Call ColorRow(mRow, vbRed)
      Status = 1
   End If
   
   ''-- Month-Year Validation
   If .TextMatrix(mRow, 2) <> "" Then
      If IsDate(.TextMatrix(mRow, 2)) = True Then
         'mMonthYear = Split(.TextMatrix(mRow, 2), "-")
        'mMonth = Left(mMonthYear(0), 3) ''mMonthYear(0)
        'mYear = mMonthYear(1)
        
        mMonth = Format(CDate(.TextMatrix(mRow, 2)), "MMM")
        mYear = Format(CDate(.TextMatrix(mRow, 2)), "YYYY")
        
        mMYear = ""
        ''-- Only Current Moth-Year and Past Month-Year are allow to upload
        If DateDiff("m", CDate(Format(Date, "yyyy-mm-01")), CDate(Format(.TextMatrix(mRow, 2), "yyyy-mm-01"))) <= 0 Then
           
          If LCase(mMonth) = "jan" Or LCase(mMonth) = "mar" Or LCase(mMonth) = "may" Or LCase(mMonth) = "jul" Or LCase(mMonth) = "aug" Or LCase(mMonth) = "oct" Or LCase(mMonth) = "dec" Then
             Mday = 31
          Else
             If LCase(mMonth) = "feb" Then
                If (Val(mYear) Mod 4) = 0 Then
                   Mday = 29
                Else
                   Mday = 28
                End If
             ElseIf LCase(mMonth) = "apr" Or LCase(mMonth) = "jun" Or LCase(mMonth) = "sep" Or LCase(mMonth) = "nov" Then
                Mday = 30
             Else
                MsgBox "Data can't loaded due to Invalid month year format (mmm-yyyy) !!!"
                Exit Sub
             End If
          End If
        Else
          ErrMsg = ErrMsg + "Given Month-Year is bigger then Current Month-Year at row " & mRow + 1 & vbCrLf
          Call ColorRow(mRow, vbRed)
          Status = 1
        End If
      Else
        ErrMsg = ErrMsg + "Invalid Month-Year is not allow." & mRow + 1 & vbCrLf
        Call ColorRow(mRow, vbRed)
        Status = 1
      End If
   Else
      ErrMsg = ErrMsg + "MonthYear is blank at row " & mRow + 1 & vbCrLf
      Call ColorRow(mRow, vbRed)
      Status = 1
   End If
   
   ''-- Validation if All Price Cell are Blank then now allow to enter and Update the row
   If Status <> 1 Then
    mPriceCell = False
    For I = 3 To .Cols - (IIf(Mday = 31, 1, IIf(Mday = 30, 2, IIf(Mday = 29, 3, IIf(Mday = 28, 4, 0)))))
        If Trim(.TextMatrix(mRow, I)) <> "" Then
            mPriceCell = True
        End If
    Next
    If mPriceCell = False Then
       ErrMsg = ErrMsg + "Price is not define in any one of the Day of define Month-Year Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
       Call ColorRow(mRow, vbRed)
       Status = 1
    End If
   End If
   
   ''-- Validation to check the previous Month-Year all days Data should be uploded
   If Status <> 1 Then
     tmp = " Select * from Mst_CommodityPrice Where StateID = " & mStateID & " And CommodityID = " & mCommodityID & " And YearMonth = (Select Max(YearMonth) from Mst_CommodityPrice Where StateID = " & mStateID & " And CommodityID = " & mCommodityID & ") "
     Call TmpTable
     If TmpRs.RecordCount > 0 Then
        mPriceCell = True
        If DateDiff("m", CDate(Format(.TextMatrix(mRow, 2), "yyyy-mm-01")), CDate(Format(TmpRs!YearMonth & "-01", "yyyy-mm-dd"))) < 0 Then
            mPreviousMonthYearDays = 0
            mPreviousMonthYearDays = NoofDayInMonth(Format(TmpRs!YearMonth & "-01", "mmm"), Format(TmpRs!YearMonth & "-01", "yyyy"))
            If mPreviousMonthYearDays > 0 Then
                
                If mPreviousMonthYearDays = 28 Then
                    If IsNull(TmpRs!Day1 + TmpRs!Day2 + TmpRs!Day3 + TmpRs!Day4 + TmpRs!Day5 + TmpRs!Day6 + TmpRs!Day7 + TmpRs!Day8 + TmpRs!Day9 + TmpRs!Day10 + TmpRs!Day11 + TmpRs!Day12 + TmpRs!Day13 + TmpRs!Day14 + TmpRs!Day15 + TmpRs!Day16 + TmpRs!Day17 + TmpRs!Day18 + TmpRs!Day19 + TmpRs!Day20 + TmpRs!Day21 + TmpRs!Day22 + TmpRs!Day23 + TmpRs!Day24 + TmpRs!Day25 + TmpRs!Day26 + TmpRs!Day27 + TmpRs!Day28) = True Then
                        mPriceCell = False
                    End If
                ElseIf mPreviousMonthYearDays = 29 Then
                    If IsNull(TmpRs!Day1 + TmpRs!Day2 + TmpRs!Day3 + TmpRs!Day4 + TmpRs!Day5 + TmpRs!Day6 + TmpRs!Day7 + TmpRs!Day8 + TmpRs!Day9 + TmpRs!Day10 + TmpRs!Day11 + TmpRs!Day12 + TmpRs!Day13 + TmpRs!Day14 + TmpRs!Day15 + TmpRs!Day16 + TmpRs!Day17 + TmpRs!Day18 + TmpRs!Day19 + TmpRs!Day20 + TmpRs!Day21 + TmpRs!Day22 + TmpRs!Day23 + TmpRs!Day24 + TmpRs!Day25 + TmpRs!Day26 + TmpRs!Day27 + TmpRs!Day28 + TmpRs!Day29) = True Then
                        mPriceCell = False
                    End If
                ElseIf mPreviousMonthYearDays = 30 Then
                    If IsNull(TmpRs!Day1 + TmpRs!Day2 + TmpRs!Day3 + TmpRs!Day4 + TmpRs!Day5 + TmpRs!Day6 + TmpRs!Day7 + TmpRs!Day8 + TmpRs!Day9 + TmpRs!Day10 + TmpRs!Day11 + TmpRs!Day12 + TmpRs!Day13 + TmpRs!Day14 + TmpRs!Day15 + TmpRs!Day16 + TmpRs!Day17 + TmpRs!Day18 + TmpRs!Day19 + TmpRs!Day20 + TmpRs!Day21 + TmpRs!Day22 + TmpRs!Day23 + TmpRs!Day24 + TmpRs!Day25 + TmpRs!Day26 + TmpRs!Day27 + TmpRs!Day28 + TmpRs!Day29 + TmpRs!Day30) = True Then
                        mPriceCell = False
                    End If
                ElseIf mPreviousMonthYearDays = 31 Then
                    If IsNull(TmpRs!Day1 + TmpRs!Day2 + TmpRs!Day3 + TmpRs!Day4 + TmpRs!Day5 + TmpRs!Day6 + TmpRs!Day7 + TmpRs!Day8 + TmpRs!Day9 + TmpRs!Day10 + TmpRs!Day11 + TmpRs!Day12 + TmpRs!Day13 + TmpRs!Day14 + TmpRs!Day15 + TmpRs!Day16 + TmpRs!Day17 + TmpRs!Day18 + TmpRs!Day19 + TmpRs!Day20 + TmpRs!Day21 + TmpRs!Day22 + TmpRs!Day23 + TmpRs!Day24 + TmpRs!Day25 + TmpRs!Day26 + TmpRs!Day27 + TmpRs!Day28 + TmpRs!Day29 + TmpRs!Day30 + TmpRs!Day31) = True Then
                        mPriceCell = False
                    End If
                End If
            Else
                ErrMsg = ErrMsg + "invalid Month-Year. Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                Call ColorRow(mRow, vbRed)
                Status = 1
            End If
        End If
                
        If mPriceCell = False Then
           ErrMsg = ErrMsg + "All Prices of " & Format(TmpRs!YearMonth & "-01", "mmm-yyyy") & " are not uploaded Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
           Call ColorRow(mRow, vbRed)
           Status = 1
        End If
        
     End If
   End If
      
   If Status <> 1 Then
    RsMst_CommodityPrice.Filter = ""
    RsMst_CommodityPrice.MoveFirst
    RsMst_CommodityPrice.Filter = " StateID = " & mStateID & " And CommodityID = " & mCommodityID & " And YearMonth = '" & Format(.TextMatrix(mRow, 2), "yyyy-mm") & "'"
    ''-- Find the Row Data for State -Commodity and Year Month
    ''-- if data is available then update the only Null Value Day but dont update the exists value if it is reapted for same day
    ''-- if data is not available then add the new row and insert the data
        Dim mStatus As Variant
        If RsMst_CommodityPrice.EOF = False Then
             
            ''For I = .Cols - 31 To Mday + 2
            For I = 3 To .Cols - (IIf(Mday = 31, 1, IIf(Mday = 30, 2, IIf(Mday = 29, 3, IIf(Mday = 28, 4, 0)))))
                If (I - 2) <= Mday Then
                    ''-- Only Back date data or Current Date Data is allow to update.
                    If DateDiff("d", CDate(Format(Date, "yyyy-mm-dd")), CDate(Format(.TextMatrix(mRow, 2), "yyyy-mm-" & (I - 2) & ""))) <= 0 Then
                        ''-- Only Non blank cells are allow to update and add
                        If Trim(.TextMatrix(mRow, I)) <> "" Then
                            ''-- Zero Value are not allow to Update or add in new One
                            If Val(.TextMatrix(mRow, I)) > 0 Then
                                ''-- Check for Min and Max Range of Commodity
                                If Val(.TextMatrix(mRow, I)) <= mMaxRate And Val(.TextMatrix(mRow, I)) >= mMinRate Then
                                   ''--Check that previous Days are are already uploaded
                                   If I > 3 Then
                                      If IsNull(RsMst_CommodityPrice.Fields(I - 1)) = True Then
                                        ErrMsg = ErrMsg + "Commodity Price of previous days are not uploaded Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                                        Call ColorRow(mRow, vbRed)
                                        RsMst_CommodityPrice.CancelUpdate
                                        mStatus = 0
                                        Exit For
                                      End If
                                   End If
                                
                                   ''-- Update the Price if Availble
                                   If IsNull(RsMst_CommodityPrice.Fields(I)) = True Then
                                    RsMst_CommodityPrice.Fields(I) = Val(IIf(.TextMatrix(mRow, I) = "", 0, Round(.TextMatrix(mRow, I), 2)))
                                    mStatus = 1
                                   Else
                                    RsMst_CommodityPrice.Fields(I) = Val(IIf(.TextMatrix(mRow, I) = "", 0, Round(.TextMatrix(mRow, I), 2)))
                                    Call ColorCell(mRow, CDbl(I), vbGreen)
                                    mStatus = 1
                                   End If
                                Else
                                   ErrMsg = ErrMsg + "Commodity Price is not between " & mMinRate & " And " & mMaxRate & " Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                                   Call ColorRow(mRow, vbRed)
                                   RsMst_CommodityPrice.CancelUpdate
                                   mStatus = 0
                                   Exit For
                                End If
                            Else
                                ErrMsg = ErrMsg + "Commodity Price Can not be Zero." & " Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                                Call ColorRow(mRow, vbRed)
                                RsMst_CommodityPrice.CancelUpdate
                                mStatus = 0
                                Exit For
                            End If
                        End If
                    Else
                        If Trim(.TextMatrix(mRow, I)) <> "" Then
                            ErrMsg = ErrMsg + "Commodity Price is not allow to update for future Date." & " Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                            Call ColorRow(mRow, vbRed)
                            RsMst_CommodityPrice.CancelUpdate
                            mStatus = 0
                            Exit For
                        End If
                    End If
                Else
                    ErrMsg = ErrMsg + "Given Month Price is more then no of Days in month Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                    Call ColorRow(mRow, vbRed)
                    RsMst_CommodityPrice.CancelUpdate
                    mStatus = 0
                    Exit For
                End If
                
            Next
            
            If mStatus <> 0 Then
                If IsNull(RsMst_CommodityPrice!CreatedBy) = True Or RsMst_CommodityPrice!CreatedBy = "" Then
                   RsMst_CommodityPrice!CreatedBy = Gen_UserLoginID
                   RsMst_CommodityPrice!CreatedDate = Now
                Else
                   RsMst_CommodityPrice!UpdatedBy = Gen_UserLoginID
                   RsMst_CommodityPrice!UpdatedDate = Now
                End If
                RsMst_CommodityPrice.Update
                RsMst_CommodityPrice.Requery
            End If
        
        Else
            
            RsMst_CommodityPrice.AddNew
            RsMst_CommodityPrice!StateID = mStateID
            RsMst_CommodityPrice!CommodityID = mCommodityID
            RsMst_CommodityPrice!MonthYear = Format(.TextMatrix(mRow, 2), "mmm-yyyy")
            RsMst_CommodityPrice!YearMonth = Format(.TextMatrix(mRow, 2), "yyyy-mm")
            RsMst_CommodityPrice!G_UID = GetGUID
            
            ''For I = .Cols - 31 To Mday + 2
            For I = 3 To .Cols - (IIf(Mday = 31, 1, IIf(Mday = 30, 2, IIf(Mday = 29, 3, IIf(Mday = 28, 4, 0)))))
                
                If (I - 2) <= Mday Then
                    ''-- Only Back date data or Current Date Data is allow to update.
                    If DateDiff("d", CDate(Format(Date, "yyyy-mm-dd")), CDate(Format(.TextMatrix(mRow, 2), "yyyy-mm-" & (I - 2) & ""))) <= 0 Then
                        ''-- Only Non blank cells are allow to update and add
                        If Trim(.TextMatrix(mRow, I)) <> "" Then
                            ''-- Zero Value are not allow to Update or add in new One
                            If Val(.TextMatrix(mRow, I)) > 0 Then
                                ''-- Check for Min and Max Range of Commodity
                                If Val(.TextMatrix(mRow, I)) <= mMaxRate And Val(.TextMatrix(mRow, I)) >= mMinRate Then
                                   ''--Check that previous Days are are already uploaded
                                   If I > 3 Then
                                      If IsNull(RsMst_CommodityPrice.Fields(I - 1)) = True Then
                                        ErrMsg = ErrMsg + "Commodity Price of previous days are not uploaded Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                                        Call ColorRow(mRow, vbRed)
                                        RsMst_CommodityPrice.CancelUpdate
                                        mStatus = 0
                                        Exit For
                                      End If
                                   End If
                                   
                                   ''-- Update the Price if Availble
                                   If IsNull(RsMst_CommodityPrice.Fields(I)) = True Then
                                      RsMst_CommodityPrice.Fields(I) = Val(IIf(.TextMatrix(mRow, I) = "", 0, Round(.TextMatrix(mRow, I), 2)))
                                      mStatus = 1
                                   Else
                                      RsMst_CommodityPrice.Fields(I) = Val(IIf(.TextMatrix(mRow, I) = "", 0, Round(.TextMatrix(mRow, I), 2)))
                                      Call ColorCell(mRow, CDbl(I), vbGreen)
                                      mStatus = 1
                                   End If
                                Else
                                   ErrMsg = ErrMsg + "Commodity Price is not between " & mMinRate & " And " & mMaxRate & " Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                                   Call ColorRow(mRow, vbRed)
                                   RsMst_CommodityPrice.CancelUpdate
                                   mStatus = 0
                                   Exit For
                                End If
                            Else
                                ErrMsg = ErrMsg + "Commodity Price Can not be Zero." & " Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                                Call ColorRow(mRow, vbRed)
                                RsMst_CommodityPrice.CancelUpdate
                                mStatus = 0
                                Exit For
                            End If
                        End If
                Else
                    If Trim(.TextMatrix(mRow, I)) <> "" Then
                        ErrMsg = ErrMsg + "Commodity Price is not allow to update for future Date." & " Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                        Call ColorRow(mRow, vbRed)
                        RsMst_CommodityPrice.CancelUpdate
                        mStatus = 0
                        Exit For
                    End If
                End If
                
                Else
                    ErrMsg = ErrMsg + "Given Month Price is more then no of Days in month Where State=" & .TextMatrix(mRow, 0) & " And Commodity=" & .TextMatrix(mRow, 1) & vbCrLf
                    Call ColorRow(mRow, vbRed)
                    RsMst_CommodityPrice.CancelUpdate
                    mStatus = 0
                    Exit For
                End If
                
            Next

            If mStatus <> 0 Then
                If IsNull(RsMst_CommodityPrice!CreatedBy) = True Or RsMst_CommodityPrice!CreatedBy = "" Then
                   RsMst_CommodityPrice!CreatedBy = Gen_UserLoginID
                   RsMst_CommodityPrice!CreatedDate = Now
                Else
                   RsMst_CommodityPrice!UpdatedBy = Gen_UserLoginID
                   RsMst_CommodityPrice!UpdatedDate = Now
                End If
                RsMst_CommodityPrice.Update
                RsMst_CommodityPrice.Requery
            End If
        
        End If
   End If
   
End With
ProgressBar1.Value = ProgressBar1.Value + 1
Next mRow

ProgressBar1.Value = 0
If ErrMsg = "" Then
   MsgBox "Data Saved To Database Successfully !!!", vbInformation
Else
   TxtErrorMsg.Text = "Following Data is Not Inserted into Database Due To Error !..." & vbCrLf & ErrMsg
   MsgBox "Data Saved To Database Successfully Except above error List.", vbExclamation
End If

CmdSave.Enabled = False
CmdGetFile.Enabled = True
CmdExit.Enabled = True

End Sub


''Private Sub CmdSave_Click()
''Dim ErrMsg, mCol As Variant
''
''Call TableMst_CommodityPrice
''Call TableMst_State
''Call TableMst_Commodity
''
''ErrMsg = ""
''ProgressBar1.Value = 0
''ProgressBar1.Max = MSHFlexGrid1.Rows - 1
''mRow = 1
''For mRow = 1 To MSHFlexGrid1.Rows - 1
''   If MSHFlexGrid1.TextMatrix(mRow, 1) = "" Then
''      Exit For
''   End If
''   Status = 0
''   wc = ""
''   RsMst_State.MoveFirst
''   RsMst_Commodity.MoveFirst
''
''   RsMst_CommodityPrice.AddNew
''
''   ''-- State Name
''   If MSHFlexGrid1.TextMatrix(mRow, 0) <> "" Then
''      RsMst_State.Find ("StateName='" & MSHFlexGrid1.TextMatrix(mRow, 0) & "'")
''      If RsMst_State.EOF = False Then
''         RsMst_CommodityPrice!StateID = RsMst_State!StateID
''         wc = " StateID=" & RsMst_CommodityPrice!StateID & ""
''      Else
''        ErrMsg = ErrMsg + "State ID is Not Found of State  " & MSHFlexGrid1.TextMatrix(mRow, 0) & "  at row " & mRow + 1 & vbCrLf
''        Status = 1
''        mCol = 0
''      End If
''   Else
''      ErrMsg = ErrMsg + "State is blank at row " & mRow + 1 & vbCrLf
''      Status = 1
''   End If
''
''   ''-- Commodity Name
''   If MSHFlexGrid1.TextMatrix(mRow, 1) <> "" Then
''      RsMst_Commodity.Find ("Commodity='" & MSHFlexGrid1.TextMatrix(mRow, 1) & "'")
''      If RsMst_Commodity.EOF = False Then
''         RsMst_CommodityPrice!CommodityID = RsMst_Commodity!CommodityID
''         mMinRate = IIf(IsNull(RsMst_Commodity!MinRate), 0, RsMst_Commodity!MinRate)
''         mMaxRate = IIf(IsNull(RsMst_Commodity!MaxRate), 0, RsMst_Commodity!MaxRate)
''         wc = wc + " and CommodityID=" & RsMst_Commodity!CommodityID & ""
''      Else
''         ErrMsg = ErrMsg + "Commodity ID is Not Found of Commodity   " & MSHFlexGrid1.TextMatrix(mRow, 1) & "  at row " & mRow + 1 & vbCrLf
''         Status = 1
''         mCol = 1
''         mMinRate = ""
''         mMaxRate = ""
''      End If
''   Else
''      ErrMsg = ErrMsg + "Commodity is blank at row " & mRow + 1 & vbCrLf
''      Status = 1
''   End If
''
''   If MSHFlexGrid1.TextMatrix(mRow, 2) <> "" Then
''      RsMst_CommodityPrice!MonthYear = MSHFlexGrid1.TextMatrix(mRow, 2)
''      RsMst_CommodityPrice!YearMonth = Format(MSHFlexGrid1.TextMatrix(mRow, 2), "yyyy-mm")
''      mMonthYear = Split(MSHFlexGrid1.TextMatrix(mRow, 2), "-")
''      mMonth = mMonthYear(0)
''      mYear = mMonthYear(1)
''      If LCase(mMonth) = "jan" Or LCase(mMonth) = "mar" Or LCase(mMonth) = "may" Or LCase(mMonth) = "jul" Or LCase(mMonth) = "aug" Or LCase(mMonth) = "oct" Or LCase(mMonth) = "dec" Then
''         Mday = 31
''      Else
''         If LCase(mMonth) = "feb" Then
''            If (Val(mYear) Mod 4) = 0 Then
''               Mday = 28
''            Else
''               Mday = 29
''            End If
''         ElseIf LCase(mMonth) = "apr" Or LCase(mMonth) = "jun" Or LCase(mMonth) = "sep" Or LCase(mMonth) = "nov" Then
''            Mday = 30
''         Else
''            MsgBox "Data can't loaded due to Invalid month year format (mmm-yyyy) !!!"
''            Exit Sub
''         End If
''      End If
''      'wc = wc + " and MonthYear='" & RsMst_CommodityPrice!MonthYear & "'"
''      wc = wc + " and YearMonth='" & RsMst_CommodityPrice!YearMonth & "'"
''
''   Else
''      ErrMsg = ErrMsg + "MonthYear is blank at row " & mRow + 1 & vbCrLf
''      Status = 1
''   End If
''
''   If Status <> 1 Then
''      tmp = "Select * from Mst_CommodityPrice where " + wc
''      Call TmpTable
''      Dim mStatus As Variant
''      If TmpRs.RecordCount = 0 Then
''         For I = MSHFlexGrid1.Cols - 31 To Mday + 2 'MSHFlexGrid1.Cols - 1
''             If Val(MSHFlexGrid1.TextMatrix(mRow, I)) <= mMaxRate And Val(MSHFlexGrid1.TextMatrix(mRow, I)) >= mMinRate Or mMinRate = 0 Then
''                RsMst_CommodityPrice.Fields(I) = Val(IIf(MSHFlexGrid1.TextMatrix(mRow, I) = "", 0, Round(MSHFlexGrid1.TextMatrix(mRow, I), 2)))
''                mStatus = 1
''             Else
''                ErrMsg = ErrMsg + "Commodity Price is not between " & mMinRate & " And " & mMaxRate & " Where State=" & MSHFlexGrid1.TextMatrix(mRow, 0) & " And Commodity=" & MSHFlexGrid1.TextMatrix(mRow, 1) & vbCrLf
''                MSHFlexGrid1.row = mRow
''                MSHFlexGrid1.Col = 0
''                For y = 0 To MSHFlexGrid1.Cols - 1
''                    MSHFlexGrid1.Col = y
''                    MSHFlexGrid1.CellBackColor = vbRed
''                Next
''                RsMst_CommodityPrice.CancelUpdate
''                mStatus = 0
''                Exit For
''             End If
''         Next
''         If mStatus <> 0 Then
''            RsMst_CommodityPrice.Update
''         End If
''      Else
''''         MSHFlexGrid1.Row = mRow
''''         MSHFlexGrid1.Col = 0
''''         For y = 0 To MSHFlexGrid1.Cols - 1
''''             MSHFlexGrid1.Col = y
''''             MSHFlexGrid1.CellBackColor = vbRed
''''         Next
''''
''''
''''         ErrMsg = ErrMsg + "Data Already Present of Row " & mRow + 1 & "  Where State=" & MSHFlexGrid1.TextMatrix(mRow, 0) & " And Commodity=" & MSHFlexGrid1.TextMatrix(mRow, 1) & vbCrLf
''''         For I = MSHFlexGrid1.Cols - 31 To Mday + 2 'MSHFlexGrid1.Cols - 1
''''             TmpRs.Fields(I) = Val(IIf(MSHFlexGrid1.TextMatrix(mRow, I) = "", 0, Round(MSHFlexGrid1.TextMatrix(mRow, I), 2)))
''''         Next
''''         'TmpRs!YearMonth = a
''''         TmpRs.Update
''''         'RsMst_CommodityPrice.CancelUpdate
''
''         MSHFlexGrid1.row = mRow
''         MSHFlexGrid1.Col = 0
''         For y = 0 To MSHFlexGrid1.Cols - 1
''             MSHFlexGrid1.Col = y
''             MSHFlexGrid1.CellBackColor = vbGreen
''         Next
''         ErrMsg = ErrMsg + "Data Already Present of Row " & mRow + 1 & " Where State=" & MSHFlexGrid1.TextMatrix(mRow, 0) & " And Commodity=" & MSHFlexGrid1.TextMatrix(mRow, 1) & " And MonthYear=" & MSHFlexGrid1.TextMatrix(mRow, 2) & vbCrLf
''         RsMst_CommodityPrice.CancelUpdate
''
''
''
''      End If
''   Else
''      MSHFlexGrid1.row = mRow
''      MSHFlexGrid1.Col = mCol
''      For y = 0 To MSHFlexGrid1.Cols - 1
''        MSHFlexGrid1.Col = y
''        MSHFlexGrid1.CellBackColor = vbRed
''      Next
''      RsMst_CommodityPrice.CancelUpdate
''   End If
''   ProgressBar1.Value = ProgressBar1.Value + 1
''Next
''
''ProgressBar1.Value = 0
''If ErrMsg = "" Then
''   MsgBox "Data Saved To Database Successfully !!!", vbInformation
''Else
''   TxtErrorMsg.Text = "Following Data is Not Inserted  into Database Due To Error !..." & vbCrLf & ErrMsg
''   MsgBox "Data Saved To Database Successfully Except above error List", vbExclamation
''End If
''CmdSave.Enabled = False
''CmdGetFile.Enabled = True
''CmdExit.Enabled = True
''End Sub

Private Sub Form_Load()
LblFile.Caption = ""
Frame1.Enabled = False
Frame2.Enabled = False
CmdSave.Enabled = False
ProgressBar1.Value = 0

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
CmdSave.Enabled = False
CmdImport.Enabled = False
CmdExit.Enabled = False
CmdReset.Enabled = False
TxtMonthYear.Value = Date
Call Grid_Head
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Index <> 1 Then
   Item.Checked = False
End If
End Sub

Private Sub ColorRow(mRowIndex As Double, mColor As ColorConstants)
Dim mR As Double, mC As Double
With MSHFlexGrid1
    .row = mRowIndex
    For mC = 0 To .Cols - 1
        .Col = mC
        .CellBackColor = mColor
    Next mC
End With
End Sub

Private Sub ColorCell(mRowIndex As Double, mColIndex As Double, mColor As ColorConstants)
Dim mR As Double, mC As Double
With MSHFlexGrid1
    .row = mRowIndex
    .Col = mColIndex
    .CellBackColor = mColor
End With
End Sub

Private Function NoofDayInMonth(mMonth_ As Variant, myear_ As Variant) As Integer
On Error GoTo ErrorMsg
Dim mMonth As Integer, mYear As Integer
Dim Mday As Integer

mMonth = Month("01-" & mMonth_ & "-" & myear_) 'mYear)
mYear = Year("01-" & mMonth_ & "-" & myear_) 'mYear)

If mMonth = 1 Or mMonth = 3 Or mMonth = 5 Or mMonth = 7 Or mMonth = 8 Or mMonth = 10 Or mMonth = 12 Then
   Mday = 31
Else
   If mMonth = 2 Then
      If (Val(mYear) Mod 4) = 0 Then
         Mday = 29
      Else
         Mday = 28
      End If
   ElseIf mMonth = 4 Or mMonth = 6 Or mMonth = 9 Or mMonth = 11 Then
      Mday = 30
   Else
      Mday = 0
   End If
End If

ErrorMsg:
    If Err.Number <> 0 Then
        Mday = 0
    End If
NoofDayInMonth = Mday
End Function

