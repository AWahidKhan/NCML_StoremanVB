VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMst_StockStatus 
   Caption         =   "Stock Status"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1020
      Left            =   10275
      TabIndex        =   15
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox CmbMonthYear 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Month And Year"
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
         TabIndex        =   16
         Top             =   420
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Error Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   15000
      Begin VB.TextBox TxtErrorMsg 
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   240
         Width           =   14535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton CmdTotalCapacityUpdate 
         Caption         =   "Total Capacity Update"
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
         Height          =   465
         Left            =   4440
         TabIndex        =   4
         Top             =   165
         Width           =   2895
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Height          =   465
         Left            =   2985
         TabIndex        =   3
         Top             =   165
         Width           =   1455
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
         Height          =   465
         Left            =   1530
         TabIndex        =   2
         Top             =   165
         Width           =   1455
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8790
         TabIndex        =   6
         Top             =   165
         Width           =   1215
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Height          =   465
         Left            =   7335
         TabIndex        =   5
         Top             =   165
         Width           =   1455
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
         Height          =   465
         Left            =   75
         TabIndex        =   1
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   690
         Width           =   705
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   11033
      _Version        =   393216
      WordWrap        =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   9360
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   7800
      TabIndex        =   8
      Top             =   0
      Width           =   135
      Begin VB.ComboBox CmbSheetName 
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
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   4740
      End
      Begin VB.Label Label1 
         Caption         =   "Select Sheet Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   10
         Top             =   360
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmMst_StockStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim mSheetFlg As Variant

Private Sub CmdGetFile_Click()

On Error GoTo errortrap
With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls)|*.xls"
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
lblFile.Caption = Path

CmbSheetName.Clear
If lblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(lblFile.Caption)
 ' Code For Creating List of Sheet in Excel Sheet
   Dim I As Integer
   Dim runApp As Boolean
   runApp = False
   For I = 1 To oWB.Worksheets.Count
       Set oWS = oWB.Worksheets(I)
       Status = 1
       CmbSheetName.AddItem oWS.Name
       If LCase(oWS.Name) = "rpt_txn_summary" Then
          runApp = True
       End If
   Next
   If runApp = False Then
      MsgBox "Please Select Proper Excel Sheet!!!"
      Exit Sub
   End If
   
   CmbSheetName.SelText = "Rpt_Txn_Summary"
   Frame2.Enabled = True
   MsgBox "Please Select Year And Month !!!"
   CmbSheetName.SetFocus
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   CmbMonthYear.Enabled = True
   CmbMonthYear.SetFocus
   oXL.Quit
End If

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub
Private Sub CmbSheetName_LostFocus()
If CmbSheetName.Text = "" Then
   mSheetFlg = Null
   Exit Sub
End If
Call Grid_Head

End Sub
Private Sub CmdExit_Click()
'If oXL.ActiveWorkbook = True Then oXL.Quit
Unload Me
End Sub

Private Sub CmdImport_Click()
If CmbSheetName.Text = "" Then
   MsgBox "Name Of Sheet Can't be blank!!!"
   CmbSheetName.SetFocus
   Exit Sub
End If
If CmbMonthYear.Text = "" Then
   MsgBox "Please Enter Year Month !!!"
   CmbMonthYear.SetFocus
   Exit Sub
End If

'-----------------------------
If CmbSheetName.Text <> "" Then
   MsgBox "Wait till next message!!!"
   Call Grid_Head
   TxtErrorMsg = ""
   ErrMsg = ""
   Status = 0
   RowCount = 0
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(lblFile.Caption)
   Call Grid_Head
   Call FillData
   CmdSave.Enabled = True
   CmdTotalCapacityUpdate.Enabled = True
   CmdReset.Enabled = True
   CmdSave.SetFocus
   MsgBox "Import Succesful!!!"
   
End If

End Sub

Private Sub CmdReset_Click()

Call Grid_Head
'CmdImport.Enabled = True
CmdExit.Enabled = True
CmdSave.Enabled = False
CmdTotalCapacityUpdate.Enabled = True
CmdGetfile.Enabled = True
If lblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
lblFile.Caption = ""
mSheetFlg = ""
TxtErrorMsg = ""
CmbSheetName.Text = ""

ProgressBar1.Value = 0

End Sub

Private Sub CmdSave_Click()
Dim ErrMsg As Variant

CmdSave.Enabled = False
If MSHFlexGrid1.Rows < 3 Then
   MsgBox "No data to save!!!"
   Exit Sub
End If

If MSHFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
MsgBox "Wait till Next Message....."
Call TableMst_StockStatus
Call TableMst_State
Call TableMst_Location
Call TableMst_CMP
Call TableMst_UnitType
Call TableMst_Commodity
Call TableMst_TmpStockStatus

Dim mStateID, mLocationID, mCMPID, mUnitTypeID, mCommodityID As Variant

ErrMsg = ""
mRow = 0
  
RowCount = MSHFlexGrid1.Rows  'oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row

ProgressBar1.Value = 0
ProgressBar1.Max = RowCount


tmp = "select locationid ,min(cmpID) As CMPID from mst_cmp where closedate is null group by locationid"
Call Tmp4Table
For c = 1 To RowCount - 2
    Status = 0
    Wc = ""
    
    mLocationID = GetLocationID(Trim(MSHFlexGrid1.TextMatrix(c, 2)))
    
    TmpRs4.MoveFirst
    TmpRs4.Find "locationid = " & mLocationID
    If Not TmpRs4.EOF Then
       mCMPID = TmpRs4!CMPID
    Else
       mCMPID = 0
    End If
    
    If mLocationID <> 0 And mCMPID <> 0 Then
       RsMst_TmpStockStatus.AddNew
       RsMst_TmpStockStatus!MonthYear = MSHFlexGrid1.TextMatrix(c, 0)
       RsMst_TmpStockStatus!LocationID = mLocationID
       RsMst_State.MoveFirst
       RsMst_State.Find ("StateName='" & MSHFlexGrid1.TextMatrix(c, 1) & "'")
       If RsMst_State.EOF = False Then
          mStateID = RsMst_State!StateID
          RsMst_TmpStockStatus!StateID = mStateID
       End If
       RsMst_TmpStockStatus!UnitTypeID = GetUnitID(MSHFlexGrid1.TextMatrix(c, 4))
       RsMst_TmpStockStatus!ClosingBags = Format(MSHFlexGrid1.TextMatrix(c, 6), "#00000000.00")
       RsMst_TmpStockStatus!ClosingWeight = Format(MSHFlexGrid1.TextMatrix(c, 7), "#00000000.00")
       RsMst_TmpStockStatus!CMPID = mCMPID
       
       RsMst_TmpStockStatus.Update
       
    ElseIf mCMPID = 0 Then  'Location
       ErrMsg = ErrMsg & "CMP " & MSHFlexGrid1.TextMatrix(c, 3) & " Not Found of Row  " & c + 1 & vbCrLf
       MSHFlexGrid1.Col = 3
       MSHFlexGrid1.Row = c
       MSHFlexGrid1.CellBackColor = vbRed
    Else
       ErrMsg = ErrMsg & "Location " & MSHFlexGrid1.TextMatrix(c, 2) & " Not Found of Row  " & c + 1 & vbCrLf
       MSHFlexGrid1.Col = 2
       MSHFlexGrid1.Row = c
       MSHFlexGrid1.CellBackColor = vbRed
    End If

   ProgressBar1.Value = ProgressBar1.Value + 1
Next

tmp = "Select MonthYear,StateID,LocationID,UnitTypeID,CMPID,Sum(ClosingBags) 'ClosingBags',Sum(ClosingWeight)'ClosingWeight'"
tmp = tmp & " From Mst_TmpStockStatus"
tmp = tmp & " Group by LocationID,MonthYear,StateID,UnitTypeID,CMPID"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       If RsMst_StockStatus.RecordCount > 0 Then
          RsMst_StockStatus.MoveFirst
          RsMst_StockStatus.Filter = "MonthYear= '" & TmpRs!MonthYear & "' and StateID= " & TmpRs!StateID & " and LocationID= " & TmpRs!LocationID & " and UnitTypeID= " & TmpRs!UnitTypeID & " and ClosingBags= " & TmpRs!ClosingBags & " and ClosingWeight=" & TmpRs!ClosingWeight & ""
       End If
       
       If RsMst_StockStatus.RecordCount = 0 Then
          RsMst_StockStatus.AddNew
          RsMst_StockStatus!ID = GetMaxID
          RsMst_StockStatus!MonthYear = TmpRs!MonthYear
          RsMst_StockStatus!StateID = TmpRs!StateID
          RsMst_StockStatus!LocationID = TmpRs!LocationID
          RsMst_StockStatus!UnitTypeID = TmpRs!UnitTypeID
          RsMst_StockStatus!ClosingBags = TmpRs!ClosingBags
          RsMst_StockStatus!ClosingWeight = TmpRs!ClosingWeight
          RsMst_StockStatus!CMPID = TmpRs!CMPID
          
          RsMst_StockStatus.Update
       End If
       RsMst_StockStatus.Filter = ""
       TmpRs.MoveNext
   Next
End If

tmp = "Delete Mst_TmpStockStatus"
Call TmpTable

ProgressBar1.Value = ProgressBar1.Max
TxtErrorMsg = ErrMsg
MsgBox "Done!!!"
End Sub

Private Function GetMaxID() As Double
Dim Retval As Double
tmp = "Select max(ID) as ID from Mst_StockStatus"
Call Tmp0Table
Retval = IIf(IsNull(TmpRs0.Fields(0)), 0, TmpRs0.Fields(0))
Retval = Retval + 1

GetMaxID = Retval
End Function

Private Sub CmdTotalCapacityUpdate_Click()
tmp = "Select MC.CMPID,ML.LocationID,ML.LocationName,ut.UnitTypeID,ut.UnitType,Sum(MG.GodownCapacity) as TotalCapacity From Mst_Location ML "
tmp = tmp & " Left Outer Join Mst_CMP MC On ML.LocationID = MC.LocationID"
tmp = tmp & " Left Outer Join Mst_Godown MG On MC.CMPID = MG.CMPID"
tmp = tmp & " Inner join Mst_unittype ut On MG.UnitTypeID =ut.UnitTypeID"
tmp = tmp & " Where MG.CloseDate Is Null And ut.UnitTypeID <> 2"
tmp = tmp & " Group By MC.CMPID,ML.LocationID,ML.LocationName ,ut.UnitTypeID,ut.UnitType"

Call TmpTable

For I = 1 To TmpRs.RecordCount
    RsMst_StockStatus.AddNew
    RsMst_StockStatus!ID = GetMaxID
    RsMst_StockStatus!MonthYear = MSHFlexGrid1.TextMatrix(1, 0)
    RsMst_StockStatus!LocationID = TmpRs!LocationID
    RsMst_StockStatus!UnitTypeID = TmpRs!UnitTypeID
    RsMst_StockStatus!ClosingWeight = TmpRs!TotalCapacity
    RsMst_StockStatus!CMPID = TmpRs!CMPID
    
    TmpRs.MoveNext
Next
MsgBox "Total Capacity Updated Successfully !!! "
CmdTotalCapacityUpdate.Enabled = False
End Sub

Private Sub Form_Load()

lblFile.Caption = ""
Frame1.Enabled = False
Frame2.Enabled = False
Frame3.Enabled = False
Call Grid_Head
ProgressBar1.Value = 0

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame3.Enabled = True
CmdImport.Enabled = False
CmdSave.Enabled = False
CmdTotalCapacityUpdate.Enabled = True
CmdReset.Enabled = False
CmbMonthYear.Enabled = True
Dim mYear As Variant
mYear = Year(Now)
CmbMonthYear.AddItem "Jan" & " " & mYear
CmbMonthYear.AddItem "Feb" & " " & mYear
CmbMonthYear.AddItem "Mar" & " " & mYear
CmbMonthYear.AddItem "Apr" & " " & mYear
CmbMonthYear.AddItem "May" & " " & mYear
CmbMonthYear.AddItem "Jun" & " " & mYear
CmbMonthYear.AddItem "Jul" & " " & mYear
CmbMonthYear.AddItem "Aug" & " " & mYear
CmbMonthYear.AddItem "Sep" & " " & mYear
CmbMonthYear.AddItem "Oct" & " " & mYear
CmbMonthYear.AddItem "Nov" & " " & mYear
CmbMonthYear.AddItem "Dec" & " " & mYear
End Sub



Private Sub FillData()

Set oWS = oWB.Worksheets(CmbSheetName.Text)
oWS.ClearArrows
mRow = 1
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
'ProgressBar1.Value = 0
'ProgressBar1.Max = rowCount
Dim MonthYear As Variant
MonthYear = CmbMonthYear.Text
For l = 1 To RowCount
    mRow = mRow + 1
    If oWS.Cells(mRow, 1) <> "" Then
        For c = 0 To MSHFlexGrid1.Cols - 1
            If c = 0 Then
               MSHFlexGrid1.TextMatrix(l, c) = MonthYear
            ElseIf c = 6 Then
               MSHFlexGrid1.TextMatrix(l, c) = Trim(oWS.Cells(l + 1, 9).Value)
            ElseIf c = 7 Then
               MSHFlexGrid1.TextMatrix(l, c) = Trim(oWS.Cells(l + 1, 13).Value)
            Else
               MSHFlexGrid1.TextMatrix(l, c) = Trim(oWS.Cells(l + 1, c).Value)
            End If
         Next
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    End If
Next
oXL.Quit
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 8
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Month Year"
    .TextMatrix(0, 1) = "State"
    .TextMatrix(0, 2) = "Location"
    .TextMatrix(0, 3) = "CMP"
    .TextMatrix(0, 4) = "Unit Type"
    .TextMatrix(0, 5) = "Commodity"
    .TextMatrix(0, 6) = "Closing Bags"
    .TextMatrix(0, 7) = "Closing Weight"
    .ColWidth(0) = 1200
    .ColWidth(1) = 2000
    .ColWidth(2) = 2200
    .ColWidth(3) = 2200
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    
End With
End Sub


