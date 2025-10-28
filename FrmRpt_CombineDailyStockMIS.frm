VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRpt_CombineDailyStockMIS 
   Caption         =   "Combine Daily StockMIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   9030
      Width           =   15015
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Stock Taken from Store-Man (Flag-0)"
         Height          =   195
         Left            =   11880
         TabIndex        =   21
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   11400
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Stock Taken From EHop (Flag-3)"
         Height          =   195
         Left            =   8880
         TabIndex        =   19
         Top             =   240
         Width           =   2340
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   8400
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Locations Not in Ehop but in Store-Man (Flag-1)"
         Height          =   195
         Left            =   4920
         TabIndex        =   17
         Top             =   240
         Width           =   3360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Locations not in Store-Man But in EHop  (Flag-2)"
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Height          =   7965
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3360
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5927
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4080
         Left            =   120
         TabIndex        =   22
         Top             =   3720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7197
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   75
         TabIndex        =   5
         Top             =   165
         Width           =   1455
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
         Height          =   350
         Left            =   4440
         TabIndex        =   4
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
         Height          =   350
         Left            =   5895
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
         Height          =   350
         Left            =   1530
         TabIndex        =   2
         Top             =   165
         Width           =   1455
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
         Left            =   2985
         TabIndex        =   1
         Top             =   165
         Width           =   1455
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   630
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   7920
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   7110
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
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   4620
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2025
      End
   End
End
Attribute VB_Name = "FrmRpt_CombineDailyStockMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim oWS1 As Excel.Worksheet
Dim oWS2 As Excel.Worksheet
Dim ErrMsg As Variant
Dim Status As Variant
Dim mSheetFlg As Variant
Dim LeaseStock, LeaseCapacity, ResCap, ResStock, LastRow, TotLeasecap, TotLeaseStock, TotUtiLease, TotFranchStock, TotResStock, TotResCap, TotStock As Long

' Variables for Total of RM
Dim rTotLeasecap, rTotLeaseStock, rTotUtiLease, rTotFranchStock, rTotResStock, rTotResCap, rTotStock As Long

' Variables for Total Stock
Dim tTotLeasecap, tTotLeaseStock, tTotUtiLease, tTotFranchStock, tTotResStock, tTotResCap, tTotStock As Long


Private Sub CmbSheetName_LostFocus()
If CmbSheetName.Text = "" Then
   mSheetFlg = Null
   Exit Sub
End If
If LCase(CmbSheetName.Text) = "store-man" Then
  mSheetFlg = "store-man"
Else
    MsgBox "Invalid Selection!!!"
    CmbSheetName.SetFocus
    Exit Sub
End If

Call Grid_Head

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CombineDailyStockMIS.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mCol As Integer
Dim mRow As Integer
Dim c As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CombineDailyStockMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Detailed Report"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = I + 1 ' + 6
    For c = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c) = Format(MSHFlexGrid1.TextMatrix(I, c), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, c) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           If c = MSHFlexGrid1.Cols - 1 And mRow > 2 Then
              Call PrintFlagValue(oWS, mRow, c, Val(MSHFlexGrid1.TextMatrix(I, c)))
           End If
           oWS.Cells(mRow, c) = MSHFlexGrid1.TextMatrix(I, c)
          
        End If
    Next
Next
oWS.Cells(2, 22) = "Flag Status"
' For Grid2( Sheet 2)
Set oWS1 = oWB.Worksheets("Sheet2")
oWS1.Name = "Concise Report"
oWS1.ClearArrows
mRow = 1
mCol = 1
For I = 0 To MSHFlexGrid2.Rows - 2
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        If c <> 0 And c <> 1 And c <> 3 Then
           If InStr(MSHFlexGrid2.TextMatrix(I, c), "/") > 0 Then
              If IsDate(MSHFlexGrid2.TextMatrix(I, c)) Then
                 oWS1.Cells(mRow, mCol) = Format(MSHFlexGrid2.TextMatrix(I, c), "dd/MMM/yyyy")
              Else
                 oWS1.Cells(mRow, mCol) = MSHFlexGrid2.TextMatrix(I, c)
              End If
           Else
              oWS1.Cells(mRow, mCol) = MSHFlexGrid2.TextMatrix(I, c)
           End If
           If I > 0 Then
              If MSHFlexGrid2.TextMatrix(I, 12) = "" And MSHFlexGrid2.TextMatrix(I - 1, 12) = "" Then
                 oWS1.Range(oWS1.Cells(mRow, 1), oWS1.Cells(mRow, mCol)).Font.Color = vbBlue
              ElseIf MSHFlexGrid2.TextMatrix(I, 12) = "" Then
                 oWS1.Range(oWS1.Cells(mRow, 1), oWS1.Cells(mRow, mCol)).Font.Color = vbRed
              End If
              
           End If
           mCol = mCol + 1
        End If
    Next
    mCol = 1
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Function PrintFlagValue(moWS As Excel.Worksheet, mRowNo As Integer, Col As Integer, mFlag As Integer)
If mFlag = 1 Then
   moWS.Cells(mRowNo, Col + 1) = "Location Not in Ehop but in Store-Man"
ElseIf mFlag = 2 Then
   moWS.Cells(mRowNo, Col + 1) = "Location not in Store-Man But in EHoP"
ElseIf mFlag = 3 Then
   moWS.Cells(mRowNo, Col + 1) = "Stock Taken From EHop"
ElseIf mFlag = 0 Then
   moWS.Cells(mRowNo, Col + 1) = "Actual Stock Taken from Store-Man"
End If
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

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
LblFile.Caption = Path

CmbSheetName.Clear
If LblFile.Caption <> "" Then
   CmdExit.Enabled = True
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
 ' Code For Creating List of Sheet in Excel Sheet
   Dim I As Integer

   For I = 1 To oWB.Worksheets.Count
       Set oWS = oWB.Worksheets(I)
       Status = 1
       Frame2.Enabled = False
       If LCase(oWS.Name) = "store-man" Or LCase(oWS.Name) = "ehop" Or LCase(oWS.Name) = "reserved stock" Then
          CmbSheetName.AddItem oWS.Name
       Else
          'Exit For
       End If
   Next
        
   If CmbSheetName.ListCount < 3 Then
      CmbSheetName.Clear
      Frame2.Enabled = False
      MsgBox "Invalid File!!!"
      Exit Sub
   End If
   CmbSheetName.Text = "store-man"
   MsgBox "Click Import To Show Records!!!"
   CmdImport.Enabled = True
   CmdReset.Enabled = True
   oXL.Quit
End If

Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CmdImport_Click()
If CmbSheetName.Text = "" Then
   MsgBox "Name Of Sheet Can't be blank!!!"
   CmbSheetName.SetFocus
   Exit Sub
End If
'-----------------------------
If CmbSheetName.Text <> "" Then
   MsgBox "Wait till next message!!!"
   CmdImport.Enabled = False
   CmdReset.Enabled = False
   CmdExit.Enabled = True
   CmdGetFile.Enabled = False
   'Call Grid_Head
   ErrMsg = ""
   Status = 0
   RowCount = 0
   I = 1
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   TxtErrorMsg = ""
   If Chk_FileHeadsStoreMan = False Then
      Exit Sub
   End If
   Call Grid_Head
   Call FillData
   CmdImport.Enabled = True
   CmdExcel.Enabled = True
   CmdExit.Enabled = True
   CmdReset.Enabled = True
   MsgBox "Import Succesful!!!"
End If
End Sub
Private Sub FillData()
Set oWS = oWB.Worksheets(CmbSheetName.Text)
Set oWS1 = oWB.Worksheets("ehop")
oWS.ClearArrows
oWS1.ClearArrows
Dim Flag As Boolean
Dim oArr
Flag = False
mRow = 2
mRow1 = 2
RowCountEhop = oWS1.Rows.Count
RowCountSM = oWS.Rows.Count
For l = 0 To RowCountSM + 2
If Trim(oWS.Cells(l + 1, 4)) = "" And l > 1 Then Exit For
  If oWS.Cells(l + 1, 1) <> "" Then
     For c = 0 To MSHFlexGrid1.Cols - 2
         If c = 1 Or c = 4 Then
            oArr = Split(Trim(oWS.Cells(l + 1, c + 1).Value), "~")
            If UBound(oArr) > 1 Then
               MSHFlexGrid1.TextMatrix(l, c) = oArr(0)
            Else
               MSHFlexGrid1.TextMatrix(l, c) = Trim(oWS.Cells(l + 1, c + 1).Value)
            End If
         Else
           If InStr(Trim(oWS.Cells(l + 1, c + 1).Value), "/") > 0 Then
              If IsDate(Trim(oWS.Cells(l + 1, c + 1).Value)) Then
                 MSHFlexGrid1.TextMatrix(l, c) = Format(Trim(oWS.Cells(l + 1, c + 1).Value), "dd/MMM/yyyy")
              Else
                 MSHFlexGrid1.TextMatrix(l, c) = Trim(oWS.Cells(l + 1, c + 1).Value)
              End If
           Else
'              oWS.Cells(mRow, c) = MSHFlexGrid1.TextMatrix(i, c)
              MSHFlexGrid1.TextMatrix(l, c) = Trim(oWS.Cells(l + 1, c + 1).Value)

           End If
         
         
         
                     End If
     Next
     MSHFlexGrid1.TextMatrix(l, 21) = 0
     MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
  End If
Next

' Add New locations from EHop   Flag = 2

For l = 2 To RowCountEhop + 2
   If Trim(oWS1.Cells(l + 1, 3)) = "" Then Exit For
   If oWS1.Cells(l + 1, 1) <> "" Then
      For c = 2 To MSHFlexGrid1.Rows - 1
         If Trim(LCase(oWS1.Cells(l + 1, 3))) = Trim(LCase(MSHFlexGrid1.TextMatrix(c, 3))) Then
            Flag = True
            Exit For
         Else
            Flag = False
            'Exit For
         End If
      Next

      If Flag = False And Trim(oWS.Cells(l + 1, 3)) <> "" Then
         LastRow = MSHFlexGrid1.Rows - 1
         MSHFlexGrid1.Row = LastRow
         For c = 0 To MSHFlexGrid1.Cols - 1
             MSHFlexGrid1.Col = c
             MSHFlexGrid1.CellBackColor = vbGreen
         Next
         MSHFlexGrid1.TextMatrix(LastRow, 1) = Trim(oWS1.Cells(l + 1, 1).Value)
         MSHFlexGrid1.TextMatrix(LastRow, 2) = Trim(oWS1.Cells(l + 1, 2).Value)
         MSHFlexGrid1.TextMatrix(LastRow, 4) = Trim(oWS1.Cells(l + 1, 4).Value)
         MSHFlexGrid1.TextMatrix(LastRow, 20) = ""

         MSHFlexGrid1.TextMatrix(LastRow, 3) = Trim(oWS1.Cells(l + 1, 3).Value)
         MSHFlexGrid1.TextMatrix(LastRow, 7) = Val(Trim(oWS1.Cells(l + 1, 6).Value)) + Val(Trim(oWS1.Cells(l + 1, 12).Value))
         MSHFlexGrid1.TextMatrix(LastRow, 10) = Val(Trim(oWS1.Cells(l + 1, 9).Value))
         MSHFlexGrid1.TextMatrix(LastRow, 11) = Val(Trim(oWS1.Cells(l + 1, 6).Value)) + Val(Trim(oWS1.Cells(l + 1, 12).Value)) + Val(Trim(oWS1.Cells(l + 1, 9).Value))


         MSHFlexGrid1.TextMatrix(LastRow, 21) = 2
         MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
      End If
   End If
Next

 ' Show Locations Which are not in Ehop   Flag = 1

For c = 2 To MSHFlexGrid1.Rows - 1
    For l = 2 To RowCountEhop + 2
        If Trim(oWS1.Cells(l + 1, 3)) = "" Then Exit For
        If Trim(LCase(MSHFlexGrid1.TextMatrix(c, 3))) = Trim(LCase(oWS1.Cells(l + 1, 3))) Then
           Flag = True
           Exit For
        Else
           Flag = False
           'Exit For
        End If
    Next

    If Flag = False And Trim(oWS1.Cells(l + 1, 3)) <> "" Then
       MSHFlexGrid1.Row = c
       For a = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.Col = a
           MSHFlexGrid1.CellBackColor = vbRed
       Next
       MSHFlexGrid1.TextMatrix(c, 21) = 1
    End If
Next

' Stock Of EHop is Greater than Store-Man   Flag = 3
For c = 2 To MSHFlexGrid1.Rows - 1
    For l = 2 To RowCountEhop + 2
        If Trim(oWS1.Cells(l + 1, 3)) = "" Then Exit For
        If oWS1.Cells(l + 1, 1) <> "" Then
           If Trim(LCase(MSHFlexGrid1.TextMatrix(c, 3))) = Trim(LCase(oWS1.Cells(l + 1, 3))) Then
              If Trim(LCase(oWS1.Cells(l + 1, 3))) = "abohar" Then
                 tmp = tmp
              End If
              If Val(MSHFlexGrid1.TextMatrix(c, 11)) = 0 Then ' < Val(oWS1.Cells(L + 1, 12)) Then
                 Flag = True
                 Exit For
              End If
           Else
              Flag = False
             'Exit For
           End If
        End If
    Next

    If Flag = True And Trim(oWS.Cells(l + 1, 3)) <> "" Then
       MSHFlexGrid1.Row = c
       For a = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.Col = a
           If MSHFlexGrid1.CellBackColor <> vbGreen Then
              MSHFlexGrid1.CellBackColor = vbYellow
              MSHFlexGrid1.TextMatrix(c, 21) = 3
           End If
       Next
       MSHFlexGrid1.TextMatrix(c, 1) = Trim(oWS1.Cells(l + 1, 1).Value)
       MSHFlexGrid1.TextMatrix(c, 2) = Trim(oWS1.Cells(l + 1, 2).Value)
       MSHFlexGrid1.TextMatrix(c, 4) = Trim(oWS1.Cells(l + 1, 4).Value)

       MSHFlexGrid1.TextMatrix(c, 20) = ""

       MSHFlexGrid1.TextMatrix(c, 7) = Val(Trim(oWS1.Cells(l + 1, 6).Value)) + Val(Trim(oWS1.Cells(l + 1, 12).Value))
       MSHFlexGrid1.TextMatrix(c, 10) = Val(Trim(oWS1.Cells(l + 1, 9).Value))
       MSHFlexGrid1.TextMatrix(c, 11) = Val(Trim(oWS1.Cells(l + 1, 6).Value)) + Val(Trim(oWS1.Cells(l + 1, 12).Value)) + Val(Trim(oWS1.Cells(l + 1, 9).Value))
    End If
Next

Call FillReservedStock
Call Grid1ColumnTotal
Call FillGridGroupStock

oXL.Quit
End Sub
Private Sub Grid1ColumnTotal()
For I = 2 To MSHFlexGrid1.Rows - 1
    For k = 5 To MSHFlexGrid1.Cols - 3
       MSHFlexGrid1.TextMatrix(0, k) = Format(Val(MSHFlexGrid1.TextMatrix(0, k)) + Val(MSHFlexGrid1.TextMatrix(I, k)), "0.000")
    Next
Next
End Sub
Private Sub FillReservedStock()
Set oWS2 = oWB.Worksheets("reserved stock")
oWS2.ClearArrows
Dim Flag As Boolean
Dim oArr
Flag = False
mRow = 2
mRow1 = 2
RowCount = oWS2.Rows.Count
For r = 1 To RowCount
    If Trim(oWS2.Cells(r, 1)) = "" Then Exit For
       For l = 2 To MSHFlexGrid1.Rows - 2
           If LCase(Trim(oWS2.Cells(r, 1))) = LCase(Trim(MSHFlexGrid1.TextMatrix(l, 3))) Then
              MSHFlexGrid1.TextMatrix(l, 19) = Val(Trim(oWS2.Cells(r, 2)))
              Exit For
           End If
       Next
Next
End Sub
Private Function Chk_FileHeadsStoreMan() As Boolean
Dim Retval As Boolean
Set oWS = oWB.Worksheets(CmbSheetName.Text)
Retval = True
If mSheetFlg = "store-man" Then
   If LCase(Trim(oWS.Cells(2, 2))) <> "region" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of B "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 3))) <> "state" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of c "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 4))) <> "location" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of D "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 5))) <> "area manager" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of E "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 6))) <> "lease capacity" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of F "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 7))) <> "lease accre qty" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of G "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 8))) <> "lease stock" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of H "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 9))) <> "franchise capacity" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of I "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 10))) <> "franchise accre qty" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of J "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 11))) <> "franchise stock" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of K "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 12))) <> "total stock" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of L "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 13))) <> "lease ncdex" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of M "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 14))) <> "lease nonncdex" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of N "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 15))) <> "lease spot" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of O "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 16))) <> "franch ncdex" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of P "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 17))) <> "franch nonncdex" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of Q "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 18))) <> "franch spot" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of R "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 19))) <> "reserved capacity" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of S "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 20))) <> "reserved stock" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of T "
      Chk_FileHeads = Retval
      Exit Function
   End If
   If LCase(Trim(oWS.Cells(2, 21))) <> "entered upto" Then
      Retval = False
      MsgBox "Please Check Sheet " & CmbSheetName.Text & " Column Name of U "
      Chk_FileHeads = Retval
      Exit Function
   End If
End If
Chk_FileHeadsStoreMan = Retval
End Function

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(1) = 800

MSHFlexGrid1.Cols = 22
MSHFlexGrid1.TextMatrix(1, 1) = " Region"
MSHFlexGrid1.TextMatrix(1, 2) = " State"
MSHFlexGrid1.TextMatrix(1, 3) = " Location"
MSHFlexGrid1.TextMatrix(1, 4) = " Area Manager"
MSHFlexGrid1.TextMatrix(1, 5) = " Lease Capacity (including Stock Management)"
MSHFlexGrid1.TextMatrix(1, 6) = " Lease Accredited Qty"
MSHFlexGrid1.TextMatrix(1, 7) = " Lease Stock (including stock under Stock Management)"
MSHFlexGrid1.TextMatrix(1, 8) = " Franchisee Capacity"
MSHFlexGrid1.TextMatrix(1, 9) = " Franchisee Accredited Qty"
MSHFlexGrid1.TextMatrix(1, 10) = " Franchisee Stock"
MSHFlexGrid1.TextMatrix(1, 11) = " Total Stock"

MSHFlexGrid1.TextMatrix(1, 12) = " Lease NCDEX"
MSHFlexGrid1.TextMatrix(1, 13) = " Lease Non-NCDEX"
MSHFlexGrid1.TextMatrix(1, 14) = " Lease Spot"
MSHFlexGrid1.TextMatrix(1, 15) = " Franchisee NCDEX"
MSHFlexGrid1.TextMatrix(1, 16) = " Franchisee Non-NCDEX"
MSHFlexGrid1.TextMatrix(1, 17) = " Franchisee Spot"
MSHFlexGrid1.TextMatrix(1, 18) = " Reserved Capacity"
MSHFlexGrid1.TextMatrix(1, 19) = " Reserved Stock"
MSHFlexGrid1.TextMatrix(1, 20) = " Entered UpTo  Date (i.e the oldest BOD date for a CMP in this location)"
MSHFlexGrid1.TextMatrix(1, 21) = " Flag"

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 2200
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 2200
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1700
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1600
MSHFlexGrid1.ColWidth(12) = 1400
MSHFlexGrid1.ColWidth(13) = 1400
MSHFlexGrid1.ColWidth(14) = 1600
MSHFlexGrid1.ColWidth(15) = 1400
MSHFlexGrid1.ColWidth(16) = 1600
MSHFlexGrid1.ColWidth(17) = 1400
MSHFlexGrid1.ColWidth(18) = 1400
MSHFlexGrid1.ColWidth(19) = 1200
MSHFlexGrid1.ColWidth(20) = 1900
MSHFlexGrid1.ColWidth(21) = 1000

Call Grid_Head1
End Sub
Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.Cols = 13
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.FixedCols = 4
MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.RowHeight(1) = 800

MSHFlexGrid2.TextMatrix(1, 0) = " Region"
MSHFlexGrid2.TextMatrix(1, 1) = " State"
MSHFlexGrid2.TextMatrix(1, 2) = " Location"
MSHFlexGrid2.TextMatrix(1, 3) = " Area Manager"
MSHFlexGrid2.TextMatrix(1, 4) = " Lease Capacity (including Stock Management)"
MSHFlexGrid2.TextMatrix(1, 5) = " Lease Stock (including stock under Stock Management)"
MSHFlexGrid2.TextMatrix(1, 6) = " Franchisee Stock"
MSHFlexGrid2.TextMatrix(1, 7) = " Total Stock"
MSHFlexGrid2.TextMatrix(1, 8) = " Reserved Capacity"
MSHFlexGrid2.TextMatrix(1, 9) = " Reserved Stock"
MSHFlexGrid2.TextMatrix(1, 10) = " Lease Capacity Utilisation"
MSHFlexGrid2.TextMatrix(1, 11) = " Entered UpTo  Date (i.e the oldest BOD date for a CMP in this location)"
MSHFlexGrid2.TextMatrix(1, 12) = " Flag Status"

MSHFlexGrid2.ColWidth(0) = 100
MSHFlexGrid2.ColWidth(1) = 0
MSHFlexGrid2.ColWidth(2) = 1800
MSHFlexGrid2.ColWidth(3) = 100
MSHFlexGrid2.ColWidth(4) = 1600
MSHFlexGrid2.ColWidth(5) = 1700
MSHFlexGrid2.ColWidth(6) = 1600
MSHFlexGrid2.ColWidth(7) = 1600
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 1600
MSHFlexGrid2.ColWidth(10) = 1600
MSHFlexGrid2.ColWidth(11) = 1900
MSHFlexGrid2.ColWidth(12) = 2200
End Sub

Private Sub CmdReset_Click()
Call Grid_Head
CmdGetFile.Enabled = True
CmdExcel.Enabled = False
CmdImport.Enabled = False
CmbSheetName.Clear
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
CmdExcel.Enabled = False
Call Grid_Head

End Sub
Private Sub FillGridGroupStock()
'Set oWS = oWB.Worksheets(CmbSheetName.Text)
'oWS.ClearArrows
Dim Flag As Boolean

Dim oArr
Flag = False
mRow = 2
mRow1 = 2

'RowCountSM = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
RowCountSM = MSHFlexGrid1.Rows  'oWS.Rows.Count
Call ClearStockRM
Call ClearStockAM
Call ClearStock
For r = 2 To RowCountSM - 2
    If CheckRM(Trim(MSHFlexGrid1.TextMatrix(r, 1))) = False Then
       For r1 = 2 To RowCountSM - 2
           'If Trim(oWS.Cells(r1, 4)) = "" Then Exit For
           If CheckAM(Trim(MSHFlexGrid1.TextMatrix(r1, 4))) = False Then
              If LCase(Trim(MSHFlexGrid1.TextMatrix(r, 1))) = LCase(Trim(MSHFlexGrid1.TextMatrix(r1, 1))) Then
                 For r2 = 1 To RowCountSM - 2
                     'If Trim(oWS.Cells(r2 + 1, 4)) = "" Then Exit For
                     If LCase(Trim(MSHFlexGrid1.TextMatrix(r1, 4))) = LCase(Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 4))) Then
                        LastRow = MSHFlexGrid2.Rows - 1
                        MSHFlexGrid2.TextMatrix(LastRow, 0) = Trim(MSHFlexGrid1.TextMatrix(r1, 1))
                        MSHFlexGrid2.TextMatrix(LastRow, 1) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 2))
                        MSHFlexGrid2.TextMatrix(LastRow, 2) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 3))
                        MSHFlexGrid2.TextMatrix(LastRow, 3) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 4))
                        MSHFlexGrid2.TextMatrix(LastRow, 4) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 5))
                        MSHFlexGrid2.TextMatrix(LastRow, 5) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 7))
                                             
                        MSHFlexGrid2.TextMatrix(LastRow, 6) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 10))
                        MSHFlexGrid2.TextMatrix(LastRow, 7) = Val(MSHFlexGrid2.TextMatrix(LastRow, 5)) + Val(MSHFlexGrid2.TextMatrix(LastRow, 6))
                        MSHFlexGrid2.TextMatrix(LastRow, 8) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 18))
                        MSHFlexGrid2.TextMatrix(LastRow, 9) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 19))
                        LeaseStock = Val(MSHFlexGrid2.TextMatrix(LastRow, 5))
                        LeaseCapacity = Val(MSHFlexGrid2.TextMatrix(LastRow, 4))
                        ResCap = Val(MSHFlexGrid2.TextMatrix(LastRow, 8))
                        ResStock = Val(MSHFlexGrid2.TextMatrix(LastRow, 9))
                        
                        
'                        If Val(MSHFlexGrid2.TextMatrix(LastRow, 4)) <> 0 Then
'                           MSHFlexGrid2.TextMatrix(LastRow, 10) = Format((LeaseStock + ResCap - ResStock) * 100 / LeaseCapacity, "0.000")
'                        Else
'                           MSHFlexGrid2.TextMatrix(LastRow, 10) = 0
'                        End If

' Change on 21-01-2010

                        If Val(MSHFlexGrid2.TextMatrix(LastRow, 4)) <> 0 Then
                           If ResCap > ResStock Then
                              MSHFlexGrid2.TextMatrix(LastRow, 10) = Format((LeaseStock + ResCap - ResStock) * 100 / LeaseCapacity, "0.000")
                           Else
                              MSHFlexGrid2.TextMatrix(LastRow, 10) = Format((LeaseStock) * 100 / LeaseCapacity, "0.000")
                           End If
                        Else
                           MSHFlexGrid2.TextMatrix(LastRow, 10) = 0
                        End If
                        
                        
                        MSHFlexGrid2.TextMatrix(LastRow, 11) = Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 20))
                        MSHFlexGrid2.TextMatrix(LastRow, 12) = GetFlagStatus(Trim(MSHFlexGrid1.TextMatrix(r2 + 1, 21)))
                        TotLeasecap = TotLeasecap + Val(MSHFlexGrid2.TextMatrix(LastRow, 4))
                        TotLeaseStock = TotLeaseStock + Val(MSHFlexGrid2.TextMatrix(LastRow, 5))
                        TotFranchStock = TotFranchStock + Val(MSHFlexGrid2.TextMatrix(LastRow, 6))
                        TotResStock = TotResStock + Val(MSHFlexGrid2.TextMatrix(LastRow, 9))
                        TotResCap = TotResCap + Val(MSHFlexGrid2.TextMatrix(LastRow, 8))
                        TotStock = TotStock + Val(MSHFlexGrid2.TextMatrix(LastRow, 7))
                        TotUtiLease = TotUtiLease + Val(MSHFlexGrid2.TextMatrix(LastRow, 10))
                        MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
                     End If
                         
                 Next
                 Call StockAM(Trim(MSHFlexGrid1.TextMatrix(r1, 4)))
              End If
           End If
       Next
       Call StockRM(Trim(MSHFlexGrid1.TextMatrix(r, 1)))
    End If
Next
Call CalculateTotal
oXL.Quit
End Sub
Private Function GetFlagStatus(mFlag As String) As String
Dim Retval As String
If mFlag = 1 Then
   Retval = "Location Not in Ehop but in Store-Man"
ElseIf mFlag = 2 Then
   Retval = "Location not in Store-Man But in EHoP"
ElseIf mFlag = 3 Then
   Retval = "Stock Taken From EHop"
ElseIf mFlag = 0 Then
   Retval = "Actual Stock Taken from Store-Man"
End If
GetFlagStatus = Retval
End Function
Private Sub CalculateTotal()
MSHFlexGrid2.TextMatrix(0, 4) = tTotLeasecap
MSHFlexGrid2.TextMatrix(0, 5) = tTotLeaseStock
MSHFlexGrid2.TextMatrix(0, 6) = tTotFranchStock
MSHFlexGrid2.TextMatrix(0, 7) = tTotStock
MSHFlexGrid2.TextMatrix(0, 8) = tTotResCap
MSHFlexGrid2.TextMatrix(0, 9) = tTotResStock
If tTotLeasecap <> 0 Then
   MSHFlexGrid2.TextMatrix(0, 10) = Format((tTotLeaseStock + tTotResCap - tTotResStock) * 100 / tTotLeasecap, "0.000")
Else
   MSHFlexGrid2.TextMatrix(0, 10) = 0
End If
End Sub
Private Sub StockAM(AMname As String)
LastRow = MSHFlexGrid2.Rows - 1
MSHFlexGrid2.TextMatrix(LastRow, 2) = AMname
MSHFlexGrid2.TextMatrix(LastRow, 4) = TotLeasecap
MSHFlexGrid2.TextMatrix(LastRow, 5) = TotLeaseStock
MSHFlexGrid2.TextMatrix(LastRow, 6) = TotFranchStock
MSHFlexGrid2.TextMatrix(LastRow, 7) = TotStock
MSHFlexGrid2.TextMatrix(LastRow, 8) = TotResCap
MSHFlexGrid2.TextMatrix(LastRow, 9) = TotResStock
If TotLeasecap <> 0 Then
   MSHFlexGrid2.TextMatrix(LastRow, 10) = Format((TotLeaseStock + TotResCap - TotResStock) * 100 / TotLeasecap, "0.000")
Else
   MSHFlexGrid2.TextMatrix(LastRow, 10) = 0
End If
rTotLeasecap = rTotLeasecap + TotLeasecap
rTotLeaseStock = rTotLeaseStock + TotLeaseStock
rTotFranchStock = rTotFranchStock + TotFranchStock
rTotStock = rTotStock + TotStock
rTotResCap = rTotResCap + TotResCap
rTotResStock = rTotResStock + TotResStock
rTotUtiLease = rTotUtiLease + TotUtiLease
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
RowColorAM (LastRow)
Call ClearStockAM
End Sub
Private Sub StockRM(RMName As String)
LastRow = MSHFlexGrid2.Rows - 1
MSHFlexGrid2.TextMatrix(LastRow, 2) = RMName
MSHFlexGrid2.TextMatrix(LastRow, 4) = rTotLeasecap
MSHFlexGrid2.TextMatrix(LastRow, 5) = rTotLeaseStock
MSHFlexGrid2.TextMatrix(LastRow, 6) = rTotFranchStock
MSHFlexGrid2.TextMatrix(LastRow, 7) = rTotStock
MSHFlexGrid2.TextMatrix(LastRow, 8) = rTotResCap
MSHFlexGrid2.TextMatrix(LastRow, 9) = rTotResStock
If rTotLeasecap <> 0 Then
   MSHFlexGrid2.TextMatrix(LastRow, 10) = Format((rTotLeaseStock + rTotResCap - rTotResStock) * 100 / rTotLeasecap, "0.000")
Else
   MSHFlexGrid2.TextMatrix(LastRow, 10) = 0
End If
tTotLeasecap = tTotLeasecap + rTotLeasecap
tTotLeaseStock = tTotLeaseStock + rTotLeaseStock
tTotFranchStock = tTotFranchStock + rTotFranchStock
tTotStock = tTotStock + rTotStock
tTotResCap = tTotResCap + rTotResCap
tTotResStock = tTotResStock + rTotResStock
tTotUtiLease = tTotUtiLease + rTotUtiLease
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
RowColorRM (LastRow)
Call ClearStockRM
End Sub
Private Function RowColorAM(RowNo As Integer)
MSHFlexGrid2.Row = RowNo
For m = 0 To MSHFlexGrid2.Cols - 1
    MSHFlexGrid2.Col = m
    MSHFlexGrid2.CellBackColor = vbYellow
    MSHFlexGrid2.CellFontBold = True
Next
End Function
Private Function RowColorRM(RowNo As Integer)
MSHFlexGrid2.Row = RowNo
For m = 0 To MSHFlexGrid2.Cols - 1
    MSHFlexGrid2.Col = m
    MSHFlexGrid2.CellBackColor = vbGreen
    MSHFlexGrid2.CellFontBold = True
Next
End Function
Private Sub ClearStockRM()
rTotLeasecap = 0
rTotLeaseStock = 0
rTotUtiLease = 0
rTotFranchStock = 0
rTotResStock = 0
rTotResCap = 0
rTotStock = 0
End Sub
Private Sub ClearStockAM()
TotLeasecap = 0
TotLeaseStock = 0
TotUtiLease = 0
TotFranchStock = 0
TotResStock = 0
TotResCap = 0
TotStock = 0
End Sub
Private Sub ClearStock()
tTotLeasecap = 0
tTotLeaseStock = 0
tTotUtiLease = 0
tTotFranchStock = 0
tTotResStock = 0
tTotResCap = 0
tTotStock = 0
End Sub
Private Function CheckRM(RM As String) As Boolean
Dim Retval As Boolean

For r = 2 To MSHFlexGrid2.Rows - 1
    If LCase(MSHFlexGrid2.TextMatrix(r, 0)) = LCase(RM) Then
       Retval = True
       Exit For
    Else
       Retval = False
    End If
Next
CheckRM = Retval
End Function
Private Function CheckAM(AM As String) As Boolean
Dim Retval As Boolean

For r = 2 To MSHFlexGrid2.Rows - 1
    If LCase(MSHFlexGrid2.TextMatrix(r, 3)) = LCase(AM) Then
       Retval = True
       Exit For
    Else
       Retval = False
    End If
Next
CheckAM = Retval
End Function

