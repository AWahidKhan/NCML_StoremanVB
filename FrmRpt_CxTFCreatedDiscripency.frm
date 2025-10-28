VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CxTFCreatedDiscripency 
   Caption         =   "CxTF Date and Created Date Discripency"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7935
      Left            =   120
      TabIndex        =   15
      Top             =   1365
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   7530
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12726
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
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
      Begin VB.CheckBox ChkAMOC 
         Appearance      =   0  'Flat
         Caption         =   "With AM and OC details"
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
         Left            =   1800
         TabIndex        =   19
         Top             =   960
         Width           =   2835
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Creation Date"
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
         Left            =   90
         TabIndex        =   17
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Transaction Date"
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
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
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
         Left            =   12615
         TabIndex        =   7
         Top             =   870
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetReport 
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
         Left            =   11520
         TabIndex        =   6
         Top             =   870
         Width           =   1095
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
         Left            =   13710
         TabIndex        =   8
         Top             =   870
         Width           =   1095
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
         Left            =   4635
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   2835
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
         Left            =   7515
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   405
         Width           =   4560
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   105447425
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   3225
         TabIndex        =   3
         Top             =   420
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   105447425
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label lblRec 
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
         Left            =   7680
         TabIndex        =   14
         Top             =   930
         Width           =   2325
      End
      Begin VB.Label lblRecordCount 
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
         Left            =   10080
         TabIndex        =   13
         Top             =   930
         Width           =   210
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
         Left            =   5520
         TabIndex        =   12
         Top             =   120
         Width           =   1065
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
         Left            =   9413
         TabIndex        =   11
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label3 
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
         Height          =   225
         Left            =   3465
         TabIndex        =   10
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   1965
         TabIndex        =   9
         Top             =   120
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmRpt_CxTFCreatedDiscripency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID As Variant
Private Sub CmdClear_Click()

Call Grid_Head
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbStateID.Text = ""
CmbLocationID.Text = ""
lblRecordCount.Caption = "-"
ChkAMOC.Value = 1

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CxTFCreatedDiscripency.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CxTFCreatedDiscripency.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
Set oWB = Nothing
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
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



Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
CmbLocationID.Clear

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.Cols = 11
'TCD.CreatedDate,TCD.CreatedBy,
'MU.Employee_Name
MSHFlexGrid1.TextMatrix(0, 0) = "State"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Txn Type"
MSHFlexGrid1.TextMatrix(0, 3) = "CxTF No."
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 5) = "CxTF Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Created Date"
MSHFlexGrid1.TextMatrix(0, 7) = "Day Difference"
MSHFlexGrid1.TextMatrix(0, 8) = "Created By"
MSHFlexGrid1.TextMatrix(0, 9) = "AM Name"
MSHFlexGrid1.TextMatrix(0, 10) = "OC Name"

MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(4) = 2200
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(8) = 2500
MSHFlexGrid1.ColWidth(9) = 2500
MSHFlexGrid1.ColWidth(10) = 2500

'MSHFlexGrid1.TextMatrix(0, 8) = "Employee Name"
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
wc = ""
Call Grid_Head
MsgBox "Wait till next message !!!"

If Option1.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where TCD.CxTFDate BETWEEN #" & TxtFromDate.Value & "# And  #" & TxtToDate.Value & "# And Left(TCD.CxTFDate,10) <> Left(TCD.CreatedDate,10) "
   Else
      wc = " Where TCD.CxTFDate BETWEEN '" & TxtFromDate.Value & " ' And  '" & TxtToDate.Value & "' And Left(TCD.CxTFDate,10) <> Left(TCD.CreatedDate,10) "
   End If
ElseIf Option2.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where TCD.CreatedDate BETWEEN #" & TxtFromDate.Value & "# And  #" & TxtToDate.Value + 1 & "# And Left(TCD.CxTFDate,10) <> Left(TCD.CreatedDate,10) "
   Else
      wc = " Where TCD.CreatedDate BETWEEN '" & TxtFromDate.Value & "' And  '" & TxtToDate.Value + 1 & "' And Left(TCD.CxTFDate,10) <> Left(TCD.CreatedDate,10) "
   End If
End If


If CmbLocationID.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbLocationID & "'"
End If

If CmbStateID.Text <> "" Then
   wc = wc & " And MS.StateName = '" & CmbStateID & "'"
End If


tmp = "Select MS.StateName,ML.LocationName,TCD.TxnType,TCD.CxTF_DetailsID,MCom.Commodity,TCD.CxTFDate,TCD.CreatedDate, DATEDIFF(day,TCD.CxTFDate,TCD.CreatedDate) as 'DayDiff', TCD.CreatedBy, MU.EmployeeName " & _
      " From Txn_CxTF_Details TCD " & _
      " Inner Join Mst_CMP MC On TCD.CMPID = MC.CMPID " & _
      "Inner Join Mst_Location ML On MC.LocationID = ML.LocationID " & _
      "Inner Join Mst_State MS On ML.StateID = MS.StateID " & _
      "Inner Join Mst_Commodity MCom on TCD.CommodityID = MCom.CommodityID " & _
      "Inner Join Mst_Users MU On TCD.CreatedBy = MU.EmployeeNo "

tmp1 = " Order By MS.StateName,ML.LocationName,TCD.CxTFDate,TCD.CreatedDate"
tmp = tmp & " " & wc & " " & tmp1

Call TmpTable

'Call dispdata
ProgressBar1.Value = 0
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 2
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!TxnType), "", TmpRs!TxnType)
       If MSHFlexGrid1.TextMatrix(I, 2) = "D" Then
           MSHFlexGrid1.TextMatrix(I, 2) = "Deposit"
       ElseIf MSHFlexGrid1.TextMatrix(I, 2) = "W" Then
           MSHFlexGrid1.TextMatrix(I, 2) = "Withdrawal"
       Else
           MSHFlexGrid1.TextMatrix(I, 2) = ""
       End If
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!CxTF_DetailsID), "", TmpRs!CxTF_DetailsID)
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!CxTFDate), "", Format(TmpRs!CxTFDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!DayDiff), "", TmpRs!DayDiff)
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!EmployeeName), "", TmpRs!EmployeeName) & "~" & IIf(IsNull(TmpRs!CreatedBy), "", TmpRs!CreatedBy)
       'MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!Employee_Name), "", TmpRs!Employee_Name)
       If ChkAMOC.Value = 1 Then
          GetAMOCName (I)
       End If
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

lblRecordCount.Caption = TmpRs.RecordCount
MsgBox "Done!!!"

End Sub
Private Sub GetAMOCName(x_ As Variant)

tmp = " Select Distinct  ME1.EmployeeName + '~'+ ME1.EmployeeNo 'AMName',ME2.EmployeeName + '~'+ ME2.EmployeeNo 'OCName'"
tmp = tmp & " From Mst_EmpCMPMapping MECM"
tmp = tmp & " Inner Join Mst_CMP MC On MECM.CMPID = MC.CMPID"
tmp = tmp & " inner join Mst_Location ML On MC.LocatioNID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Employee ME On MECM.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee ME1 On ME.DirectReportingTo = ME1.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee ME2 On ME1.EmployeeID = ME2.DirectReportingTo And ME2.DesignationID = 10"
tmp = tmp & " Where ME.Flag = 'Y' And ML.LocationName = '" & MSHFlexGrid1.TextMatrix(x_, 1) & "'"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x_, 9) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname) ' & "~" & IIf(IsNull(TmpRs!Employee_Name), "", TmpRs!Employee_Name)
   MSHFlexGrid1.TextMatrix(x_, 10) = IIf(IsNull(TmpRs1!OCName), "", TmpRs1!OCName) ' & "~" & IIf(IsNull(TmpRs!OCName), "", TmpRs!OCName)
End If

End Sub


Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Call CmdClear_Click

End Sub

Private Sub Option1_Click()

Call Grid_Head

End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub
