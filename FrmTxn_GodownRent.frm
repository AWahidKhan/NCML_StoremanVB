VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_GodownRent 
   Caption         =   "Godown Rent"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   13425
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   1365
      Width           =   15015
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   435
         TabIndex        =   11
         Top             =   195
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12938
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
      Begin VB.Label LblTotalRecord 
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
         Left            =   11520
         TabIndex        =   23
         Top             =   120
         Width           =   90
      End
      Begin VB.Label Label7 
         Caption         =   "Total Record Found :"
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
         Left            =   9120
         TabIndex        =   22
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Show/Hide"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12240
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
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
         Height          =   450
         Left            =   7590
         TabIndex        =   21
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton SaveCmd 
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
         Height          =   450
         Left            =   6615
         TabIndex        =   20
         Top             =   660
         Width           =   975
      End
      Begin VB.CommandButton CmdShowAll 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7590
         TabIndex        =   19
         Top             =   195
         Width           =   975
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H0000C000&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8625
         TabIndex        =   17
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8625
         TabIndex        =   15
         Top             =   195
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8625
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton CmdShow 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6600
         TabIndex        =   8
         Top             =   195
         Width           =   975
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
         ItemData        =   "FrmTxn_GodownRent.frx":0000
         Left            =   5040
         List            =   "FrmTxn_GodownRent.frx":0002
         TabIndex        =   5
         Top             =   720
         Width           =   1455
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
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox CmbCMPMName 
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
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox TxtCMPMName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Godown is Close or Agreement  Started/Expiring in Current month "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9210
         TabIndex        =   18
         Top             =   772
         Width           =   5775
      End
      Begin VB.Label Label5 
         Caption         =   "Agreement is not approved by FAG"
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
         Left            =   9195
         TabIndex        =   16
         Top             =   210
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Godown Agreement is Expired"
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
         Left            =   9195
         TabIndex        =   14
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CMPM Name"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmTxn_GodownRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mEmployeeID, mEmployeeName, mDirectReportingTo As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mUserTypeFlag As Variant
Dim mSelRow As Variant
Dim mSelCol As Variant
Dim mFlagChk, mClaimID As Variant
Dim mMonth, mExpenseFrom, mExpenseTo, mTotalDays, mRow As Variant
Dim mMsg, mGridMsg As String
Dim oSMTP As SMTPSession
Dim mYear As Variant
Dim mColorRow As Boolean

Private Sub ChkCheckAll_Click()

If MSHFlexGrid1.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed And MSHFlexGrid1.CellBackColor <> vbYellow Then
          MSHFlexGrid1.row = I
          MSHFlexGrid1.TextMatrix(I, 0) = strChecked
       End If
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbCMPMName_GotFocus()
tmp = CmbCMPMName.Text
Call TableMst_Employee(" Where DesignationID=6 And Flag = 'Y' ")
CmbCMPMName.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbCMPMName.AddItem RsMst_Employee!EmployeeName
    RsMst_Employee.MoveNext
Next
CmbCMPMName.Text = tmp

End Sub
'----
Private Sub CmbCMPMName_LostFocus()
Call Grid_Head
If CmbCMPMName.Text = "" Then
   Exit Sub
End If
RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeName = '" & CmbCMPMName.Text & "'"
If RsMst_Employee.EOF Then
   MsgBox "Invalid selection "
   CmbCMPMName.SetFocus
   Exit Sub
End If
mEmployeeID = RsMst_Employee!EmployeeID
End Sub
Private Sub CmbMonth_LostFocus()
Call Grid_Head
If CmbMonth.Text = "" Then Exit Sub
mStatus = False
For I = 0 To 11
    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
       CmbMonth.Text = CmbMonth.List(I)
       mStatus = True
       Exit For
    End If
Next

If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If
mMonth = CmbMonth.ListIndex + 1
If CmbYear.Text <> "" Then
   mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))
End If
Call Grid_Head

End Sub

Private Sub CmbYear_GotFocus()
If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next


If Month(Now) = 12 Then
   CmbYear.AddItem Year(Now) + 1
End If

CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
Call Grid_Head

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
Dim myr As Variant

flg = "N"

If Month(Now) = 12 Then
   myr = Year(Now) + 1
Else
   myr = Year(Now)
End If


For I = 2010 To myr 'Year(Now)
    If CmbYear.Text = I Then
       flg = "Y"
       Exit For
    End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

mYear = CmbYear.Text
mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownRent.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownRent.xls")
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
                
        If MSHFlexGrid1.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbRed
        ElseIf MSHFlexGrid1.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid1.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbYellow
        End If
              
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdShow_Click()
Dim vMonth As Variant

Call Grid_Head
ChkCheckAll.Value = 0

If CmbCMPMName.Visible = True Then
   If CmbCMPMName.Text = "" Then
      MsgBox "Please Select CMPM Name From List !!!"
      CmbCMPMName.SetFocus
      Exit Sub
   End If
Else
   mEmployeeID = Gen_EmployeeID
End If

If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If
mColorRow = False


'Commented on 27 Dec 2010 'Parag
'If mMonth <= 3 Then
'   MsgBox " Month earlier to April 2010 is not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'Parag
'If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "head cmg" Then
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   If mMonth <= 11 And Val(CmbYear.Text) <= 2013 Then
      MsgBox " For processing rent earlier to December Month contact Mr. Manu Agrawal in CMG Department !!!"
      CmbMonth.SetFocus
      Exit Sub
   End If
End If


If mMonth <= 3 And Val(CmbYear.Text) <= 2010 Then
   MsgBox " Month earlier to April 2010 is not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

Dim mSelDate, mCurrDate, mFutDate As Variant

mSelDate = CmbYear.Text & "-" & mMonth & "-01"
mCurrDate = Year(Now) & "-" & Month(Now) & "-01"
mFutDate = DateAdd("m", 1, mCurrDate)
mFutDate = Format(mFutDate, "yyyy-MM-dd")
mSelDate = Format(mSelDate, "yyyy-MM-dd")


'Commented on 30th Sept 2013
'If mSelDate >= mFutDate Then
'   MsgBox "Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If


If Day(Date) >= Val(Gen_BlockGodRentProcessDay) Then
   If mSelDate > mFutDate Then
      MsgBox "Future Month not Allowed !!!"
      CmbMonth.SetFocus
      Exit Sub
   End If
Else
   If mSelDate >= mFutDate Then
      MsgBox "Future Month not Allowed !!!"
      CmbMonth.SetFocus
      Exit Sub
   End If
End If

'Commented on 27 Dec 2010 'Parag
'vMonth = Month(Now) + 1
'
'If mMonth > vMonth Then
'   MsgBox "Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbMonth.Text = "" Then
'   MsgBox "Please Select Month"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox "Please Select Year "
'   CmbYear.SetFocus
'   Exit Sub
'End If

''If mUserTypeFlag = "cmpm" Then
''   mEmployeeID = Gen_EmployeeID
''End If

mExpenseFrom = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mExpenseTo = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(mExpenseFrom), "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If

MsgBox "Please Wait till next message !!"

tmp = " Select GodownID,Sum(WHRWeight) as 'Stock' From"
'Parag
'tmp = tmp & " Mst_gsl  WHERE (Flag <> 'E') AND (Flag <> 'Z') group by GodownID"
tmp = tmp & " Mst_gsl  WHERE (Status <> 'E') AND (Status <> 'Z') group by GodownID"
Call Tmp1Table

tmp = " Select ClaimDate,ExpenseTo,LessorID,GodownID,Flag,LAID,ExpiryDate "
tmp = tmp & " From ( "
tmp = tmp & " Select TGR.ClaimDate,TGR.ExpenseTo,TGR.LessorID,TGRD.GodownID,TGR.Flag,TGRD.LAID,TLA.ExpiryDate from Txn_GodownRent TGR"
tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID=TGRD.LAID"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TGR.Flag <> 'S' And TGR.ExpenseTo = #" & Format(mExpenseTo, Gen_DatFormat) & "#"
Else
   tmp = tmp & " Where TGR.Flag <> 'S' And TGR.ExpenseTo = '" & Format(mExpenseTo, Gen_DatFormat) & "'"
End If
tmp = tmp & " Union All "
tmp = tmp & " Select TGR.ClaimDate,TGR.ExpenseTo,TGR.LessorID,TGRD.GodownID,'A' 'Flag',TGRD.LAID,TLA.ExpiryDate from Txn_GodownRent TGR"
tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID=TGRD.LAID"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TGR.Flag = 'S' And TGR.ExpenseTo = #" & Format(mExpenseTo, Gen_DatFormat) & "#"
Else
   tmp = tmp & " Where TGR.Flag = 'S' And TGR.ExpenseTo = '" & Format(mExpenseTo, Gen_DatFormat) & "'"
End If
tmp = tmp & ") A "
tmp = tmp & " order by Flag,LAID desc "

Call Tmp2Table


tmp = "Select TLA.LAID 'AgreementID',MLOC.LocationName,MC.CMPName,MG.GodownNo,MG.Address,ML.LessorName,TLA.StartDate,TLA.ExpiryDate,"
tmp = tmp & " MG.Area , TGLA.RentAmount, MG.GodownID, ML.LessorID, MG.CloseDate,TLA.Flag,MG.StartDate as 'GodownStartDate' ,TLA.AgreemenActualtStartDate"
tmp = tmp & " From Mst_EmpCMPMapping MEC"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MEC.CMPID"
tmp = tmp & " Inner Join Mst_Location MLOC on MLOC.LocationID=MC.LocationID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail TGLA On TGLA.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGLA.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Lessor ML on ML.LessorID=TGLA.LessorID"
tmp = tmp & " Where MEC.EmployeeID = " & mEmployeeID & ""

' Added by Parag on 1st Oct 2010
If Gen_DBType = "MDB" Then
   tmp = tmp & " And  MG.StartDate <= #" & Format(mExpenseTo, Gen_DatFormat) & "#"
Else
   tmp = tmp & " And MG.StartDate <= '" & Format(mExpenseTo, Gen_DatFormat) & "'"
End If
'''''Add this Line After Adding AgreementnotReceivedCheckbox in LessorAgreement on 10 july 2012 ''''
tmp = tmp & " And TLA.AgreenotReceieveFlag = '0'"
'''''Add this Line After Adding AgreementnotReceivedCheckbox in LessorAgreement on 10 july 2012 ''''

'''----Added by sagar on 30-04-14 for not considering Rejected Agreements
tmp = tmp & " And TLA.Flag <> 'R'"
'''-----

Call TmpTable


tmp = "Select TLA.LAID 'AgreementID',MLOC.LocationName,MC.CMPName,MG.GodownNo,MG.Address,ML.LessorName,TLA.StartDate,TLA.ExpiryDate,"
tmp = tmp & " MG.Area , TGLA.RentAmount, MG.GodownID, ML.LessorID, MG.CloseDate,TLA.Flag,MG.StartDate as 'GodownStartDate' ,TLA.AgreemenActualtStartDate"
tmp = tmp & " From Mst_EmpCMPMapping MEC"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MEC.CMPID"
tmp = tmp & " Inner Join Mst_Location MLOC on MLOC.LocationID=MC.LocationID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail TGLA On TGLA.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGLA.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Lessor ML on ML.LessorID=TGLA.LessorID"
tmp = tmp & " Where MEC.EmployeeID = " & mEmployeeID & ""

' Added by Parag on 1st Oct 2010
If Gen_DBType = "MDB" Then
   tmp = tmp & " And  TLA.StartDate <= #" & Format(mExpenseTo, Gen_DatFormat) & "# And TLA.ExpiryDate >= #" & Format(mExpenseFrom, Gen_DatFormat) & "#"
Else
   tmp = tmp & " And TLA.StartDate <= '" & Format(mExpenseTo, Gen_DatFormat) & "' And TLA.ExpiryDate >= '" & Format(mExpenseFrom, Gen_DatFormat) & "'"
End If
'''''Add this Line After Adding AgreementnotReceivedCheckbox in LessorAgreement on 10 july 2012 ''''
tmp = tmp & " And TLA.AgreenotReceieveFlag = '0'"

'''----Added by sagar on 30-04-14 for not considering Rejected Agreements
tmp = tmp & " And TLA.Flag <> 'R'"
'''-----


Call Tmp4Table


mRow = 1
If TmpRs.RecordCount > 0 Then
   For m = 1 To TmpRs.RecordCount
       ' incase Godown is not closed
'       If Format(mExpenseFrom, "yyyymm") = Format(TmpRs!StartDate, "yyyymm") Then
'          mExpenseFrom = TmpRs!StartDate
'       End If
       
       If IsNull(TmpRs!CloseDate) Then
          Call Check_Data
       Else ' if godown is closed
       ' 01/09/2010 <= 30/07/2010 sept July
          'If Format(TmpRs!CloseDate, "yyyymm") <= Format(mExpenseFrom, "yyyymm") Then
          
          'Commented Above condition on 15 th sept as instructed by Mayur Sir.
          If Format(TmpRs!CloseDate, "yyyymm") >= Format(mExpenseFrom, "yyyymm") Then
             Call Check_Data
          End If
       End If
       TmpRs.MoveNext
   Next
End If

'Dim mDeletedRow As Variant
'mDeletedRow = 0
'If Check1.Value = 0 Then
'   For I = 1 To MSHFlexGrid1.Rows - 1
'       MSHFlexGrid1.Row = I
'       If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit For
'
'       MSHFlexGrid1.Col = 0
'       If MSHFlexGrid1.CellBackColor = vbRed Then
'          MSHFlexGrid1.RemoveItem (I)
'          I = I - 1
'          mDeletedRow = mDeletedRow + 1
'       End If
'
'   Next
'   'MSHFlexGrid1.Row = I
'   For x = 1 To mDeletedRow
'       MSHFlexGrid1.AddItem (x)
'   Next
'End If
'MSHFlexGrid1.Refresh

LblTotalRecord.Caption = MSHFlexGrid1.Rows - 2
MsgBox "Done !!"
End Sub

Private Sub Check_Data()
Dim Mday, Nday As Variant
If IsNull(TmpRs!CloseDate) = False And Format(TmpRs!CloseDate, "yyyymm") < Format(mExpenseFrom, "yyyymm") Then
   Exit Sub
End If
TmpRs4.Filter = ""
With MSHFlexGrid1
If TmpRs!GodownID = 5295 Then
   a = a
End If
TmpRs2.Filter = "GodownID=" & TmpRs!GodownID & " And LessorID=" & TmpRs!LessorID & " And ExpenseTo=" & Format(mExpenseTo, "yyyy-mm-dd") & " And LAID=" & TmpRs!AgreementID
If TmpRs2.RecordCount > 0 Then
   For I = 1 To TmpRs2.RecordCount
       'TmpRs4.Filter = ""
       If LCase(TmpRs2!Flag) = "r" Then
          TmpRs4.Filter = "GodownID=" & TmpRs!GodownID & " And LessorID=" & TmpRs!LessorID & " And AgreementID=" & TmpRs!AgreementID & " "
          If TmpRs4.RecordCount = 0 Then
             Call dispdata(mRow)
          Else
             If TmpRs!AgreementID = TmpRs4!AgreementID Then
                Call dispdata(mRow)
             'Else
             '   mRow = mRow + 1
             End If
          End If
          TmpRs4.Filter = ""
          'mRow = mRow + 1
          Exit For
       Else
                    
''          If CDate(TmpRs2!ExpiryDate) < CDate(mExpenseTo) And CDate(TmpRs2!ExpiryDate) >= CDate(mExpenseFrom) And CDate(TmpRs!StartDate) > CDate(mExpenseFrom) And CDate(TmpRs!StartDate) < CDate(mExpenseTo) Then
''             If TmpRs!AgreementID <> TmpRs2!laid Then
''                Call dispdata(mRow)
''                Exit For
''             Else
''                TmpRs2.Filter = ""
''                Exit Sub
''             End If
''          Else
             TmpRs2.Filter = ""
             Exit Sub
''          End If
       End If
       TmpRs2.MoveNext
    Next
Else
'   Call dispdata(mRow)
   'mRow = mRow + 1

   TmpRs4.Filter = "GodownID=" & TmpRs!GodownID & " And LessorID=" & TmpRs!LessorID & " And AgreementID=" & TmpRs!AgreementID & " "
   If TmpRs4.RecordCount = 0 Then
      Call dispdata(mRow)
   Else
      If TmpRs!AgreementID = TmpRs4!AgreementID Then
         Call dispdata(mRow)
      'Else
      '   mRow = mRow + 1
      End If
   End If
   TmpRs4.Filter = ""
End If

If mColorRow = True Then
   If Format(TmpRs!StartDate, "yyyymm") > Format(mExpenseFrom, "yyyymm") Then 'And Format(TmpRs!ExpiryDate, "yyyymm") < Format(mExpenseFrom, "yyyymm") Then
       'Mark row as red
       .row = mRow - 1
       If .row <> 0 Then
          For o = 0 To .Cols - 1
              .Col = o
              .CellBackColor = vbRed
          Next
       End If
   End If

    
    
   If Format(TmpRs!CloseDate, "yyyymm") = Format(mExpenseFrom, "yyyymm") Then
       'Mark row as red
       .row = mRow - 1
       If .row <> 0 Then
          For o = 0 To .Cols - 1
              .Col = o
              .CellBackColor = vbGreen
          Next
       End If
       Mday = Val(Format(mExpenseTo, "dd"))
       Nday = Val(Format(TmpRs!CloseDate, "dd"))
       .TextMatrix(mRow - 1, 9) = (Val(.TextMatrix(mRow - 1, 15)) * (Nday)) / (Mday)
       .TextMatrix(mRow - 1, 9) = Format(.TextMatrix(mRow - 1, 9), "##########0")
   End If
    
   If Format(TmpRs!ExpiryDate, "yyyymm") = Format(mExpenseFrom, "yyyymm") Then
      'Mark row as red
      .row = mRow - 1
      If .row <> 0 Then
         For o = 0 To .Cols - 1
             .Col = o
             .CellBackColor = vbGreen
         Next
      End If
      Mday = Val(Format(mExpenseTo, "dd"))
      ' This code will take care of Godown Close and Expiry which ever is less.
      If TmpRs!ExpiryDate > TmpRs!CloseDate Then
         Nday = Val(Format(TmpRs!CloseDate, "dd"))
      Else
          Nday = Val(Format(TmpRs!ExpiryDate, "dd"))
      End If
      .TextMatrix(mRow - 1, 9) = (Val(.TextMatrix(mRow - 1, 15)) * (Nday)) / (Mday)
      .TextMatrix(mRow - 1, 9) = Format(.TextMatrix(mRow - 1, 9), "##########0")
   End If
   
      'Code to capture start date of the aggrement and Cal Rent
   
   If Format(TmpRs!StartDate, "yyyymm") = Format(mExpenseFrom, "yyyymm") Then
      .row = mRow - 1
      
      If Format(TmpRs!StartDate, "yyyymm") = Format(mExpenseFrom, "yyyymm") Then
         'Mday = Val(Format(TmpRs!ExpiryDate, "dd"))
         If .CellBackColor <> vbGreen Then
            .TextMatrix(mRow - 1, 9) = Val(.TextMatrix(mRow - 1, 15))
            Mday = Val(Format(mExpenseTo, "dd"))
            Nday = Mday - Val(Format(TmpRs!StartDate, "dd")) + 1
            .TextMatrix(mRow - 1, 9) = (Val(.TextMatrix(mRow - 1, 15)) * (Nday)) / (Mday)
         End If

'      Else

      End If
      
      
      
      .TextMatrix(mRow - 1, 9) = Format(.TextMatrix(mRow - 1, 9), "##########0")
      
      .row = mRow - 1
      If .row <> 0 Then
         For o = 0 To .Cols - 1
             .Col = o
             .CellBackColor = vbGreen
         Next
      End If
    
    End If
    
   If Format(TmpRs!StartDate, "yyyymm") = Format(TmpRs!ExpiryDate, "yyyymm") Then
      .row = mRow - 1
      .TextMatrix(mRow - 1, 9) = Val(.TextMatrix(mRow - 1, 15))
      Mday = Val(Format(mExpenseTo, "dd"))
      mSDay = Val(Format(TmpRs!StartDate, "dd"))
      Nday = Val(Format(TmpRs!ExpiryDate, "dd"))
      .TextMatrix(mRow - 1, 9) = (Val(.TextMatrix(mRow - 1, 15)) * ((Nday - mSDay) + 1)) / (Mday)
      
      .TextMatrix(mRow - 1, 9) = Format(.TextMatrix(mRow - 1, 9), "##########0")
      
      '.Row = mRow - 1
      If .row <> 0 Then
         For o = 0 To .Cols - 1
             .Col = o
             .CellBackColor = vbGreen
         Next
      End If
    
    End If
    
    
    If Format(TmpRs!ExpiryDate, "yyyymm") < Format(mExpenseFrom, "yyyymm") Then
       'Mark row as red
       .row = mRow - 1
       If .row <> 0 Then
          For o = 0 To .Cols - 1
              .Col = o
              .CellBackColor = vbRed
         Next
      End If
    End If

    If Format(TmpRs!StartDate, "yyyymm") > Format(mExpenseFrom, "yyyymm") Then
       'Mark row as red
       .row = mRow - 1
       If .row <> 0 Then
          For o = 0 To .Cols - 1
              .Col = o
              .CellBackColor = vbRed
         Next
      End If
    End If


    If Format(TmpRs!Flag) <> "A" And .CellBackColor <> vbRed Then
       'Mark row as Yellow for unapproved Agreement by FAG
       .row = mRow - 1
       If .row <> 0 Then
          For o = 0 To .Cols - 1
              .Col = o
              .CellBackColor = vbYellow
          Next
      End If
    End If
    mColorRow = False
End If

End With
TmpRs2.Filter = ""
                  
End Sub

Private Sub dispdata(mROw_ As Variant)

If TmpRs1.RecordCount > 0 Then
   TmpRs1.MoveFirst
End If
With MSHFlexGrid1
    TmpRs1.Find "GodownID=" & TmpRs!GodownID & ""
    If Not TmpRs1.EOF Then
       .TextMatrix(mROw_, 10) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    Else
       .TextMatrix(mROw_, 10) = 0
    End If
   .TextMatrix(mROw_, 1) = TmpRs!AgreementID
   .TextMatrix(mROw_, 2) = TmpRs!Locationname
   .TextMatrix(mROw_, 3) = TmpRs!CMPName
   .TextMatrix(mROw_, 4) = TmpRs!GodownNo
   .TextMatrix(mROw_, 5) = TmpRs!Address
   .TextMatrix(mROw_, 6) = TmpRs!LessorName
   .TextMatrix(mROw_, 7) = Format(TmpRs!ExpiryDate, "dd-MMM-yyyy")
   .TextMatrix(mROw_, 8) = TmpRs!Area
   .TextMatrix(mROw_, 9) = TmpRs!RentAmount
   .TextMatrix(mROw_, 11) = TmpRs!GodownID
   .TextMatrix(mROw_, 12) = TmpRs!LessorID
   .TextMatrix(mROw_, 14) = Format(IIf(IsNull(TmpRs!CloseDate), "", TmpRs!CloseDate), "dd-MMM-yyyy")
   .TextMatrix(mROw_, 15) = TmpRs!RentAmount
   
   If CmbCMPMName.Visible = False Then
      .TextMatrix(mROw_, 16) = TxtCMPMName.Text & "~" & Gen_UserLoginID
   Else
      .TextMatrix(mROw_, 16) = GetEmployeeNonName(mEmployeeID)
   End If
   .TextMatrix(mROw_, 17) = mEmployeeID
   
   .TextMatrix(mROw_, 18) = Format(TmpRs!GodownStartDate, "dd-MMM-yyyy")
   .TextMatrix(mROw_, 19) = Format(TmpRs!StartDate, "dd-MMM-yyyy")
   .TextMatrix(mROw_, 20) = IIf(IsNull(TmpRs!AgreemenActualtStartDate), "", Format(TmpRs!AgreemenActualtStartDate, "dd-MMM-yyyy"))
   
   .row = mROw_
   .Col = 0
   .CellFontName = "Wingdings"
   .CellFontSize = 12
   .CellAlignment = flexAlignCenterCenter
   .Text = strUnChecked
                  
   .Rows = .Rows + 1
    mRow = mRow + 1
    mColorRow = True
End With
End Sub

Private Sub CmdShowAll_Click()
SaveCmd.Enabled = False
Dim vMonth As Variant

Call Grid_Head
ChkCheckAll.Value = 0
CmbCMPMName.Text = ""
LblTotalRecord.Caption = "-"

If CmbMonth.Text = "" Then
   MsgBox " Blank Month not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Blank Year not allowed  !!!"
   CmbYear.SetFocus
   Exit Sub
End If

'Commented on 27 Dec 2010 'Parag
'If mMonth <= 3 Then
'   MsgBox " Month earlier to April 2010 is not allowed  !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If


If mMonth <= 3 And Val(CmbYear.Text) <= 2010 Then
   MsgBox " Month earlier to April 2010 is not allowed  !!!"
   CmbMonth.SetFocus
   Exit Sub
End If

mColorRow = False

Dim mSelDate, mCurrDate, mFutDate As Variant

mSelDate = CmbYear.Text & "-" & mMonth & "-01"
mCurrDate = Year(Now) & "-" & Month(Now) & "-01"
mFutDate = DateAdd("m", 1, mCurrDate)
mFutDate = Format(mFutDate, "yyyy-MM-dd")
mSelDate = Format(mSelDate, "yyyy-MM-dd")


'Commented on 30th Sept 2013
'If mSelDate >= mFutDate Then
'   MsgBox "Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If


If Day(Date) >= Val(Gen_BlockGodRentProcessDay) Then
   If mSelDate > mFutDate Then
      MsgBox "Future Month not Allowed !!!"
      CmbMonth.SetFocus
      Exit Sub
   End If
Else
   If mSelDate >= mFutDate Then
      MsgBox "Future Month not Allowed !!!"
      CmbMonth.SetFocus
      Exit Sub
   End If
End If


'Commented on 27 Dec 2010 'Parag
'vMonth = Month(Now) + 1
'
'If mMonth > vMonth Then
'   MsgBox "Future Month not Allowed !!!"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbMonth.Text = "" Then
'   MsgBox "Please Select Month"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'
'If CmbYear.Text = "" Then
'   MsgBox "Please Select Year "
'   CmbYear.SetFocus
'   Exit Sub
'End If

''If mUserTypeFlag = "cmpm" Then
''   mEmployeeID = Gen_EmployeeID
''End If


mExpenseFrom = Format(CmbYear.Text & "-" & mMonth & "-" & "01", "yyyy-mm-dd")
mExpenseTo = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

If Format(Gen_BlockingDate, "yyyy-mm") >= Format(CDate(mExpenseFrom), "yyyy-mm") Then
   MsgBox "Select next month from " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   CmbMonth.SetFocus
   Exit Sub
End If


MsgBox "Please Wait till next message !!"


tmp = " Select GodownID,Sum(WHRWeight) as 'Stock' From"
'Parag
'tmp = tmp & " Mst_gsl  WHERE (Flag <> 'E') AND (Flag <> 'Z') group by GodownID"
tmp = tmp & " Mst_gsl  WHERE (Status <> 'E') AND (Status <> 'Z') group by GodownID"
Call Tmp1Table

'Commented after Payment process enabled
'tmp = " Select TGR.ClaimDate,TGR.ExpenseTo,TGR.LessorID,TGRD.GodownID,TGR.Flag,TGRD.LAID,TLA.ExpiryDate from Txn_GodownRent TGR"
'tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
'tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID=TGRD.LAID"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Where TGR.ExpenseTo = #" & Format(mExpenseTo, Gen_DatFormat) & "#"
'Else
'   tmp = tmp & " Where TGR.ExpenseTo = '" & Format(mExpenseTo, Gen_DatFormat) & "'"
'End If
'tmp = tmp & " order by TGR.Flag "

'tmp = tmp & " Where  TGR.Flag <> 'R'"

tmp = " Select ClaimDate,ExpenseTo,LessorID,GodownID,Flag,LAID,ExpiryDate "
tmp = tmp & " From ( "
tmp = tmp & " Select TGR.ClaimDate,TGR.ExpenseTo,TGR.LessorID,TGRD.GodownID,TGR.Flag,TGRD.LAID,TLA.ExpiryDate from Txn_GodownRent TGR"
tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID=TGRD.LAID"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TGR.Flag <> 'S' And TGR.ExpenseTo = #" & Format(mExpenseTo, Gen_DatFormat) & "#"
Else
   tmp = tmp & " Where TGR.Flag <> 'S' And TGR.ExpenseTo = '" & Format(mExpenseTo, Gen_DatFormat) & "'"
End If
tmp = tmp & " Union All "
tmp = tmp & " Select TGR.ClaimDate,TGR.ExpenseTo,TGR.LessorID,TGRD.GodownID,'A' 'Flag',TGRD.LAID,TLA.ExpiryDate from Txn_GodownRent TGR"
tmp = tmp & " Inner Join Txn_GodownRentDetail TGRD On TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID=TGRD.LAID"
If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TGR.Flag = 'S' And TGR.ExpenseTo = #" & Format(mExpenseTo, Gen_DatFormat) & "#"
Else
   tmp = tmp & " Where TGR.Flag = 'S' And TGR.ExpenseTo = '" & Format(mExpenseTo, Gen_DatFormat) & "'"
End If
tmp = tmp & ") A "
tmp = tmp & " order by Flag,LAID desc "

Call Tmp2Table

Call TableMst_Employee
Call CmbCMPMName_GotFocus
mRow = 1



For k = 0 To CmbCMPMName.ListCount - 1
    If RsMst_Employee.RecordCount > 0 Then
       RsMst_Employee.MoveFirst
    End If
    RsMst_Employee.Find " EmployeeName='" & CmbCMPMName.List(k) & "'"
    If Not RsMst_Employee.EOF Then
       mEmployeeID = RsMst_Employee!EmployeeID
    End If
    
    tmp = "Select TLA.LAID 'AgreementID',MLOC.LocationName,MC.CMPName,MG.GodownNo,MG.Address,ML.LessorName,TLA.StartDate,TLA.ExpiryDate,"
    tmp = tmp & " MG.Area , TGLA.RentAmount, MG.GodownID, ML.LessorID, MG.CloseDate,TLA.Flag,MG.StartDate as 'GodownStartDate',TLA.AgreemenActualtStartDate"
    tmp = tmp & " From Mst_EmpCMPMapping MEC"
    tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MEC.CMPID"
    tmp = tmp & " Inner Join Mst_Location MLOC on MLOC.LocationID=MC.LocationID"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID=MC.CMPID"
    tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail TGLA On TGLA.GodownID=MG.GodownID"
    tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGLA.LAID=TLA.LAID"
    tmp = tmp & " Inner Join Mst_Lessor ML on ML.LessorID=TGLA.LessorID"
    tmp = tmp & " Where MEC.EmployeeID = " & mEmployeeID & ""

    ' Missed by parag added by Mayur Jan 2011 06 (Added by Parag on 1st Oct 2010)
    If Gen_DBType = "MDB" Then
       tmp = tmp & " And  MG.StartDate <= #" & Format(mExpenseTo, Gen_DatFormat) & "#"
    Else
       tmp = tmp & " And MG.StartDate <= '" & Format(mExpenseTo, Gen_DatFormat) & "'"
    End If

    '''''Add this Line After Adding AgreementnotReceivedCheckbox in LessorAgreement on 10 july 2012 ''''
    tmp = tmp & " And TLA.AgreenotReceieveFlag = '0'"
    '''''Add this Line After Adding AgreementnotReceivedCheckbox in LessorAgreement on 10 july 2012 ''''
    
    '''----Added by sagar on 30-04-14 for not considering Rejected Agreements
    tmp = tmp & " And TLA.Flag <> 'R'"
    '''-----
    
    Call TmpTable

    tmp = "Select TLA.LAID 'AgreementID',MLOC.LocationName,MC.CMPName,MG.GodownNo,MG.Address,ML.LessorName,TLA.StartDate,TLA.ExpiryDate,"
    tmp = tmp & " MG.Area , TGLA.RentAmount, MG.GodownID, ML.LessorID, MG.CloseDate,TLA.Flag,MG.StartDate as 'GodownStartDate' ,TLA.AgreemenActualtStartDate"
    tmp = tmp & " From Mst_EmpCMPMapping MEC"
    tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MEC.CMPID"
    tmp = tmp & " Inner Join Mst_Location MLOC on MLOC.LocationID=MC.LocationID"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID=MC.CMPID"
    tmp = tmp & " Inner Join Txn_GodownLessorAgreeDetail TGLA On TGLA.GodownID=MG.GodownID"
    tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TGLA.LAID=TLA.LAID"
    tmp = tmp & " Inner Join Mst_Lessor ML on ML.LessorID=TGLA.LessorID"
    tmp = tmp & " Where MEC.EmployeeID = " & mEmployeeID & ""
    
    ' Added by Parag on 1st Oct 2010
    If Gen_DBType = "MDB" Then
       tmp = tmp & " And  TLA.StartDate <= #" & Format(mExpenseTo, Gen_DatFormat) & "# And TLA.ExpiryDate >= #" & Format(mExpenseFrom, Gen_DatFormat) & "#"
    Else
       tmp = tmp & " And TLA.StartDate <= '" & Format(mExpenseTo, Gen_DatFormat) & "' And TLA.ExpiryDate >= '" & Format(mExpenseFrom, Gen_DatFormat) & "'"
    End If
    '''''Add this Line After Adding AgreementnotReceivedCheckbox in LessorAgreement on 10 july 2012 ''''
    tmp = tmp & " And TLA.AgreenotReceieveFlag = '0'"
    
    '''----Added by sagar on 30-04-14 for not considering Rejected Agreements
    tmp = tmp & " And TLA.Flag <> 'R'"
    '''-----
    
    Call Tmp4Table



    If TmpRs.RecordCount > 0 Then
       For m = 1 To TmpRs.RecordCount
           ' incase Godown is not closed
           If IsNull(TmpRs!CloseDate) Then
              Call Check_Data
           Else ' if godown is closed
              If Format(TmpRs!CloseDate, "yyyymm") >= Format(mExpenseFrom, "yyyymm") Then
                 Call Check_Data
              End If
           End If
           TmpRs.MoveNext
       Next
    End If
Next

LblTotalRecord.Caption = MSHFlexGrid1.Rows - 2
MsgBox "Done !!"

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
'SaveCmd.Enabled = False

Call CheckAccess
Call Fill_MonthYear
Call Grid_Head
Call Grid_Head1
'Call Fill_MstLessor
End Sub
Private Sub Fill_MonthYear()
For I = 1 To 12
   CmbMonth.AddItem MonthName(I), I - 1
Next

End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Refresh
    .Rows = 2
    .Cols = 21
    .WordWrap = True
    
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "Agreement ID"
    .TextMatrix(0, 2) = "Location"
    .TextMatrix(0, 3) = "CMP Name"
    .TextMatrix(0, 4) = "Godown No"
    .TextMatrix(0, 5) = "Address"
    .TextMatrix(0, 6) = "Lessor Name"
    .TextMatrix(0, 7) = "Agreement Expiry Date"
    .TextMatrix(0, 8) = "Area In Sq. ft."
    .TextMatrix(0, 9) = "Rent Payable "
    .TextMatrix(0, 10) = "Stock"
    .TextMatrix(0, 11) = "GodownID"
    .TextMatrix(0, 12) = "Lessor ID"
    .TextMatrix(0, 13) = "CheckFlag"
    .TextMatrix(0, 14) = "Godown Closed Date"
    .TextMatrix(0, 15) = "Monthly Rent Amount"
    .TextMatrix(0, 16) = "CMPM Name"
    .TextMatrix(0, 17) = "CMPM ID"
    .TextMatrix(0, 18) = "Godown Start Date"
    .TextMatrix(0, 19) = "Agreement Start Date"
    .TextMatrix(0, 20) = "Agreement Actual Start Date"
    
    .ColWidth(0) = 700
    .ColWidth(1) = 1500
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 1200
    .ColWidth(5) = 3500
    .ColWidth(6) = 1800
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    .ColWidth(10) = 1500
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 1500
    .ColWidth(15) = 1500
    .ColWidth(16) = 2200
    .ColWidth(17) = 1200
    
    .RowHeight(0) = 500
    
End With
End Sub
Private Sub Grid_Head1()
With MSHFlexGrid2
    .Rows = 2
    .Cols = 3
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Lessor ID"
    .TextMatrix(0, 1) = "Lessor Name"
    .ColWidth(0) = 1200
    .ColWidth(1) = 2500
    
End With
End Sub
Private Sub Fill_MstLessor()
'Call TableMst_Lessor
'If RsMst_Lessor.RecordCount > 0 Then
'   For i = 1 To RsMst_Lessor.RecordCount
'       MSHFlexGrid2.TextMatrix(i, 0) = RsMst_Lessor!LessorID
'       MSHFlexGrid2.TextMatrix(i, 1) = RsMst_Lessor!LessorName
'       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
'       RsMst_Lessor.MoveNext
'   Next
'End If
MSHFlexGrid2.Rows = 1
MSHFlexGrid2.Clear
Dim flg As Variant
For y = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(y, 0) = strChecked Then
       flg = "Y"
       For o = 1 To MSHFlexGrid2.Rows - 1
          If MSHFlexGrid2.TextMatrix(o, 0) = MSHFlexGrid1.TextMatrix(y, 12) Then
             flg = "N"
             Exit For
          End If
       Next
       If flg = "Y" Then
          I = MSHFlexGrid2.Rows
          MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
          MSHFlexGrid2.TextMatrix(I, 0) = MSHFlexGrid1.TextMatrix(y, 12)
          MSHFlexGrid2.TextMatrix(I, 1) = MSHFlexGrid1.TextMatrix(y, 6)
       End If
  End If
Next
End Sub
Private Sub CheckAccess()
CmdShowAll.Visible = False


'If LCase(Gen_UserRole) = "superadmin" Then
'   CmdShowAll.Visible = True
'End If
'If LCase(Gen_UserRole) = "admin" Then
'   CmdShowAll.Visible = True
'End If

If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "headcmg" Then
   TxtCMPMName.Visible = False
   CmbCMPMName.Visible = True
   mUserTypeFlag = "power"
   CmdShowAll.Visible = True

ElseIf Gen_DesignationID = 6 Then
   TxtCMPMName.Visible = True
   TxtCMPMName.Text = Gen_UserName
   CmbCMPMName.Visible = False
   mUserTypeFlag = "cmpm"
Else
   Frame1.Enabled = False
   Frame2.Enabled = False
End If
End Sub

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
    If .Col = 0 Then  '.Row = 1 And
       If .TextMatrix(.row, .Col) = "" Then Exit Sub
       'If (.TextMatrix(.Row, 12)) = "" Then
       .Col = 0
       If (.CellBackColor) <> vbRed And (.CellBackColor) <> vbYellow Then
          If .TextMatrix(.row, .Col) = strUnChecked Then
             .TextMatrix(.row, .Col) = strChecked
          Else
             .TextMatrix(.row, .Col) = strUnChecked
          End If
       Else
          If .CellBackColor = vbRed Then
             MsgBox "Agreement is expired please request CMG department to change expiry date!!!"
          ElseIf .CellBackColor = vbYellow Then
             MsgBox "Agreement is not approved by FAG !!!"
          End If
       End If
    End If
End With
End Sub
Private Sub SaveCmd_Click()

If MSHFlexGrid1.Rows = 2 Then
   MsgBox "No Record Present To Save !!!"
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "admin" And LCase(Gen_UserRole) <> "superadmin" And LCase(Gen_UserRole) <> "headcmg" Then
   If mMonth <= 11 And Val(CmbYear.Text) <= 2013 Then
      MsgBox " For processing rent earlier to December Month contact Mr. Bhagwat in CMG Department !!!"
      CmbMonth.SetFocus
      Exit Sub
   End If
End If


SaveCmd.Enabled = False

tmp = "Select ME.EmployeeID,ME.EmployeeName from Mst_Employee ME"
tmp = tmp & " Inner Join Mst_Employee EMP on EMP.DirectReportingTo=ME.EmployeeID"
tmp = tmp & " Where EMP.EmployeeID = " & mEmployeeID & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   mDirectReportingTo = TmpRs!EmployeeID
Else
   MsgBox "Direct Reporting Not Available, You Can't Create Claim !!!"
   Exit Sub
End If

Call Fill_MstLessor 'you will get Distinct LessorID from selected List

mGridMsg = "<tr  bgcolor=" & "#E9CFEC" & "  style=" & "font:bold" & " >"  '
For mMail = 0 To MSHFlexGrid1.Cols - 4
    If mMail = 0 Then
       mGridMsg = mGridMsg & " <td> Claim ID </td>"
    ElseIf mMail = 11 Then
       mGridMsg = mGridMsg & " <td> " & MSHFlexGrid1.TextMatrix(0, 14) & "</td>"
    ElseIf mMail = 12 Then
       mGridMsg = mGridMsg & " <td> " & MSHFlexGrid1.TextMatrix(0, 15) & "</td>"
    Else
       mGridMsg = mGridMsg & " <td> " & MSHFlexGrid1.TextMatrix(0, mMail) & "</td>"
    End If
Next
mGridMsg = mGridMsg & "</tr>"


Call TableTxn_GodownRent
Call TableTxn_GodownRentDetail
Call TableTxn_GodownLessorAgreeDetail

Dim mLessorID As Double
For I = 1 To MSHFlexGrid2.Rows - 1
    mFlagChk = 0
    If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit For
    mLessorID = MSHFlexGrid2.TextMatrix(I, 0)
    Call GetClaimDetails(mLessorID)
Next

If MSHFlexGrid2.Rows > 1 Then
   Call Gen_Email(mGridMsg)
   MsgBox "Done !"
   Call Grid_Head
Else
   MsgBox "No record to save !!!"
End If

ChkCheckAll.Value = 0

End Sub

Private Sub GetClaimDetails(mLessorID_ As Double)

mExpenseTo = Format(CmbYear.Text & "-" & mMonth & "-" & mTotalDays, "yyyy-mm-dd")

RsTxn_GodownRent.AddNew
'mClaimID = GetMaxClaimID
'RsTxn_GodownRent!ClaimID = mClaimID
RsTxn_GodownRent!ClaimDate = Date
RsTxn_GodownRent!LessorID = mLessorID_ ' Val(MSHFlexGrid1.TextMatrix(j, 10))
RsTxn_GodownRent!Amount = 0 'Val(MSHFlexGrid1.TextMatrix(j, 7))
RsTxn_GodownRent!ExpenseFrom = mExpenseFrom
RsTxn_GodownRent!ExpenseTo = mExpenseTo
RsTxn_GodownRent!ForwardToID = mDirectReportingTo ' mEmployeeID
RsTxn_GodownRent!Flag = "P"
RsTxn_GodownRent!BatchID = 0
RsTxn_GodownRent!CMPMID = Val(MSHFlexGrid1.TextMatrix(1, 17))
RsTxn_GodownRent!G_UID = GetGUID
mG_UIDRent = RsTxn_GodownRent!G_UID
RsTxn_GodownRent!CreatedBy = Gen_UserLoginID
RsTxn_GodownRent!CreatedDate = Now
RsTxn_GodownRent!Remarks = "Being Rent for the month of " & CmbMonth.Text & " - " & CmbYear.Text
RsTxn_GodownRent.Update

tmp = "Select ClaimID From Txn_GodownRent Where G_UID='" & mG_UIDRent & "'"
Call Tmp7Table

If TmpRs7.RecordCount > 0 Then
'   mCxTF_DetailsID = TmpRs!CxTF_DetailsID
   mClaimID = TmpRs7!ClaimID
End If

For j = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(j, 0) = strChecked And Val(MSHFlexGrid1.TextMatrix(j, 12)) = mLessorID_ Then
       mFlagChk = 1
    Else
       mFlagChk = 0
    End If
    If mFlagChk = 1 Then
       RsTxn_GodownRentDetail.AddNew
'       RsTxn_GodownRentDetail!ClaimDetailID = GetMaxClaimDetailID
       RsTxn_GodownRentDetail!ClaimID = mClaimID
       RsTxn_GodownRentDetail!GodownID = Val(MSHFlexGrid1.TextMatrix(j, 11))
       RsTxn_GodownRentDetail!ExpenseTypeID = GetExpenseTypeID("Warehouse Rent")
       RsTxn_GodownRentDetail!Amount = Val(MSHFlexGrid1.TextMatrix(j, 9))
       RsTxn_GodownRentDetail!CreatedBy = Gen_UserLoginID
       RsTxn_GodownRentDetail!CreatedDate = Now
       RsTxn_GodownRentDetail!LAID = Val(MSHFlexGrid1.TextMatrix(j, 1))
       RsTxn_GodownRentDetail.Update
     
     
       RsTxn_GodownRent!Amount = RsTxn_GodownRent!Amount + Val(MSHFlexGrid1.TextMatrix(j, 9))
       RsTxn_GodownRent.Update
       
       'Code To Update RentpaidUpTo Date of Godown Lessor Agreement Detail table

       tmp = "Select Max(RentPaidUpto) as 'MaxRentPaidDate' From Txn_GodownLessorAgreeDetail "
       tmp = tmp & " Where LAID=" & Val(MSHFlexGrid1.TextMatrix(j, 1)) & " And GodownID=" & Val(MSHFlexGrid1.TextMatrix(j, 11)) & " And LessorID=" & Val(MSHFlexGrid1.TextMatrix(j, 12)) & ""
       Call TmpTable
       
       RsTxn_GodownLessorAgreeDetail.Filter = "LAID=" & Val(MSHFlexGrid1.TextMatrix(j, 1)) & " And GodownID=" & Val(MSHFlexGrid1.TextMatrix(j, 11)) & " And LessorID=" & Val(MSHFlexGrid1.TextMatrix(j, 12)) & ""
       
       If TmpRs.RecordCount > 0 Then
          If CDate(MSHFlexGrid1.TextMatrix(j, 7)) <= mExpenseTo Then
             mExpenseTo = CDate(MSHFlexGrid1.TextMatrix(j, 7))
          End If
          
'          If IsNull(TmpRs!MaxRentPaidDate) = True Then
             RsTxn_GodownLessorAgreeDetail!RentPaidUpTo = CDate(mExpenseTo)
             RsTxn_GodownLessorAgreeDetail.Update
'          ElseIf TmpRs!MaxRentPaidDate < CDate(mExpenseTo) Then
'             RsTxn_GodownLessorAgreeDetail!RentPaidUpTo = CDate(mExpenseTo)
'             RsTxn_GodownLessorAgreeDetail.Update
'          End If
'
          
          
          
       End If
       
       RsTxn_GodownLessorAgreeDetail.Filter = ""
      
       mGridMsg = mGridMsg & "<tr>"
       For mMail = 0 To MSHFlexGrid1.Cols - 4
           If mMail = 0 Then
              mGridMsg = mGridMsg & " <td> " & mClaimID & "</td>"
           ElseIf mMail = 11 Then
              mGridMsg = mGridMsg & " <td> " & IIf(MSHFlexGrid1.TextMatrix(j, 14) = "", "&nbsp", MSHFlexGrid1.TextMatrix(j, 14)) & "</td>"
           ElseIf mMail = 12 Then
              mGridMsg = mGridMsg & " <td> " & IIf(MSHFlexGrid1.TextMatrix(j, 15) = "", "&nbsp", MSHFlexGrid1.TextMatrix(j, 15)) & "</td>"
           Else
              mGridMsg = mGridMsg & " <td> " & IIf(MSHFlexGrid1.TextMatrix(j, mMail) = "", "&nbsp", MSHFlexGrid1.TextMatrix(j, mMail)) & "</td>"
           End If
       Next
       mGridMsg = mGridMsg & "</tr>"

    End If
Next
End Sub


Private Function GetMaxClaimID()
Dim Retval As Double
tmp = "Select max(ClaimID) from Txn_GodownRent"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClaimID = Retval
End Function

Private Function GetMaxClaimDetailID()
Dim Retval As Double
tmp = "Select max(ClaimDetailID) from Txn_GodownRentDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxClaimDetailID = Retval
End Function

Private Function Gen_Email(mMailData_ As String)
Dim TOADD, TOCC As Variant
Dim mCMPMName, mEmpName_ As Variant

If CmbCMPMName.Visible = True Then
   mCMPMName = GetEmployeeNonName(mEmployeeID)
Else
   mCMPMName = GetEmployeeNonName(Gen_EmployeeID)
End If


mMsg = ""
mMsg = mMsg & "<tr style=" & "font:bold" & " > <td Colspan=2><C><b>Godown Rent Claim (From Store-Man)</b></C></td><tr>"
mMsg = mMsg & " <tr><td><b>CMPM Name</b></td><td>" & mCMPMName & "</td></tr>"
mMsg = mMsg & " <tr><td><b>Rent From</b> </td><td>" & Format(mExpenseFrom, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & " <tr><td><b>Rent To</b></td><td>" & Format(mExpenseTo, "dd-MMM-yyyy") & "</td></tr>"
mEmpName_ = GetEmployeeNonName(mDirectReportingTo)
mMsg = mMsg & " <tr><td><b>Forwarded To</b></td><td>" & mEmpName_ & "</td></tr>"
mMsg = mMsg & "<tr> <td Colspan=2><b>Details</b></td><tr>"

mMsg = mMsg & "<tr> <td Colspan=2><table border=1>" & mGridMsg & "</table></td> </tr> "
mMsgSubject = "Godown Rent Claim"
mForwardToEmailID = GetEmployeeEmailID(mDirectReportingTo)


tmp = "Select ME.EmployeeID,ME.EmployeeName,ME.EmployeeEmailID from Mst_Employee ME"
tmp = tmp & " Inner Join Mst_Employee EMP on EMP.DirectReportingTo=ME.EmployeeID"
tmp = tmp & " Where EMP.EmployeeID = " & mDirectReportingTo & ""
Call TmpTable
If TmpRs.RecordCount > 0 Then
   mForwardToEmailID = mForwardToEmailID & "," & TmpRs!EmployeeEmailID
End If
TOADD = mForwardToEmailID
TOCC = Gen_UserEmailID
Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
End Function
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face=Zurich BT color = Blue><table border = 1 cellpadding=1 cellspacing =1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = TOCC_ '"mayur.d@ncml.com"
'  If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
'     ToADD_ = "bhagwat.k@ncml.com"
'  End If
'  .SendTo = ToADD_
'  'TOCC_ = TOCC_ & ",cmg@ncml.com"
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  '.BCC = "mayur.d@ncml.com"
'  .BCC = "parag.r@ncml.com"
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = TOCC_
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "vaishali.k@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If


End Function

Private Function GetAgreeEndDate(mLAID_ As Variant) As Variant

End Function
