VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_LoginDetails 
   Caption         =   "Login Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   14775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6975
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   12303
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.CheckBox ChkSummary 
         Caption         =   "Summary"
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
         Left            =   4500
         TabIndex        =   20
         Top             =   180
         Width           =   1245
      End
      Begin VB.TextBox TxtSMenuName 
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
         Left            =   8925
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   5520
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Form Access Detail"
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
         Left            =   2040
         TabIndex        =   17
         Top             =   120
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Login Detail"
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
         Left            =   150
         TabIndex        =   16
         Top             =   120
         Width           =   1485
      End
      Begin VB.TextBox TxtEmployeeID 
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
         Left            =   2910
         TabIndex        =   3
         Top             =   840
         Width           =   1485
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
         Height          =   420
         Left            =   13110
         TabIndex        =   7
         Top             =   1500
         Width           =   1335
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
         Height          =   420
         Left            =   11775
         TabIndex        =   6
         Top             =   1500
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
         Height          =   420
         Left            =   10440
         TabIndex        =   5
         Top             =   1500
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1365
         _ExtentX        =   2408
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
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
         Format          =   111542275
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   1500
         TabIndex        =   2
         Top             =   840
         Width           =   1380
         _ExtentX        =   2434
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
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
         Format          =   111542275
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.TextBox TxtEmployeeName 
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
         Left            =   4425
         TabIndex        =   4
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label6 
         Caption         =   "Menu Name"
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
         Left            =   11055
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label5 
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
         Left            =   5520
         TabIndex        =   14
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Employee Name"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   600
         Width           =   1575
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
         Left            =   7920
         TabIndex        =   13
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label Label3 
         Caption         =   "Employee ID"
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
         Left            =   3045
         TabIndex        =   12
         Top             =   600
         Width           =   1215
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
         Left            =   255
         TabIndex        =   11
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   1785
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmRpt_LoginDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkSummary_Click()

If ChkSummary.Value = 0 Then
   If Option1.Value = True Then
      Call Grid_Head
   Else
      Call Grid_HeadFormDetail
   End If
Else
   If Option1.Value = True Then
      Call Grid_HeadSummary
   Else
      Call Grid_HeadFormSummary
   End If
End If

End Sub

Private Sub CmdClear_Click()
Option1.Value = True
Call Grid_Head
TxtEmployeeID.Text = ""
TxtEmployeeName.Text = ""
TxtToDate.Value = Date
TxtFromDate.Value = Date
LblTotalRecord.Caption = "-"
TxtSMenuName = ""
End Sub

Private Sub CmdExcel_Click()
If MSHFlexGrid1.Rows < 3 And MSHFlexGrid1.TextMatrix(1, 0) = "" Then MsgBox "No Record Found": Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_LoginDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_LoginDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
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
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
LblTotalRecord.Caption = "-"
If TxtToDate.Value < TxtFromDate.Value Then
   MsgBox "To Date must be more than From Date", vbInformation
   TxtFromDate.SetFocus
   Exit Sub
End If

If Option1.Value = True Then
   If ChkSummary.Value = 0 Then
      Call Grid_Head
   
      tmp = "Select LD.Login_Date ,LD.EmployeeID,MU.EmployeeName,MD.Designation,ME.EmployeeEmailID,ME.MobileNo"
      tmp = tmp & " From Mst_Employee ME"
      tmp = tmp & " Right Join  Log_LoginDetail LD on ME.EmployeeNo=LD.EmployeeID"
      tmp = tmp & " LEFT Join Mst_Designation MD on MD.DesignationID=ME.DesignationID"
      tmp = tmp & " Inner Join Mst_Users MU on MU.EmployeeNo=LD.EmployeeID"
      
      If Gen_DBType = "MDB" Then
         tmp = tmp & " Where   (LD.Login_Date >= #" & TxtFromDate.Value & "# and LD.Login_Date <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM #) "
      Else
         tmp = tmp & " Where   (LD.Login_Date >= '" & TxtFromDate.Value & "' and LD.Login_Date <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM' ) "
      End If
      If TxtEmployeeID.Text <> "" Then
         tmp = tmp & " and LD.EmployeeID='" & TxtEmployeeID.Text & "'"
      End If
      If TxtEmployeeName.Text <> "" Then
         tmp = tmp & " and MU.EmployeeName like '" & TxtEmployeeName.Text & "%'"
      End If
      
      tmp = tmp & " Order by LD.Login_Date "
   Else
      Call Grid_HeadSummary
      
      tmp = " Select LD.EmployeeID,MU.EmployeeName,MD.Designation,ME.EmployeeEmailID,ME.MobileNo,Count(LD.Login_Date) 'Login'"
      tmp = tmp & " From Mst_Employee ME"
      tmp = tmp & "  Right Join  Log_LoginDetail LD on ME.EmployeeNo=LD.EmployeeID"
      tmp = tmp & " LEFT Join Mst_Designation MD on MD.DesignationID=ME.DesignationID"
      tmp = tmp & " Inner Join Mst_Users MU on MU.EmployeeNo=LD.EmployeeID"
     
      If Gen_DBType = "MDB" Then
         tmp = tmp & " Where   (LD.Login_Date >= #" & TxtFromDate.Value & "# and LD.Login_Date <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM #) "
      Else
         tmp = tmp & " Where   (LD.Login_Date >= '" & TxtFromDate.Value & "' and LD.Login_Date <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM' ) "
      End If
      If TxtEmployeeID.Text <> "" Then
         tmp = tmp & " and LD.EmployeeID='" & TxtEmployeeID.Text & "'"
      End If
      If TxtEmployeeName.Text <> "" Then
         tmp = tmp & " and MU.EmployeeName like '" & TxtEmployeeName.Text & "%'"
      End If
      
      tmp = tmp & " Group by LD.EmployeeID,MU.EmployeeName,MD.Designation,ME.EmployeeEmailID,ME.MobileNo"
      tmp = tmp & " Order By LD.EmployeeID,MU.EmployeeName"
      
      
   End If


ElseIf Option2.Value = True Then
   If ChkSummary.Value = 0 Then
      Call Grid_HeadFormDetail
      
      tmp = "Select LFAD.LogID,MU.EmployeeNo,MU.EmployeeName,MM.MenuName,LFAD.Concurrency from Log_FormAccessDetail LFAD "
      tmp = tmp & " Inner join Mst_Users MU On LFAD.UserID = MU.UsersID"
      tmp = tmp & " Inner join Mst_MenuDetail MM On LFAD.MenuID = MM.MenuID"
      
      If Gen_DBType = "MDB" Then
         tmp = tmp & " Where   (LFAD.Concurrency >= #" & TxtFromDate.Value & "# and LFAD.Concurrency <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM #) "
      Else
         tmp = tmp & " Where   (LFAD.Concurrency >= '" & TxtFromDate.Value & "' and LFAD.Concurrency <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM' ) "
      End If
      If TxtEmployeeID.Text <> "" Then
         tmp = tmp & " and MU.EmployeeNo='" & TxtEmployeeID.Text & "'"
      End If
      If TxtEmployeeName.Text <> "" Then
         tmp = tmp & " and MU.EmployeeName like '" & TxtEmployeeName.Text & "%'"
      End If
     
      If TxtSMenuName.Text <> "" Then
         tmp = tmp & " and MM.MenuName like '" & TxtSMenuName.Text & "%'"
      End If
     
      tmp = tmp & " Order by LFAD.Concurrency "
      
   Else
      Call Grid_HeadFormSummary
      
      tmp = " Select MU.EmployeeNo,MU.EmployeeName,MM.MenuName,Count(MU.EmployeeNo) 'NoOfCount' From Log_FormAccessDetail LFAD"
      tmp = tmp & " Inner join Mst_Users MU On LFAD.UserID = MU.UsersID"
      tmp = tmp & " Inner join Mst_MenuDetail MM On LFAD.MenuID = MM.MenuID"
      
      If Gen_DBType = "MDB" Then
         tmp = tmp & " Where   (LFAD.Concurrency >= #" & TxtFromDate.Value & "# and LFAD.Concurrency <= #" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM #) "
      Else
         tmp = tmp & " Where   (LFAD.Concurrency >= '" & TxtFromDate.Value & "' and LFAD.Concurrency <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "  11:59:59 PM' ) "
      End If
      If TxtEmployeeID.Text <> "" Then
         tmp = tmp & " and MU.EmployeeNo='" & TxtEmployeeID.Text & "'"
      End If
      If TxtEmployeeName.Text <> "" Then
         tmp = tmp & " and MU.EmployeeName like '" & TxtEmployeeName.Text & "%'"
      End If
     
      If TxtSMenuName.Text <> "" Then
         tmp = tmp & " and MM.MenuName like '" & TxtSMenuName.Text & "%'"
      End If
      
      tmp = tmp & " Group By MM.MenuName,MU.EmployeeNo,MU.EmployeeName"
      
   End If
End If


Call TmpTable

If TmpRs.RecordCount > 0 Then

   For mRow = 1 To TmpRs.RecordCount
       For mCol = 0 To MSHFlexGrid1.Cols - 1
           If InStr(TmpRs.Fields(mCol), "/") > 0 Then
              If IsDate(TmpRs.Fields(mCol)) Then
                 MSHFlexGrid1.TextMatrix(mRow, mCol) = Format(TmpRs.Fields(mCol), "dd-MMM-yyyy hh:mm:ss AM/PM")
              Else
                 MSHFlexGrid1.TextMatrix(mRow, mCol) = IIf(IsNull(TmpRs.Fields(mCol)), "", TmpRs.Fields(mCol))
              End If
           Else
              MSHFlexGrid1.TextMatrix(mRow, mCol) = IIf(IsNull(TmpRs.Fields(mCol)), "", TmpRs.Fields(mCol))
           End If
           'MSHFlexGrid1.TextMatrix(mRow, mCol) = TmpRs.Fields(mCol)
       Next
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
   LblTotalRecord.Caption = TmpRs.RecordCount
Else
   MsgBox " No Record Found "
   Exit Sub
End If
MsgBox "Done !!!", vbInformation
End Sub
Public Sub Grid_Head()

With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 6
    .RowHeight(0) = 400
    .AllowUserResizing = flexResizeColumns
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(0, 0) = "Login Date "
    .TextMatrix(0, 1) = "Employee ID"
    .TextMatrix(0, 2) = "Employee Name"
    .TextMatrix(0, 3) = "Designation"
    .TextMatrix(0, 4) = "Email ID"
    .TextMatrix(0, 5) = "Mobile No"
    .ColWidth(0) = 2600
    .ColWidth(1) = 1500
    .ColWidth(2) = 3500
    .ColWidth(3) = 1800
    .ColWidth(4) = 3000
    .ColWidth(5) = 1800
End With
End Sub
Public Sub Grid_HeadSummary()

With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 6
    .RowHeight(0) = 700
    .AllowUserResizing = flexResizeColumns
    .Font.Size = 10
    .WordWrap = True
    'Employee id , Employee name, Department , Email id , mobile no, count of (Date)
    .TextMatrix(0, 0) = "Employee ID"
    .TextMatrix(0, 1) = "Employee Name"
    .TextMatrix(0, 2) = "Designation"
    .TextMatrix(0, 3) = "Email ID"
    .TextMatrix(0, 4) = "Mobile No"
    .TextMatrix(0, 5) = "No of Times Log in"
    .ColWidth(0) = 2600
    .ColWidth(1) = 3500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1800
    .ColWidth(4) = 3000
    .ColWidth(5) = 1800
End With
End Sub

Public Sub Grid_HeadFormDetail()

With MSHFlexGrid1
    
    .Clear
    .Rows = 2
    .Cols = 5
    .RowHeight(0) = 400
    .AllowUserResizing = flexResizeColumns
    .Font.Size = 10
    .WordWrap = True
    
    .TextMatrix(0, 0) = "Log ID "
    .TextMatrix(0, 1) = "Employee ID"
    .TextMatrix(0, 2) = "Employee Name"
    .TextMatrix(0, 3) = "Menu Name"
    .TextMatrix(0, 4) = "Concurrency"

    .ColWidth(0) = 2600
    .ColWidth(1) = 1500
    .ColWidth(2) = 3500
    .ColWidth(3) = 1800
    .ColWidth(4) = 3000
    
    
End With

End Sub
Public Sub Grid_HeadFormSummary()

With MSHFlexGrid1
    
    .Clear
    .Rows = 2
    .Cols = 4
    .RowHeight(0) = 400
    .AllowUserResizing = flexResizeColumns
    .Font.Size = 10
    .WordWrap = True
    
    .TextMatrix(0, 0) = "Employee ID"
    .TextMatrix(0, 1) = "Employee Name"
    .TextMatrix(0, 2) = "Menu Name"
    .TextMatrix(0, 3) = "No of times access"

    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 3500
    .ColWidth(3) = 1800
End With

End Sub


Private Sub Form_Load()
Call Grid_Head

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Call CmdClear_Click

End Sub

Private Sub Option1_Click()

Call Grid_Head

TxtSMenuName.Visible = False
Label6.Visible = False
'ChkSummary.Visible = False

If Option2.Value = True Then
   ChkSummary.Visible = True
   ChkSummary.Value = 0
   TxtSMenuName.Visible = True
   Label6.Visible = True
   Call Grid_HeadFormDetail
End If



End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub
