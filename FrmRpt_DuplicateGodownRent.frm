VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_DuplicateGodownRent 
   Caption         =   "Duplicate Godown Rent"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.OptionButton Opt2 
         Caption         =   "Detail View"
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
         Left            =   7440
         TabIndex        =   13
         Top             =   413
         Width           =   1455
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Simple View"
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
         Left            =   5520
         TabIndex        =   12
         Top             =   413
         Value           =   -1  'True
         Width           =   1695
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
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1935
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
         ItemData        =   "FrmRpt_DuplicateGodownRent.frx":0000
         Left            =   3720
         List            =   "FrmRpt_DuplicateGodownRent.frx":0002
         TabIndex        =   8
         Top             =   360
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
         Height          =   375
         Left            =   12375
         TabIndex        =   3
         Top             =   600
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
         Left            =   11175
         TabIndex        =   2
         Top             =   600
         Width           =   1215
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
         Left            =   13590
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         TabIndex        =   11
         Top             =   413
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   3120
         TabIndex        =   10
         Top             =   413
         Width           =   495
      End
      Begin VB.Label Label2 
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
         Left            =   11190
         TabIndex        =   5
         Top             =   255
         Width           =   2295
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   13590
         TabIndex        =   4
         Top             =   255
         Width           =   90
      End
   End
End
Attribute VB_Name = "FrmRpt_DuplicateGodownRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mMonth As Variant
Dim mYear As Variant
Private Sub CmbMonth_LostFocus()
Call Grid_Head1
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
End Sub

Private Sub CmdClear_Click()
CmbMonth.Text = ""
CmbYear.Text = ""
Opt1.Value = True
Opt2.Value = False
LblTotalRecords.Caption = "-"
Call Grid_Head1
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_DuplicateGodownRent.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_DuplicateGodownRent.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 1 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
               If Opt2.Value = True And c = 4 Then
                  oWS.Cells(mRow, c + 1) = "'" & (MSHFlexGrid1.TextMatrix(I, c))
               Else
                  oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "dd-mmm-yyyy")
               End If
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()
If CmbMonth.Text = "" Then
   MsgBox "Please Select Month !!"
   CmbMonth.SetFocus
   Exit Sub
End If

If CmbYear.Text = "" Then
   MsgBox " Please Select Year !!"
   CmbYear.SetFocus
   Exit Sub
End If
'2010-07-01

If Opt1.Value = True Then
   Call Grid_Head1
   tmp = "Select TG.ClaimID,ML.LessorName,TG.ClaimDate,TG.ExpenseFrom,TG.ExpenseTo,TG.Amount,TG.ApprovalRemarks,TG.Flag"
   tmp = tmp & " From Txn_GodownRent TG"
   tmp = tmp & " Inner Join Mst_Lessor ML On ML.LessorID=TG.LessorID"
   tmp = tmp & " Where Month(TG.ExpenseFrom) = " & mMonth & " and Year(TG.ExpenseFrom)=" & Val(CmbYear.Text) & " and TG.Flag <> 'R'"
   tmp = tmp & " And TG.LessorID in ("
   tmp = tmp & "    select LessorID from Txn_GodownRent"
   tmp = tmp & "    where Month(ExpenseFrom) =" & mMonth & " and Year(ExpenseFrom)=" & Val(CmbYear.Text) & " and Flag <> 'R'"
   tmp = tmp & "    group by lessorid"
   tmp = tmp & "    Having Count(lessorid) > 1"
   tmp = tmp & "    ) Order By TG.ClaimID,TG.ClaimDate,TG.ExpenseFrom,TG.ExpenseTo,ML.LessorName "
ElseIf Opt2.Value = True Then
   Call Grid_Head2
   tmp = " Select TGR.ClaimID,MLe.LessorName,ML.LocationName,MC.CMPName,MG.GodownNo ,TG.ClaimDate,TG.ExpenseFrom,TG.ExpenseTo,TG.Amount,TG.ApprovalRemarks,TG.Flag"
   tmp = tmp & " From Txn_GodownRent TG"
   tmp = tmp & " Inner Join Mst_Lessor MLe On MLe.LessorID=TG.LessorID"
   tmp = tmp & " Inner Join Txn_GodownRentDetail TGR On TGR.ClaimID=TG.ClaimID"
   tmp = tmp & " Inner join Mst_Godown MG On MG.GodownID=TGR.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MC ON MC.CMPID=MG.CMPID"
   tmp = tmp & " Inner Join  Mst_Location ML On ML.LocationID = MC.LocationID"
   tmp = tmp & " Where Month(TG.ExpenseFrom) = " & mMonth & " And Year(TG.ExpenseFrom)=" & Val(CmbYear.Text) & " and TG.Flag <> 'R'"
   tmp = tmp & " And TG.LessorID in ("
   tmp = tmp & "    select LessorID from Txn_GodownRent"
   tmp = tmp & "    where Month(ExpenseFrom) =" & mMonth & " and Year(ExpenseFrom)=" & Val(CmbYear.Text) & "  and Flag <> 'R'"
   tmp = tmp & "    group by lessorid"
   tmp = tmp & "    Having Count(LessorID) > 1 ) Order By ML.LocationName,TGR.ClaimID,MLe.LessorName"
End If

Call TmpTable
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       For c = 0 To MSHFlexGrid1.Cols - 1
           If InStr(TmpRs.Fields(c), "/") = 2 Then
              If IsDate(TmpRs.Fields(c)) Then
                 If Opt2.Value = True And c = 4 Then
                    MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
                 Else
                    MSHFlexGrid1.TextMatrix(I, c) = Format(TmpRs.Fields(c), "dd-MMM-yyyy")
                 End If
              Else
                 MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
              End If
           Else
              MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
           End If
       Next
       If LCase(MSHFlexGrid1.TextMatrix(I, MSHFlexGrid1.Cols - 1)) = "a" Then
          MSHFlexGrid1.TextMatrix(I, MSHFlexGrid1.Cols - 1) = "Approved"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, MSHFlexGrid1.Cols - 1)) = "p" Then
         MSHFlexGrid1.TextMatrix(I, MSHFlexGrid1.Cols - 1) = "Pending"
       ElseIf LCase(MSHFlexGrid1.TextMatrix(I, MSHFlexGrid1.Cols - 1)) = "s" Then
         MSHFlexGrid1.TextMatrix(I, MSHFlexGrid1.Cols - 1) = "Settled"
       End If
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
Else
   MsgBox "No Record Found !!!"
End If
LblTotalRecords.Caption = TmpRs.RecordCount
End Sub
Private Sub Grid_Head1()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 8
     .AllowUserResizing = flexResizeColumns
     .TextMatrix(0, 0) = "Claim ID"
     .TextMatrix(0, 1) = "Lessor Name"
     .TextMatrix(0, 2) = "Claim Date"
     .TextMatrix(0, 3) = "Expense From"
     .TextMatrix(0, 4) = "Expense To"
     .TextMatrix(0, 5) = "Amount"
     .TextMatrix(0, 6) = "Approval Remarks"
     .TextMatrix(0, 7) = "Status"
     .ColWidth(0) = 900
     .ColWidth(1) = 3000
     .ColWidth(2) = 1400
     .ColWidth(3) = 1400
     .ColWidth(4) = 1400
     .ColWidth(5) = 1200
     .ColWidth(6) = 4000
     .ColWidth(7) = 1100
     .RowHeight(0) = 500
     .Font.Size = 10
End With
End Sub
Private Sub Grid_Head2()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 11
     .AllowUserResizing = flexResizeColumns
     .TextMatrix(0, 0) = "Claim ID"
     .TextMatrix(0, 1) = "Lessor Name"
     .TextMatrix(0, 2) = "Location Name"
     .TextMatrix(0, 3) = "CMP Name"
     .TextMatrix(0, 4) = "Godown No"
     .TextMatrix(0, 5) = "Claim Date"
     .TextMatrix(0, 6) = "Expense From"
     .TextMatrix(0, 7) = "Expense To"
     .TextMatrix(0, 8) = "Amount"
     .TextMatrix(0, 9) = "Approval Remarks"
     .TextMatrix(0, 10) = "Status"
     .ColWidth(0) = 900
     .ColWidth(1) = 3000
     .ColWidth(2) = 2200
     .ColWidth(3) = 2200
     .ColWidth(4) = 900
     .ColWidth(5) = 1400
     .ColWidth(6) = 1400
     .ColWidth(7) = 1400
     .ColWidth(8) = 1200
     .ColWidth(9) = 5300
     .ColWidth(10) = 1000
     .RowHeight(0) = 500
     .Font.Size = 10
End With
End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
LblTotalRecords.Caption = "-"
Call Grid_Head1

For I = 1 To 12
   CmbMonth.AddItem MonthName(I), I - 1
Next
End Sub

Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For I = 2010 To Year(Now)
    CmbYear.AddItem I
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
Call Grid_Head1

If CmbYear.Text = "" Then Exit Sub
Dim flg As Variant
flg = "N"
For I = 2010 To Year(Now)
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
End Sub

Private Sub Opt1_Click()
Call Grid_Head1
LblTotalRecords.Caption = "-"
End Sub

Private Sub Opt2_Click()
Call Grid_Head2
LblTotalRecords.Caption = "-"
End Sub
