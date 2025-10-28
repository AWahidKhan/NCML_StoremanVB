VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_LogInvoiceDeletion 
   Caption         =   "Invoice Deletion Log"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8955
      Left            =   120
      TabIndex        =   6
      Top             =   690
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   8565
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15000
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
         Height          =   360
         Left            =   8055
         TabIndex        =   3
         Top             =   210
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
         Height          =   360
         Left            =   5625
         TabIndex        =   2
         Top             =   210
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
         Height          =   360
         Left            =   6840
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   1215
         TabIndex        =   9
         Top             =   203
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
         Left            =   3600
         TabIndex        =   10
         Top             =   203
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
         Left            =   2760
         TabIndex        =   12
         Top             =   263
         Width           =   735
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
         Left            =   135
         TabIndex        =   11
         Top             =   263
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "--"
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
         Left            =   12720
         TabIndex        =   5
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   10320
         TabIndex        =   4
         Top             =   240
         Width           =   2280
      End
   End
End
Attribute VB_Name = "FrmRpt_LogInvoiceDeletion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Grid_Head()
Label5.Caption = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
With MSHFlexGrid1
   .Clear
   .Cols = 8
   .Rows = 2
      
   .TextMatrix(0, 0) = "Document Type"
   .TextMatrix(0, 1) = "Document ID"
   .TextMatrix(0, 2) = "Location"
   .TextMatrix(0, 3) = "Client"
   .TextMatrix(0, 4) = "Invoice Amount"
   .TextMatrix(0, 5) = "Deletion Remarks"
   .TextMatrix(0, 6) = "Deleted By"
   .TextMatrix(0, 7) = "Deleted On"
   
   .ColWidth(0) = 2000
   .ColWidth(1) = 1500
   .ColWidth(2) = 1500
   .ColWidth(3) = 1500
   .ColWidth(4) = 1200
   .ColWidth(5) = 3800
   .ColWidth(6) = 2800
   .ColWidth(7) = 1400
   
   .RowHeight(0) = 700
   .WordWrap = True
End With
End Sub

Private Sub CmbMenuName_GotFocus()
tmp = CmbMenuName.Text
Call TableMst_MenuDetail

CmbMenuName.Clear
If RsMst_MenuDetail.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_MenuDetail.RecordCount
   CmbMenuName.AddItem RsMst_MenuDetail!MenuName
   RsMst_MenuDetail.MoveNext
Next
CmbMenuName.Text = tmp
End Sub

Private Sub CmbMenuName_LostFocus()
If Trim(CmbMenuName.Text) = "" Then Exit Sub

RsMst_MenuDetail.MoveFirst
RsMst_MenuDetail.Find "MenuName = '" & CmbMenuName.Text & "'"
If RsMst_MenuDetail.EOF Then
   MsgBox "Invalid selection "
   CmbMenuName.SetFocus
   Exit Sub
End If
mMenuID = RsMst_MenuDetail!MenuID
End Sub

Private Sub CmbMenuType_LostFocus()
If Trim(CmbMenuType.Text) = "" Then Exit Sub
If LCase(CmbMenuType.Text) <> "admin" And LCase(CmbMenuType.Text) <> "master" And LCase(CmbMenuType.Text) <> "transaction" And LCase(CmbMenuType.Text) <> "report" And LCase(CmbMenuType.Text) <> "cog update" Then
   MsgBox "Invalid Selection!!!"
   CmbMenuType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbUserRole_GotFocus()
tmp = CmbUserRole.Text

Call TableMst_UserRole

CmbUserRole.Clear
If RsMst_UserRole.RecordCount = 0 Then
   MsgBox "No User Role found"
   CmbUserRole.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_UserRole.RecordCount
   CmbUserRole.AddItem RsMst_UserRole!GroupName
   RsMst_UserRole.MoveNext
Next
CmbUserRole.Text = tmp
End Sub

Private Sub CmbUserRole_LostFocus()
If Trim(CmbUserRole.Text) = "" Then Exit Sub

RsMst_UserRole.MoveFirst
RsMst_UserRole.Find "GroupName = '" & CmbUserRole.Text & "'"
If RsMst_UserRole.EOF Then
   MsgBox "Invalid selection!!!"
   CmbUserRole.SetFocus
   Exit Sub
End If

mUserRoleID = RsMst_UserRole!UserGRoupID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_LogInvoiceDeletion.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_LogInvoiceDeletion.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
       oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
Dim mDate As Date
Call Grid_Head
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


MsgBox "Please wait till next Message!!!"
mDate = "09/17/2013"

If TxtFromDate.Value >= CDate(mDate) Then
   tmp = " Select DocumentType,DcoumentID,L.LocationName,c.Clientname ,a.Amount,a.DeletetionRemark,u.EmployeeName+'~'+a.DeletedBy 'DeletedBy',a.DeletedOn"
   tmp = tmp & " From Log_InvoiceDeletetion a"
   tmp = tmp & " Inner Join Mst_Location l on a.sm_prefix = l.sm_prefix"
   tmp = tmp & " Inner Join Mst_Client c on a.clientIDRow = c.clientIDRow"
   tmp = tmp & " Left Join Mst_Users u on a.DeletedBy = u.EmployeeNo"
   tmp = tmp & " Where DeletedOn >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And DeletedOn <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
   tmp = tmp & " Order By DocumentType,L.LocationName,DcoumentID"
Else
   tmp = "Select "
   tmp = tmp & " LogTbl.DocumentType, LogTbl.DcoumentID, LogTbl.LocationName, LogTbl.Clientname,"
   tmp = tmp & " LogTbl.Amount , LogTbl.DeletetionRemark, LogTbl.DeletedBy, LogTbl.DeletedOn"
   tmp = tmp & " From"
   tmp = tmp & " ("
   tmp = tmp & " Select DocumentType,DcoumentID,L.LocationName,c.Clientname ,a.Amount,a.DeletetionRemark,u.EmployeeName+'~'+a.DeletedBy 'DeletedBy',a.DeletedOn"
   tmp = tmp & " From Log_InvoiceDeletetion a"
   tmp = tmp & " Inner Join Mst_Location l on a.sm_prefix = l.sm_prefix"
   tmp = tmp & " Inner Join Mst_Client c on a.clientIDRow = c.clientIDRow"
   tmp = tmp & " Left Join Mst_Users u on a.DeletedBy = u.EmployeeNo"
   tmp = tmp & " Where a.DeletedOn >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "' And DeletedOn <= '" & Format(TxtToDate.Value, Gen_DatFormat) & "'"
   tmp = tmp & " Union"
   tmp = tmp & " Select 'NNNR' as DocumentType,P.InvoiceNo as 'DcoumentID', L.LocationName,c.Clientname, P.Amount,'' as DeletetionRemark, u.EmployeeName+'~'+p.UpdatedBy 'DeletedBy',p.UpdatedOn as DeletedOn"
   tmp = tmp & " From  Log_Txn_ProInvoice P"
   tmp = tmp & " Inner Join Mst_Location l on P.sm_prefix = l.sm_prefix"
   tmp = tmp & " Inner Join Mst_Client c on P.clientIDRow = c.clientIDRow"
   tmp = tmp & " Left Join Mst_Users u on p.Updatedby = u.EmployeeNo"
   tmp = tmp & " Where p.UpdatedOn < ='" & Format(mDate, Gen_DatFormat) & "' And p.UpdatedOn >= '" & Format(TxtFromDate.Value, Gen_DatFormat) & "') LogTbl"
   tmp = tmp & " Order By LogTbl.DocumentType,LogTbl.LocationName,LogTbl.DcoumentID"
End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 3
   CmdExcel.Enabled = True
   Label5.Caption = TmpRs.RecordCount
   With MSHFlexGrid1
      .Rows = .Rows + TmpRs.RecordCount
      For I = 1 To TmpRs.RecordCount
         For C = 0 To TmpRs.Fields.Count - 1
             If C = 7 Then
                MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", Format(TmpRs.Fields(C), "dd-MMM-yyyy"))
             Else
                MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
             End If
         Next
         If LCase(MSHFlexGrid1.TextMatrix(I, 0)) = "rj" Then
            MSHFlexGrid1.TextMatrix(I, 0) = "Rejected"
         ElseIf LCase(MSHFlexGrid1.TextMatrix(I, 0)) = "sd" Then
            MSHFlexGrid1.TextMatrix(I, 0) = "Seven Days"
         End If
         'MSHFlexGrid1.TextMatrix(i, 7) = Format(MSHFlexGrid1.TextMatrix(i, 7), "dd-mmm-yyyy hh:mmm:ss")
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
   End With
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
   MsgBox "No record found!!!"
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
Call CmdClear_Click
End Sub


