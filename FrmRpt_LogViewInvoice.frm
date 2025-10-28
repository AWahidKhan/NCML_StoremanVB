VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_LogViewInvoice 
   Caption         =   "Log View Invoice Report"
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
      Height          =   7875
      Left            =   120
      TabIndex        =   12
      Top             =   1905
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   7485
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   13
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
      Height          =   1725
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15000
      Begin VB.ComboBox CmbPrinted 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1250
         Width           =   5955
      End
      Begin VB.TextBox TxtEntryNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   12480
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TxtInvoiceNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10080
         TabIndex        =   5
         Top             =   600
         Width           =   2350
      End
      Begin VB.ComboBox CmbState 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   3435
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
         Left            =   12480
         TabIndex        =   8
         Top             =   1110
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
         Height          =   495
         Left            =   11265
         TabIndex        =   0
         Top             =   1110
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
         Height          =   495
         Left            =   13695
         TabIndex        =   9
         Top             =   1110
         Width           =   1215
      End
      Begin VB.ComboBox CmbLocation 
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
         Left            =   3575
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox CmbInvoiceType 
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
         ItemData        =   "FrmRpt_LogViewInvoice.frx":0000
         Left            =   6120
         List            =   "FrmRpt_LogViewInvoice.frx":0016
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   3900
      End
      Begin VB.Label Label5 
         Caption         =   "Printed By"
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
         Left            =   2550
         TabIndex        =   17
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Entry No "
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
         Left            =   13282
         TabIndex        =   16
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Invoice No "
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
         Left            =   10753
         TabIndex        =   15
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "State"
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
         Left            =   1410
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   4415
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Invoice Type"
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
         Left            =   7500
         TabIndex        =   10
         Top             =   240
         Width           =   1170
      End
   End
End
Attribute VB_Name = "FrmRpt_LogViewInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mSMPrefix, mEmployeeNo, mStatus As Variant

Private Sub CmbInvoiceType_LostFocus()

If CmbInvoiceType.Text = "" Then
   Exit Sub
End If

If LCase(CmbInvoiceType) <> "nnnr" And LCase(CmbInvoiceType) <> "rejected" And LCase(CmbInvoiceType) <> "reservation" And LCase(CmbInvoiceType) <> "7days" And LCase(CmbInvoiceType.Text) <> "adhoc" And LCase(CmbInvoiceType.Text) <> "dp" And LCase(CmbInvoiceType) <> "commtrack" And LCase(CmbInvoiceType) <> "rrn" Then
   MsgBox "Invalid Invoice Type"
   Exit Sub
End If

If LCase(CmbInvoiceType) = "nnnr" Then
   mStatus = "NNNR"
ElseIf LCase(CmbInvoiceType) = "reservation" Then
   mStatus = "Reservation"
ElseIf LCase(CmbInvoiceType) = "7days" Then
   mStatus = "SD"
ElseIf LCase(CmbInvoiceType) = "rejected" Then
   mStatus = "RJ"
ElseIf LCase(CmbInvoiceType) = "adhoc" Then
   mStatus = "ADHOC"
ElseIf LCase(CmbInvoiceType) = "dp" Then
   mStatus = "DP"
ElseIf LCase(CmbInvoiceType) = "commtrack" Then
   mStatus = "COMM"
ElseIf LCase(CmbInvoiceType) = "rrn" Then
   mStatus = "RRN"
End If

End Sub

Private Sub CmbLocation_GotFocus()
If mStateID = "" Then Exit Sub

tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("where " & Gen_WcLocation & " And StateID=" & mStateID & "")
End If
CmbLocation.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location.Fields("LocationID")
mSMPrefix = RsMst_Location.Fields("SM_Prefix")
End Sub



Private Sub CmbPrinted_GotFocus()
tmp = CmbPrinted.Text
Call TableMst_Employee("Where Flag='Y'")
CmbPrinted.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbPrinted.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Employee.RecordCount
    CmbPrinted.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbPrinted.Text = tmp
End Sub

Private Sub CmbPrinted_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbPrinted.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbPrinted, "~")
EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))


RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbPrinted.SetFocus
   Exit Sub
End If

mEmployeeNo = RsMst_Employee!EmployeeNo
End Sub


Private Sub CmbState_GotFocus()
tmp = CmbState.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub


Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   CmbLocation.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If
If RsMst_State!StateName <> CmbState.Text Then
   CmbLocation.Text = ""
End If

mStateID = RsMst_State!StateID
End Sub


Private Sub CmdClear_Click()
Call Grid_Head1
CmbState.Text = ""
CmbLocation.Text = ""
CmbInvoiceType.Text = ""
TxtInvoiceNo.Text = ""
TxtEntryNo.Text = ""
CmbPrinted.Text = ""
CmdExcel.Enabled = False
ProgressBar1.Value = 0
End Sub
Public Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 7
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "State"
MSHFlexGrid1.TextMatrix(0, 1) = "Location"
MSHFlexGrid1.TextMatrix(0, 2) = "Invoice Type"
MSHFlexGrid1.TextMatrix(0, 3) = "Invoice No"
MSHFlexGrid1.TextMatrix(0, 4) = "File Name(Entry No)"
MSHFlexGrid1.TextMatrix(0, 5) = "Printed By"
MSHFlexGrid1.TextMatrix(0, 6) = "Printed On"

MSHFlexGrid1.ColWidth(0) = 2000
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 4000
MSHFlexGrid1.ColWidth(6) = 2000

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "LogViewInvoiceReport.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "LogViewInvoiceReport.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 1
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
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub

Private Sub CmdGetReport_Click()

Dim SR, SSR As Variant
Dim tmp1 As String
Dim wc As String
wc = ""


Call Grid_Head1
MsgBox ("Wait till next message")
ProgressBar1.Value = 0

tmp = "Select MS.StateName,ML.LocationName,LPH.DocType,LPH.DocNo,LPH.FileName,ME.EmployeeName+'~'+ME.EmployeeNo as EmployeeName,LPH.Concurrency"
tmp = tmp & " From Log_Print_History LPH"
tmp = tmp & " Inner Join Mst_Location ML on LPH.SM_Prefix=ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join Mst_Employee ME on LPH.UserId=ME.EmployeeNo  "

wc = " Where LPH.DocType in ('NNNR','Reservation','SD','RJ','ADHOC','DP','COMM','RRN')"


If CmbState.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
End If

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbInvoiceType.Text <> "" Then
   wc = GenWc(wc, mStatus, "LPH.DocType", "S", "N")
End If

If TxtInvoiceNo.Text <> "" Then
   wc = GenWc(wc, TxtInvoiceNo, "LPH.DocNo", "N", "N")
End If

If TxtEntryNo.Text <> "" Then
   wc = GenWc(wc, TxtEntryNo, "LPH.FileName", "S", "N")
End If

If CmbPrinted.Text <> "" Then
   wc = GenWc(wc, mEmployeeNo, "ME.EmployeeNo", "S", "N")
End If


tmp = tmp & wc
tmp = tmp & " Order By MS.StateName,ML.LocationName,LPH.DocNo,LPH.Concurrency"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs.RecordCount + 2
   For i = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(TmpRs!DocType), "", TmpRs!DocType)
       MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs!DocNo), "", TmpRs!DocNo)
       MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs!FileName), "", TmpRs!FileName)
       MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs!EmployeeName), "", TmpRs!EmployeeName)
       MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRs!Concurrency), "", Format(TmpRs!Concurrency, "dd-mmm-yyyy hh:mm"))
       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   CmdExcel.Enabled = True
   MsgBox ("Done!!!")
Else
   MsgBox "No record found !!!"
   Call Grid_Head1
   Exit Sub
End If
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
CmdClear_Click
End Sub


