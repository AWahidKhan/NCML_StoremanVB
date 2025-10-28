VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CompFumgation 
   Caption         =   "Comparative Fumigation"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   10395
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8160
      Left            =   120
      TabIndex        =   12
      Top             =   1695
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7830
         Left            =   120
         TabIndex        =   13
         Top             =   195
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13811
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1710
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   15000
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
         Left            =   6705
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   705
         Width           =   3810
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
         Left            =   3135
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   705
         Width           =   3510
      End
      Begin VB.ComboBox CmbCMP 
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
         Left            =   10575
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   705
         Width           =   4290
      End
      Begin VB.ComboBox CmbGodown 
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
         Left            =   1620
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1170
         Width           =   5025
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
         Left            =   8017
         TabIndex        =   7
         Top             =   1170
         Width           =   1245
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
         Left            =   6705
         TabIndex        =   6
         Top             =   1170
         Width           =   1290
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
         Left            =   9285
         TabIndex        =   8
         Top             =   1170
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   330
         Left            =   1620
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   60686337
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
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
         Format          =   60686337
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label7 
         Caption         =   "Next Fumigation Date To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1597
         TabIndex        =   19
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "Godown No"
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
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   12203
         TabIndex        =   17
         Top             =   405
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4658
         TabIndex        =   16
         Top             =   405
         Width           =   465
      End
      Begin VB.Label Label5 
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
         Left            =   10575
         TabIndex        =   15
         Top             =   1230
         Width           =   2325
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
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
         Left            =   12945
         TabIndex        =   14
         Top             =   1230
         Width           =   75
      End
      Begin VB.Label Label3 
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
         Left            =   8228
         TabIndex        =   11
         Top             =   405
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Next Fumigation Date From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   1500
      End
   End
End
Attribute VB_Name = "FrmRpt_CompFumgation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mGodownID As Variant


Private Sub CmbCMP_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Select Location."
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCMP.Text

Call TableMst_CMP("Where LocationID =" & mLocationID & "")

CmbCMP.Clear

If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No CMP found  !!!!! "
   CmbCMP.SetFocus
   Exit Sub
End If
RsMst_CMP.MoveFirst
For I = 1 To RsMst_CMP.RecordCount
    CmbCMP.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMP.Text = tmp
End Sub

Private Sub CmbCMP_LostFocus()
If tmp <> CmbCMP.Text Then
   CmbGodown.Clear
   CmbGodown.Text = ""
End If
If CmbCMP.Text = "" Then
   mCMPID = 0
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"


If RsMst_CMP.EOF Then
   MsgBox "Invalid CMP Selection "
   CmbCMP.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodown_GotFocus()
If CmbCMP.Text = "" Then
   MsgBox "Select CMP."
   CmbCMP.SetFocus
   Exit Sub
End If
'
temp = CmbGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID & " And CloseDate Is Null")
CmbGodown.Clear

If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Godown found  !!!!! "
   CmbGodown.SetFocus
   Exit Sub
End If
RsMst_Godown.MoveFirst
For I = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = temp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then
   mGodownID = 0
   Exit Sub
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"

If RsMst_Godown.EOF Then
   MsgBox "Invalid Godown Selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Select State."
   CmbStateID.SetFocus
   Exit Sub
End If
tmp = CmbLocation.Text
Call TableMst_Location(" WHERE StateID = " & mStateID)

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub

Private Sub CmbLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbLocation")
End If
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text

If Gen_WcState <> "" Then
   Call TableMst_State(" WHERE " & Gen_WcState)
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

Private Sub CmdClear_Click()
Call Grid_Head
Label22.Caption = 0
CmbLocation.Text = ""
CmdExcel.Enabled = False
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CompFumgation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CompFumgation.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MMM/dd/yyyy")
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
If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "Form date Should be Less Than To date."
   TxtFromDate.SetFocus
   Exit Sub
End If

Label22.Caption = ""

tmp = "SELECT MS.StateName, ML.LocationName, MC.CMPName, MG.GodownNo, Mcm.Commodity, TFD.StackNo, MV.VendorName,"
tmp = tmp & " SUM(TFD.BalanceQty) As Bal, SUM(TFD.ActualQty) As Act, (SUM(TFD.BalanceQty)-SUM(TFD.ActualQty)) as balQty, "
tmp = tmp & " TF.Rate, TF.Amount, TF.FumigationDate, TF.NextFumigationDate, TF.Remark"
tmp = tmp & " FROM Txn_Fumigation TF"
tmp = tmp & " INNER JOIN Txn_FumigationDetail TFD ON TFD.FumigationID=TF.FumigationID"
tmp = tmp & " INNER JOIN Mst_Godown MG ON MG.GodownID=TF.GodownID"
tmp = tmp & " INNER JOIN Mst_CMP MC ON MC.CMPID=MG.CMPID"
tmp = tmp & " INNER JOIN Mst_Vendor MV ON MV.VendorID=TF.VendorID"
tmp = tmp & " INNER JOIN Mst_Location ML ON ML.LocationID=MC.LocationID"
tmp = tmp & " INNER JOIN Mst_State MS ON MS.StateID=ML.StateID"
tmp = tmp & " INNER JOIN Mst_Commodity MCm ON MCm.CommodityID=TFD.CommodityID"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TF.NextFumigationDate >= #" & Format(TxtFromDate, Gen_DatFormat) & "# And TF.NextFumigationDate <= #" & Format(TxtToDate, Gen_DatFormat) & "#"
Else
   tmp = tmp & " Where TF.NextFumigationDate >= '" & Format(TxtFromDate, Gen_DatFormat) & "' And TF.NextFumigationDate <= '" & Format(TxtToDate, Gen_DatFormat) & "'"
End If

If CmbStateID.Text <> "" Then
   tmp = tmp & " And MS.StateName='" & CmbStateID.Text & "'"
Else
   If Gen_WcState <> "" Then
      tmp = tmp & " And MS." & Gen_WcState
   End If
End If

If CmbLocation.Text <> "" Then
   tmp = tmp & " And ML.LocationName='" & CmbLocation.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

If CmbCMP.Text <> "" Then
   tmp = tmp & " And MC.CMPName='" & CmbCMP.Text & "'"
End If

If CmbGodown.Text <> "" Then
   tmp = tmp & " And MG.GodownNo='" & CmbGodown.Text & "'"
End If

tmp = tmp & " GROUP BY MS.StateName, MC.CMPName, ML.LocationName, MG.GodownNo, Mcm.Commodity, TFD.StackNo, MV.VendorName,"
tmp = tmp & " TF.FumigationDate , TF.NextFumigationDate, TF.Rate, TF.Amount, TF.Remark"
'tmp = tmp & "  order by LocationName,CxTFNo,CxTFDate"

Call TmpTable

Call dispdata

MsgBox ("Done!!!")
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Frame1.Enabled = False
   Exit Sub
End If
Call Grid_Head
Frame1.Enabled = True
Call CmdClear_Click
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 15

MSHFlexGrid1.TextMatrix(0, 0) = "State"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 4) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 5) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 6) = "Vendor Name"
MSHFlexGrid1.TextMatrix(0, 7) = "Total Qty in Godown"
MSHFlexGrid1.TextMatrix(0, 8) = "Fumigated Qty"
MSHFlexGrid1.TextMatrix(0, 9) = "Balance Qty to be Fumigated"
MSHFlexGrid1.TextMatrix(0, 10) = "Rate"
MSHFlexGrid1.TextMatrix(0, 11) = "Amount to be Paid"
MSHFlexGrid1.TextMatrix(0, 12) = "Fumigation Date"
MSHFlexGrid1.TextMatrix(0, 13) = "Due Date for Fumigation"
MSHFlexGrid1.TextMatrix(0, 14) = "Remarks"


MSHFlexGrid1.ColWidth(0) = 1500
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 2000
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200
MSHFlexGrid1.ColWidth(14) = 2000

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True
End Sub

Public Sub dispdata()
Dim x As Double
Call Grid_Head
MsgBox ("Wait till Next Message")
x = 1

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       DoEvents
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(I, 12) = Format(MSHFlexGrid1.TextMatrix(I, 12), "dd-mmm-yyyy")
       MSHFlexGrid1.TextMatrix(I, 13) = Format(MSHFlexGrid1.TextMatrix(I, 13), "dd-mmm-yyyy")
       TmpRs.MoveNext
   Next
   Label22.Refresh
   Label22.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If
End Sub

