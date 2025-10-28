VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_PendingPrint 
   Caption         =   "Print Pending List"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   135
      TabIndex        =   2
      Top             =   120
      Width           =   15015
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   9870
         TabIndex        =   11
         Top             =   285
         Width           =   1215
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   6015
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   795
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
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   255
         Width           =   4455
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         Height          =   375
         Left            =   7440
         TabIndex        =   4
         Top             =   285
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   8655
         TabIndex        =   3
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Prefix"
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
         Left            =   5400
         TabIndex        =   10
         Top             =   315
         Visible         =   0   'False
         Width           =   600
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
         Left            =   120
         TabIndex        =   8
         Top             =   308
         Width           =   855
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
         Left            =   11160
         TabIndex        =   6
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label2 
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
         Left            =   13560
         TabIndex        =   5
         Top             =   360
         Width           =   75
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   975
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_PendingPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClear_Click()
CmbLocation.Text = ""
TxtPreFix = ""
Call Grid_Head
Label2.Caption = ""
CmdExcel.Enabled = False
End Sub

Private Sub CmdGetReport_Click()

Dim wc As Variant
Call Grid_Head
MsgBox "Wait till next message!!!"
wc = ""
If CmbLocation.Text <> "" Then
   wc = " And L.LocationName = '" & CmbLocation.Text & "' "
Else
   If Gen_WcLocation <> "" Then
      wc = " And SM_PreFix in (Select Distinct SM_PreFix From Mst_Location Where " & Gen_WcLocation & ")"
      wc = " And l." & Gen_WcLocation
   End If
End If

'tmp = "SELECT 'WHR' AS TT_Type, WHRNo, 'WHR' AS TxnType, SM_Prefix, WHRDate, Flag " & _
'      " From dbo.Txn_WHR_Detail " & _
'      " WHERE (Flag IS NULL) " & wc & _
'      " Union All " & _
'      " SELECT     'CxTF' AS TT_Type, CxTFNo, TxnType, SM_Prefix, CxTFDate, Flag " & _
'      " From dbo.Txn_CxTF_Details " & _
'      " WHERE (Flag IS NULL) " & wc

tmp = " SELECT 'WHR' AS TT_Type, a.WHRNo, 'WHR' AS TxnType,l.locationid,l.locationname,a.WHRDate,a.Flag,s.statename"
tmp = tmp & " From dbo.Txn_WHR_Detail  a"
tmp = tmp & " Inner Join Mst_location l On a.sm_prefix = l.sm_prefix"
tmp = tmp & " Inner Join Mst_state s On l.stateID = s.stateID"
tmp = tmp & " Where (a.Flag Is Null)" & wc
tmp = tmp & " Union All"
tmp = tmp & " SELECT 'CxTF' AS TT_Type, b.CxTF_DetailsID, b.TxnType,l.locationid,l.locationname,b.CxTFDate,b.Flag,s.statename "
tmp = tmp & " From dbo.Txn_CxTF_Details  b"
tmp = tmp & " Inner Join Mst_location l On b.sm_prefix = l.sm_prefix"
tmp = tmp & " Inner Join Mst_CMP c On b.CMPID = c.CMPID"
tmp = tmp & " Inner Join Mst_state s On l.stateID = s.stateID"
tmp = tmp & " Where (b.Flag Is Null) and c.CloseDate is null " & wc

Call TmpTable

Call dispdata

End Sub
Private Sub GetAM(x_ As Variant, mLocationName_ As Variant)

tmp = " Select Distinct ML.LocationName,EMP1.EmployeeName + '~' + EMP1.EmployeeNo As 'EMPloyee',"
tmp = tmp & " MD.Designation"
tmp = tmp & " From Mst_EMPCMPMapping MECM"
tmp = tmp & " Inner Join Mst_CMP MC On MECM.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Employee EMP On MECm.EmployeeID = EMP.EmployeeID"
tmp = tmp & " Inner Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID"
tmp = tmp & " Where ML.LocationName = '" & mLocationName_ & "' And EMP.Flag = 'Y' "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.TextMatrix(x_, 8) = IIf(IsNull(TmpRs1!Employee), "", TmpRs1!Employee)
End If


End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_PrintPending.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_PrintPending.xls")
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

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If

Call CmdClear_Click
'Call Grid_Head
'CmdExcel.Enabled = False
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9
MSHFlexGrid1.AllowUserResizing = flexResizeBoth
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "Transaction_Type"
MSHFlexGrid1.TextMatrix(0, 1) = "WHR No/CxTF No"
MSHFlexGrid1.TextMatrix(0, 2) = "TxnType"
MSHFlexGrid1.TextMatrix(0, 3) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 4) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 5) = "WHR Date/CxTF Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Flag"
MSHFlexGrid1.TextMatrix(0, 7) = "State Name"
MSHFlexGrid1.TextMatrix(0, 8) = "AM"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 0
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1600

MSHFlexGrid1.RowHeight(0) = 800

End Sub

'''Public Sub DispData()
'''Dim x As Double
'''Dim mRWCID_ As Variant
''''x = 0
'''
'''If TmpRs.RecordCount > 0 Then
'''   For i = 1 To TmpRs.RecordCount
'''       'x = MSHFlexGrid1.Rows
'''       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'''       For c = 0 To TmpRs.Fields.Count - 1
'''           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'''       Next
'''       If MSHFlexGrid1.TextMatrix(i, 0) = "WHR" Then
'''          mRWCID_ = Format((MSHFlexGrid1.TextMatrix(i, 1) / 1000000), "#####")
'''          mRWCID_ = mRWCID_ & "-" & (MSHFlexGrid1.TextMatrix(i, 1) - (mRWCID_ * 1000000))
'''          MSHFlexGrid1.TextMatrix(i, 1) = mRWCID_
'''       End If
'''       TmpRs.MoveNext
'''   Next
'''
'''   CmdExcel.Enabled = True
'''Else
'''   Call Grid_Head
'''End If
'''
'''Label2.Refresh
'''Label2.Caption = TmpRs.RecordCount
'''
'''End Sub

Public Sub dispdata()
Dim x As Double
Dim mRWCID_ As Variant

'Call TableMst_Location
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1

       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs.Fields(0)), "", TmpRs.Fields(0))
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs.Fields(1)), "", TmpRs.Fields(1))
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs.Fields(2)), "", TmpRs.Fields(2))
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs.Fields(3)), "", TmpRs.Fields(3))
       'Location Name
'       RsMst_Location.MoveFirst
'       RsMst_Location.Find "SM_PreFix='" & TmpRs.Fields(3) & "'"
'       If Not RsMst_Location.EOF Then
'          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(RsMst_Location!LocationName), "", RsMst_Location!LocationName)
'       End If
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs.Fields(4)), "", TmpRs.Fields(4))
       MSHFlexGrid1.TextMatrix(I, 5) = Format(IIf(IsNull(TmpRs.Fields(5)), "", TmpRs.Fields(5)), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs.Fields(6)), "", TmpRs.Fields(6)) '
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs.Fields(7)), "", TmpRs.Fields(7))
       Call GetAM(I, MSHFlexGrid1.TextMatrix(I, 4))
              
       If MSHFlexGrid1.TextMatrix(I, 0) = "WHR" Then
          mRWCID_ = Format((MSHFlexGrid1.TextMatrix(I, 1) / 1000000), "#####")
          mRWCID_ = mRWCID_ & "-" & (MSHFlexGrid1.TextMatrix(I, 1) - (mRWCID_ * 1000000))
          MSHFlexGrid1.TextMatrix(I, 1) = mRWCID_
       End If
       
       TmpRs.MoveNext
   Next

   CmdExcel.Enabled = True
   
Else
   Call Grid_Head
End If

Label2.Refresh
Label2.Caption = TmpRs.RecordCount
MsgBox "Done!!!"

End Sub


Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location("")
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!Locationname
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
   TxtPreFix = ""
   Exit Sub
End If
' Code for clearing Godown combox in case of change
If CmbLocation.Text <> tmp Then
   TxtPreFix = ""
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
End Sub
