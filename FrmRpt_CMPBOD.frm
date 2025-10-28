VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_CMPBOD 
   Caption         =   "CMP BOD Check"
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
      Height          =   8655
      Left            =   120
      TabIndex        =   8
      Top             =   855
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   9
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
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbStateID 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   368
         Width           =   2955
      End
      Begin VB.ComboBox CmbLocation 
         Height          =   315
         Left            =   3082
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   368
         Width           =   3120
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
         Left            =   10200
         TabIndex        =   11
         Top             =   330
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
         Height          =   375
         Left            =   9000
         TabIndex        =   5
         Top             =   330
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
         Left            =   7800
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TxtDate 
         Height          =   330
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   111935489
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label3 
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
         Left            =   1335
         TabIndex        =   13
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label34 
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
         Height          =   225
         Left            =   4245
         TabIndex        =   12
         Top             =   120
         Width           =   810
      End
      Begin VB.Label LblDate 
         AutoSize        =   -1  'True
         Caption         =   "BOD Date"
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
         Left            =   6480
         TabIndex        =   10
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label2 
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
         Left            =   14040
         TabIndex        =   7
         Top             =   360
         Width           =   90
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
         Left            =   11520
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_CMPBOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSLocationID, mStateID As Variant

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbLocation.Clear
   CmbStateID.SetFocus
   Exit Sub
End If

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID = " & mStateID & " ")
Else
   Call TableMst_Location(" Where StateID = " & mStateID & " And " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
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
RsMst_Location.Find "Locationname = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
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
   CmbLocation.Text = ""
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
Label2.Caption = 0
TxtDate.Value = Date
CmdExcel.Enabled = False
End Sub

Private Sub CmdGetReport_Click()
Label2.Caption = ""
tmp = "Select MC.CMPID,MC.CMPName,ML.LocationName,MC.TelephoneNo,MC.MobileNo,MC.CurrentDate"
tmp = tmp & " From Mst_CMP MC Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS On MS.StateID = ML.StateID"
tmp = tmp & " Where MC.CloseDate Is Null "

If Gen_DBType = "MDB" Then
   tmp = tmp & " And MC.CurrentDate <= #" & Format(TxtDate, Gen_DatFormat) & "# "
Else
   tmp = tmp & " And MC.CurrentDate <= '" & Format(TxtDate, Gen_DatFormat) & "' "
End If
If CmbStateID.Text <> "" Then
   tmp = tmp & " And MS.StateID = " & mStateID & ""
Else
   If Gen_WcState <> "" Then
      tmp = tmp & " And MS." & Gen_WcState
   End If
End If
If CmbLocation.Text <> "" Then
   tmp = tmp & " And ML.LocationID = " & mSLocationID & ""
Else
   If Gen_WcState <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If

tmp = tmp & " Order by ML.LocationName,MC.CMPName"
Call TmpTable

Call dispdata

MsgBox ("Done!!!")

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CMPBOD.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CMPBOD.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MMM/dd/yyyy")
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
Call Grid_Head
TxtDate.Value = Date
CmdExcel.Enabled = False
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 7



MSHFlexGrid1.TextMatrix(0, 0) = "Area Manager"
MSHFlexGrid1.TextMatrix(0, 1) = "CMP ID"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Telephone No"
MSHFlexGrid1.TextMatrix(0, 5) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 6) = "Current Date"


MSHFlexGrid1.ColWidth(0) = 3500
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.ColWidth(2) = 2800
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1600

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
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, C + 1) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(I, 6) = Format(MSHFlexGrid1.TextMatrix(I, 6), "dd-mmm-yyyy")
       
       tmp = "Select ML.LocationID,ML.LocationName,EMP1.EmployeeName +'~'+ EMP1.EmployeeNo as 'AM'"
       tmp = tmp & " From Mst_EMPCMPMapping ECM"
       tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID"
       tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
       tmp = tmp & " Inner Join Mst_Designation MD On EMP1.DesignationID = MD.DesignationID"
       tmp = tmp & " Right Join Mst_Location ML On ECM.LocationID = ML.LocationID"
       tmp = tmp & " Where ML.LocationName='" & MSHFlexGrid1.TextMatrix(I, 3) & "' And EMP.Flag = 'Y' "
       tmp = tmp & " Group By ML.LocationID,ML.LocationName,EMP1.EmployeeNo,EMP1.EmployeeName, MD.Designation"
       tmp = tmp & " Order By ML.LocationName"
       
       Call Tmp1Table
       If TmpRs1.RecordCount > 0 Then
          For r = 1 To TmpRs1.RecordCount
              If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
                 MSHFlexGrid1.TextMatrix(I, 0) = MSHFlexGrid1.TextMatrix(I, 0) & "," & IIf(IsNull(TmpRs1!AM), "", TmpRs1!AM)
              Else
                 MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!AM), "", TmpRs1!AM)
              End If
              TmpRs1.MoveNext
          Next
       End If
       
       TmpRs.MoveNext
   Next
   Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If

End Sub


