VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_QQSGodownMapping 
   Caption         =   "QQS Godown Mapping"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14685
      Begin VB.ComboBox CmbState 
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
         ItemData        =   "FrmRpt_QQSGodownMapping.frx":0000
         Left            =   165
         List            =   "FrmRpt_QQSGodownMapping.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   2040
         Width           =   3735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Godown Not Mapped to QQS "
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
         Left            =   6120
         TabIndex        =   22
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         Caption         =   "QQS Not Mapped Godown"
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
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "QQS Godown Mapping Detail"
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
         Left            =   135
         TabIndex        =   20
         Top             =   240
         Width           =   3255
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
         Height          =   350
         Left            =   11805
         TabIndex        =   19
         Top             =   2265
         Width           =   1000
      End
      Begin VB.ComboBox CmbSBaseLoc 
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
         ItemData        =   "FrmRpt_QQSGodownMapping.frx":0004
         Left            =   3960
         List            =   "FrmRpt_QQSGodownMapping.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   2910
      End
      Begin VB.ComboBox CmbSQQS 
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
         ItemData        =   "FrmRpt_QQSGodownMapping.frx":0008
         Left            =   3045
         List            =   "FrmRpt_QQSGodownMapping.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1155
         Width           =   3735
      End
      Begin VB.ComboBox CmbSCMPM 
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
         ItemData        =   "FrmRpt_QQSGodownMapping.frx":000C
         Left            =   6780
         List            =   "FrmRpt_QQSGodownMapping.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1155
         Width           =   3720
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
         Height          =   350
         Left            =   12810
         TabIndex        =   4
         Top             =   2265
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
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
         Height          =   350
         Left            =   10800
         TabIndex        =   3
         Top             =   2265
         Width           =   1000
      End
      Begin VB.ComboBox CmbSDesignation 
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
         ItemData        =   "FrmRpt_QQSGodownMapping.frx":0010
         Left            =   10545
         List            =   "FrmRpt_QQSGodownMapping.frx":0012
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1155
         Width           =   3525
      End
      Begin VB.TextBox TxtSID 
         Alignment       =   1  'Right Justify
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
         Left            =   6900
         MaxLength       =   12
         TabIndex        =   1
         Top             =   2040
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   1590
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   49217537
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   135
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   49217537
         CurrentDate     =   36494
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
         Left            =   12120
         TabIndex        =   25
         Top             =   1920
         Width           =   90
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
         Left            =   1725
         TabIndex        =   24
         Top             =   1710
         Width           =   600
      End
      Begin VB.Label Label14 
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
         Left            =   4815
         TabIndex        =   17
         Top             =   1710
         Width           =   1440
      End
      Begin VB.Label Label13 
         Caption         =   "QQS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4620
         TabIndex        =   16
         Top             =   855
         Width           =   585
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "CMPM"
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
         Left            =   8340
         TabIndex        =   15
         Top             =   855
         Width           =   600
      End
      Begin VB.Label Label28 
         Caption         =   "Total Records Found :"
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
         Left            =   9600
         TabIndex        =   14
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
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
         Left            =   375
         TabIndex        =   13
         Top             =   855
         Width           =   975
      End
      Begin VB.Label Label45 
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
         Left            =   1995
         TabIndex        =   12
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "StoreMan Role"
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
         Left            =   11580
         TabIndex        =   11
         Top             =   855
         Width           =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "QQS Godown Mapping ID"
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
         Left            =   6990
         TabIndex        =   10
         Top             =   1800
         Width           =   2310
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6735
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11880
      _Version        =   393216
      WordWrap        =   -1  'True
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
Attribute VB_Name = "FrmRpt_QQSGodownMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mEmployeeID, mSEmployeeID, mSDirectReportID, mSDesignationID, mSLocationID As Variant
Dim mStateID As Variant

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
For I = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub


Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub


Private Sub CmdClear_Click()
Call Grid_Head1

Frame1.Enabled = True

Option1.Value = True
Option2.Value = False
Option3.Value = False

SFrom = Date - 30
STo = Date

CmbSQQS.Text = ""
CmbSCMPM.Text = ""
CmbSDesignation.Text = ""
CmbSBaseLoc.Text = ""
TxtSID.Text = ""
LblTotalRecords.Caption = ""

SFrom.Enabled = True
STo.Enabled = True
CmbSQQS.Enabled = True
CmbSCMPM.Enabled = True
CmbSDesignation.Enabled = True
CmbSBaseLoc.Enabled = True
TxtSID.Enabled = True

End Sub
Public Sub Grid_Head2()
With MSHFlexGrid1
     .Clear
     
     .Rows = 3
     .FixedRows = 2
     .FixedCols = 0
     .WordWrap = True
     .Cols = 4
     
     .TextMatrix(1, 0) = "QQS Name"
     .TextMatrix(1, 1) = "Status"
     .TextMatrix(1, 2) = "CMPM"
     .TextMatrix(1, 3) = "Base Location Name"
     
     .ColWidth(0) = 3000
     .ColWidth(1) = 1500
     .ColWidth(2) = 3000
     .ColWidth(3) = 3000
End With
End Sub
Public Sub Grid_Head3()
With MSHFlexGrid1
     .Clear
     
     .Rows = 3
     .FixedRows = 2
     .FixedCols = 0
     .WordWrap = True
     .Cols = 6
      
     .TextMatrix(1, 0) = "State Name"
     .TextMatrix(1, 1) = "Location Name"
     .TextMatrix(1, 2) = "CMP Name"
     .TextMatrix(1, 3) = "Godown No"
     .TextMatrix(1, 4) = "Godown Capacity"
     .TextMatrix(1, 5) = "Start Date"
     
     .ColWidth(0) = 2000
     .ColWidth(1) = 2000
     .ColWidth(2) = 3000
     .ColWidth(3) = 2000
     .ColWidth(4) = 2000
     .ColWidth(5) = 2000
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "QQS Godown Mapping Report.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "QQS Godown Mapping Report.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

For I = 1 To MSHFlexGrid1.Rows - 1
    mRow = I '+ 1 ' + 6
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()

CmdExcel.Enabled = False

If Option1.Value = True Then
   Call Grid_Head1

   If SFrom.Value > STo.Value Then
      MsgBox "'From Date' must be less than 'To Date'!!! "
      SFrom.SetFocus
      Exit Sub
   End If

   tmp = " Select TQM.QQSMapID,TQM.EntryDate,ME.EmployeeName+'~'+ME.EmployeeNo as FullName,ME.Flag,MEE.EmployeeName+'~'+MEE.EmployeeNo as FullName1,ML.LocationName,MD.Designation,TQM.Remarks,"
   tmp = tmp & " MC.CMPName,MG.GodownNo,MLD.LocationName,MS.StateName,MG.GodownCapacity,MG.StartDate,MG.AgreementEndDate,MU.UnitType"
   tmp = tmp & " From Txn_QQSGodownMapping TQM"
   tmp = tmp & " Inner Join Mst_Employee ME on TQM.EmployeeID=ME.EmployeeID"
   tmp = tmp & " Inner Join Mst_Employee MEE on ME.DirectReportingTo=MEE.EmployeeID"
   tmp = tmp & " Inner join Mst_Location ML on MEE.BaseLocationID=ML.LocationID"
   tmp = tmp & " Inner join Mst_Designation MD on MEE.DesignationID=MD.DesignationID"
   tmp = tmp & " Inner Join Txn_QQSGodownMappingDetail TQMD on TQM.QQSMapID=TQMD.QQSMapID"
   tmp = tmp & " Inner Join Mst_Godown MG on TQMD.GodownID=MG.GodownID"
   tmp = tmp & " Inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
   tmp = tmp & " Inner join Mst_Location MLD on MC.LocationID=MLD.LocationID"
   tmp = tmp & " Inner join Mst_State MS on ML.StateID=MS.StateID"
   tmp = tmp & " Inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"

   If Gen_DBType = "MDB" Then
      Wc1 = " Where TQM.EntryDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "#"
   Else
      Wc1 = " Where TQM.EntryDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "'"
   End If
   
   If CmbSQQS.Text <> "" Then
      Wc1 = Wc1 & " And ME.EmployeeID=" & mSEmployeeID & ""
   End If
   
   If CmbSCMPM.Text <> "" Then
      Wc1 = Wc1 & " And MEE.EmployeeID=" & mSDirectReportID & ""
   End If
    
   If CmbSDesignation.Text <> "" Then
      Wc1 = Wc1 & " And MD.Designation='" & CmbSDesignation.Text & "'"
   End If
   
   If TxtSID.Text <> "" Then
      Wc1 = Wc1 & " And TQM.QQSMapID='" & Val(TxtSID.Text) & "'"
   End If
   
   If CmbSBaseLoc.Text <> "" Then
      Wc1 = Wc1 & " And ML.LocationName='" & CmbSBaseLoc.Text & "'"
   End If
   
   If CmbState.Text <> "" Then
      Wc1 = Wc1 & " And MS.StateID='" & mStateID & "'"
   End If
   
   tmp = tmp & Wc1
    
   tmp = tmp & " order By TQM.QQSMapID"

ElseIf Option2.Value = True Then
   Call Grid_Head2
   
   Wc1 = Wc1 & " And ME.Flag='Y'"
   
   If CmbSBaseLoc.Text <> "" Then
      Wc1 = Wc1 & " And ML.LocationName='" & CmbSBaseLoc.Text & "'"
   End If
   
   If CmbState.Text <> "" Then
      Wc1 = Wc1 & " And MS.StateID='" & mStateID & "'"
   End If
   
   tmp = " Select ME.EmployeeName+'~'+ME.EmployeeNO as QQSName,ME.Flag as Status,ME1.EmployeeName+'~'+ME1.EmployeeNO AS CMPM,ML.LocationName AS Location"
   tmp = tmp & " From Mst_Employee ME"
   tmp = tmp & " Inner join Mst_Employee ME1 on ME.DirectReportingTo=ME1.EmployeeID"
   tmp = tmp & " Inner join Mst_Location ML on ME.BaseLocationID=ML.LocationID"
   tmp = tmp & " Inner join Mst_State MS on ML.StateID=MS.StateID"
   tmp = tmp & " Where ME.DesignationID=8 And ME.EmployeeID Not In (Select EmployeeID From Txn_QQSGodownMapping)"
   tmp = tmp & Wc1
   tmp = tmp & " Union All "
   tmp = tmp & " Select ME.EmployeeName+'~'+ME.EmployeeNO as QQSName,ME.Flag as Status,ME1.EmployeeName+'~'+ME1.EmployeeNO AS CMPM,ML.LocationName AS Location"
   tmp = tmp & " From Mst_Employee ME"
   tmp = tmp & " Inner join Mst_Employee ME1 on ME.DirectReportingTo=ME1.EmployeeID"
   tmp = tmp & " Inner join Mst_Location ML on ME.BaseLocationID=ML.LocationID"
   tmp = tmp & " Inner join Mst_State MS on ML.StateID=MS.StateID"
   tmp = tmp & " Where ME.DesignationID=8 And ME.EmployeeID In (Select EmployeeID From Txn_QQSGodownMapping Where FreeFlag = 'Y' )"
   tmp = tmp & Wc1
   'tmp = tmp & " order By QQSName,ME1.EmployeeName,ML.LocationName"
   
ElseIf Option3.Value = True Then
   Call Grid_Head3
   tmp = " Select MS.StateName,ML.LocationName,MC.CMPName,MG.GodownNO,MG.GodownCapacity,MG.StartDate"
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
   tmp = tmp & " Where MG.CloseDate Is Null And GodownID Not In (Select GodownID From Txn_QQSGodownMappingDetail a inner join Txn_QQSGodownMapping b on a.QQSMapID = b.QQSMapID Inner join Mst_employee e on b.EmployeeID = e.EmployeeID where e.FLag = 'Y' )"
   
   If CmbSBaseLoc.Text <> "" Then
      Wc1 = Wc1 & " And ML.LocationName='" & CmbSBaseLoc.Text & "'"
   End If
   
   If CmbState.Text <> "" Then
      Wc1 = Wc1 & " And MS.StateID='" & mStateID & "'"
   End If
   tmp = tmp & Wc1
   tmp = tmp & " order By MS.StateName,ML.LocationName,MC.CMPName,MG.GodownNO"
End If
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   CmdExcel.Enabled = True
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 3
   For I = 2 To TmpRs1.RecordCount + 1
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       If Option1.Value = True Then
          MSHFlexGrid1.TextMatrix(I, 13) = Format(MSHFlexGrid1.TextMatrix(I, 13), "dd-MMM-yyyy")
          MSHFlexGrid1.TextMatrix(I, 14) = Format(MSHFlexGrid1.TextMatrix(I, 14), "dd-MMM-yyyy")
          MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-MMM-yyyy")
          If MSHFlexGrid1.TextMatrix(I, 3) = "Y" Then
             MSHFlexGrid1.TextMatrix(I, 3) = "Active"
          Else
             MSHFlexGrid1.TextMatrix(I, 3) = "DeActive"
          End If
          MSHFlexGrid1.TextMatrix(I, 1) = Format(MSHFlexGrid1.TextMatrix(I, 1), "dd-MMM-yyyy")
       ElseIf Option2.Value = True Then
          If MSHFlexGrid1.TextMatrix(I, 1) = "Y" Then
             MSHFlexGrid1.TextMatrix(I, 1) = "Active"
          Else
             MSHFlexGrid1.TextMatrix(I, 1) = "DeActive"
          End If
       Else
          MSHFlexGrid1.TextMatrix(I, 5) = Format(MSHFlexGrid1.TextMatrix(I, 5), "dd-MMM-yyyy")
       End If
       TmpRs1.MoveNext
   Next
End If
LblTotalRecords.Refresh
LblTotalRecords.Caption = MSHFlexGrid1.Rows - 3
MsgBox "Done!!!"
End Sub
Public Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 16

MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.FixedCols = 0
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.MergeCells = flexMergeRestrictRows

MSHFlexGrid1.TextMatrix(0, 0) = "QQS Detail"
MSHFlexGrid1.TextMatrix(0, 1) = "QQS Detail"
MSHFlexGrid1.TextMatrix(0, 2) = "QQS Detail"
MSHFlexGrid1.TextMatrix(0, 3) = "QQS Detail"
MSHFlexGrid1.TextMatrix(0, 4) = "QQS Detail"
MSHFlexGrid1.TextMatrix(0, 5) = "QQS Detail"
MSHFlexGrid1.TextMatrix(0, 6) = "QQS Detail"
MSHFlexGrid1.TextMatrix(0, 7) = "QQS Detail"

MSHFlexGrid1.TextMatrix(0, 8) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 9) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 10) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 11) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 12) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 13) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 14) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 15) = "Godown Detail"

MSHFlexGrid1.TextMatrix(1, 0) = "QQS Map ID"
MSHFlexGrid1.TextMatrix(1, 1) = "Date"
MSHFlexGrid1.TextMatrix(1, 2) = "QQS"
MSHFlexGrid1.TextMatrix(1, 3) = "Status"
MSHFlexGrid1.TextMatrix(1, 4) = "CMPM"
MSHFlexGrid1.TextMatrix(1, 5) = "BaseLocation"
MSHFlexGrid1.TextMatrix(1, 6) = "Designation"
MSHFlexGrid1.TextMatrix(1, 7) = "Remarks"

MSHFlexGrid1.TextMatrix(1, 8) = "CMP"
MSHFlexGrid1.TextMatrix(1, 9) = "Godown No"
MSHFlexGrid1.TextMatrix(1, 10) = "Location"
MSHFlexGrid1.TextMatrix(1, 11) = "State"
MSHFlexGrid1.TextMatrix(1, 12) = "Capacity"
MSHFlexGrid1.TextMatrix(1, 13) = "Start Date"
MSHFlexGrid1.TextMatrix(1, 14) = "Agreement End Date"
MSHFlexGrid1.TextMatrix(1, 15) = "Unit Type"

MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 1400
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200

MSHFlexGrid1.MergeRow(0) = True
MSHFlexGrid1.MergeCol(1) = True
MSHFlexGrid1.MergeCol(2) = True
MSHFlexGrid1.MergeCol(3) = True
MSHFlexGrid1.MergeCol(4) = True
MSHFlexGrid1.MergeCol(5) = True
MSHFlexGrid1.MergeCol(6) = True
MSHFlexGrid1.MergeCol(7) = True
MSHFlexGrid1.MergeCol(8) = True
MSHFlexGrid1.MergeCol(9) = True
MSHFlexGrid1.MergeCol(10) = True
MSHFlexGrid1.MergeCol(11) = True
MSHFlexGrid1.MergeCol(12) = True
MSHFlexGrid1.MergeCol(13) = True
MSHFlexGrid1.MergeCol(14) = True
MSHFlexGrid1.MergeCol(15) = True

MSHFlexGrid1.RowHeight(1) = 800

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call CmdClear_Click
End Sub


Private Sub Option1_Click()
SFrom.Enabled = True
STo.Enabled = True
CmbSQQS.Enabled = True
CmbSCMPM.Enabled = True
CmbSDesignation.Enabled = True
CmbSBaseLoc.Enabled = True
CmbState.Enabled = True
TxtSID.Enabled = True
End Sub

Private Sub Option2_Click()
SFrom.Enabled = False
STo.Enabled = False
CmbSQQS.Enabled = False
CmbSCMPM.Enabled = False
CmbSDesignation.Enabled = False
TxtSID.Enabled = False
CmbSBaseLoc.Enabled = True
CmbState.Enabled = True
End Sub

Private Sub Option3_Click()
SFrom.Enabled = False
STo.Enabled = False
CmbSQQS.Enabled = False
CmbSCMPM.Enabled = False
CmbSDesignation.Enabled = False
TxtSID.Enabled = False
CmbSBaseLoc.Enabled = True
CmbState.Enabled = True
End Sub
Private Sub CmbSQQS_GotFocus()
tmp = CmbSQQS.Text
Call TableMst_Employee("Where DesignationID=8 And Flag='Y'")
CmbSQQS.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSQQS.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSQQS.Text = tmp
End Sub

Private Sub CmbSQQS_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSQQS.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

EmpNo_ = ""
EmpName_ = ""

strSplitString = Split(CmbSQQS.Text, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   CmbSQQS.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSQQS.SetFocus
   Exit Sub
End If

mSEmployeeID = RsMst_Employee!EmployeeID
mSDirectReportID = RsMst_Employee!DirectReportingTo

End Sub


Private Sub CmbSCMPM_GotFocus()
tmp = CmbSCMPM.Text
Call TableMst_Employee("Where EmployeeID in (Select DirectReportingTo From Mst_Employee Where DesignationID=8 And Flag='Y')")
CmbSCMPM.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSCMPM.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSCMPM.Text = tmp
End Sub

Private Sub CmbSCMPM_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSCMPM.Text = "" Then
   mEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSCMPM, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   CmbSCMPM.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSCMPM.SetFocus
   Exit Sub
End If

mSDirectReportID = RsMst_Employee!EmployeeID
End Sub


Private Sub CmbSDesignation_GotFocus()
tmp = CmbSDesignation.Text
Call TableMst_Designation("")
CmbSDesignation.Clear
If RsMst_Designation.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbSDesignation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Designation.RecordCount
    CmbSDesignation.AddItem RsMst_Designation!Designation
    RsMst_Designation.MoveNext
Next
CmbSDesignation.Text = tmp
End Sub

Private Sub CmbSDesignation_LostFocus()
If CmbSDesignation.Text = "" Then
   mSDesignationID = Null
   Exit Sub
End If

RsMst_Designation.MoveFirst
RsMst_Designation.Find "Designation = '" & CmbSDesignation.Text & "'"

If RsMst_Designation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbSDesignation.SetFocus
   Exit Sub
End If
mSDesignationID = RsMst_Designation!DesignationID
End Sub


Private Sub CmbSBaseLoc_GotFocus()

If CmbState.Text = "" Then
   MsgBox "Please Select State"
   CmbState.SetFocus
   Exit Sub
End If

tmp = CmbSBaseLoc.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
CmbSBaseLoc.Clear

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSBaseLoc.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSBaseLoc.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSBaseLoc.Text = tmp
End Sub


Private Sub CmbSBaseLoc_LostFocus()
If CmbSBaseLoc.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSBaseLoc.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbSBaseLoc.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location.Fields("LocationID")
End Sub


