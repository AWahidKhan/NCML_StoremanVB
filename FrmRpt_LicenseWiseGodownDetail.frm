VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_LicenseWiseGodownDetail 
   Caption         =   "License Wise Godown Detail"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
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
      Height          =   2565
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   14685
      Begin VB.TextBox TxtExpiryOn 
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
         Height          =   345
         Left            =   2130
         MaxLength       =   200
         TabIndex        =   29
         Top             =   2055
         Width           =   1785
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
         ItemData        =   "FrmRpt_LicenseWiseGodownDetail.frx":0000
         Left            =   11280
         List            =   "FrmRpt_LicenseWiseGodownDetail.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1650
         Width           =   3285
      End
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
         ItemData        =   "FrmRpt_LicenseWiseGodownDetail.frx":0004
         Left            =   90
         List            =   "FrmRpt_LicenseWiseGodownDetail.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1650
         Width           =   4095
      End
      Begin VB.ComboBox CmbLicenseNo 
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
         ItemData        =   "FrmRpt_LicenseWiseGodownDetail.frx":0008
         Left            =   6735
         List            =   "FrmRpt_LicenseWiseGodownDetail.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   3720
      End
      Begin VB.ComboBox CmbLicenseType 
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
         ItemData        =   "FrmRpt_LicenseWiseGodownDetail.frx":000C
         Left            =   3000
         List            =   "FrmRpt_LicenseWiseGodownDetail.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   3735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Godown Not Mapped to any License"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Godown Wise License Details"
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
         TabIndex        =   14
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "License Wise Godown Details"
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
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
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
         Left            =   13410
         TabIndex        =   12
         Top             =   2145
         Width           =   1125
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
         ItemData        =   "FrmRpt_LicenseWiseGodownDetail.frx":0010
         Left            =   4230
         List            =   "FrmRpt_LicenseWiseGodownDetail.frx":0012
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1650
         Width           =   2910
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
         Left            =   12285
         TabIndex        =   11
         Top             =   2145
         Width           =   1125
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
         Left            =   11160
         TabIndex        =   10
         Top             =   2145
         Width           =   1125
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
         ItemData        =   "FrmRpt_LicenseWiseGodownDetail.frx":0014
         Left            =   7155
         List            =   "FrmRpt_LicenseWiseGodownDetail.frx":0016
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1650
         Width           =   4125
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
         Left            =   10515
         MaxLength       =   12
         TabIndex        =   5
         Top             =   960
         Width           =   2490
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   1545
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
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
         Format          =   47775745
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
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
         Format          =   47775745
         CurrentDate     =   36494
      End
      Begin VB.Label lblExpiryOn 
         AutoSize        =   -1  'True
         Caption         =   "Expire Within Days"
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
         Left            =   120
         TabIndex        =   30
         Top             =   2107
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "Godown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12442
         TabIndex        =   27
         Top             =   1380
         Width           =   960
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
         Left            =   1837
         TabIndex        =   26
         Top             =   1380
         Width           =   600
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "License No"
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
         Left            =   8085
         TabIndex        =   25
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label Label13 
         Caption         =   "License Type"
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
         Left            =   4215
         TabIndex        =   24
         Top             =   660
         Width           =   1305
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
         Left            =   9960
         TabIndex        =   23
         Top             =   2160
         Width           =   90
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
         Left            =   5265
         TabIndex        =   22
         Top             =   1380
         Width           =   840
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
         Left            =   7440
         TabIndex        =   21
         Top             =   2160
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
         Left            =   330
         TabIndex        =   20
         Top             =   660
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
         Left            =   1950
         TabIndex        =   19
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "CMP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8917
         TabIndex        =   18
         Top             =   1380
         Width           =   600
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "License Mapping ID"
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
         Left            =   10875
         TabIndex        =   17
         Top             =   645
         Width           =   1785
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6615
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11668
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   9480
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmRpt_LicenseWiseGodownDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLicenseTypeID, mLicenseNo, mStateID, mLocationID, mCMPID, mGodownID As Variant

Private Sub CmbGodown_GotFocus()

If CmbCMP.Text = "" Then
    MsgBox "Select CMP First!!!"
    CmbCMP.SetFocus
    Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown(" where CMPID = " & mCMPID)

CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then
   Exit Sub
End If

If RsMst_Godown.RecordCount > 0 Then RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
mGodownNO = RsMst_Godown!GodownNo

End Sub

Private Sub CmbLicenseNo_GotFocus()

tmp = CmbLicenseNo.Text
If CmbLicenseType.Text = "" Then
   Call TableMst_License("")
Else
   Call TableMst_License("where LicenseTypeID=" & mLicenseTypeID & "")
End If
CmbLicenseNo.Clear
If RsMst_License.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_License.RecordCount
    CmbLicenseNo.AddItem RsMst_License!LicenseNo
    RsMst_License.MoveNext
Next
CmbLicenseNo.Text = tmp

End Sub


Private Sub CmbLicenseNo_LostFocus()

If CmbLicenseNo.Text = "" Then
   Exit Sub
End If

If RsMst_License.RecordCount > 0 Then RsMst_License.MoveFirst
RsMst_License.Find "LicenseNo = '" & CmbLicenseNo.Text & "'"

If RsMst_License.EOF Then
   MsgBox "Invalid Selection "
   CmbLicenseNo.SetFocus
   Exit Sub
End If

mLicenseNo = RsMst_License!LicenseNo
End Sub


Private Sub CmbLicenseType_GotFocus()
tmp = CmbLicenseType.Text
Call TableMst_LicenseType("")
CmbLicenseType.Clear
If RsMst_LicenseType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_LicenseType.RecordCount
    CmbLicenseType.AddItem RsMst_LicenseType!LicenseTypeName
    RsMst_LicenseType.MoveNext
Next
CmbLicenseType.Text = tmp
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
For I = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub


Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   CmbSBaseLoc.Text = ""
   CmbCMP.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If
If RsMst_State.RecordCount > 0 Then RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If

If CmbState.Text <> tmp Then
   CmbSBaseLoc.Text = ""
   CmbCMP.Text = ""
   CmbGodown.Text = ""
End If

mStateID = RsMst_State!StateID

End Sub


Private Sub CmdClear_Click()
Call Grid_Head1
Frame1.Enabled = True

ProgressBar1.Value = 0
Option1.Value = True
Option2.Value = False
Option3.Value = False

SFrom = Date - 30
STo = Date

CmbLicenseType.Text = ""
CmbCMP.Text = ""
CmbSBaseLoc.Text = ""
TxtSID.Text = ""
LblTotalRecords.Caption = ""
CmbState.Text = ""
CmbSBaseLoc.Text = ""
CmbCMP.Text = ""
CmbGodown.Text = ""
CmbLicenseNo.Text = ""

CmbLicenseType.Enabled = True
CmbLicenseNo.Enabled = True
TxtSID.Enabled = True
CmdExcel.Enabled = False

End Sub
Public Sub Grid_Head2()
With MSHFlexGrid1
     .Clear
     .Rows = 3
     .FixedRows = 2
     .FixedCols = 0
     .WordWrap = True
     .Cols = 11
     
     .TextMatrix(1, 0) = "Godown No"
     .TextMatrix(1, 1) = "Godown ID"
     .TextMatrix(1, 2) = "CMP Name"
     .TextMatrix(1, 3) = "Location Name"
     .TextMatrix(1, 4) = "State Name"
     .TextMatrix(1, 5) = "Godown Capacity"
     .TextMatrix(1, 6) = "Start Date"
     .TextMatrix(1, 7) = "Agreement End Date"
     .TextMatrix(1, 8) = "Unit Type"
     .TextMatrix(1, 9) = "License Type Name"
     .TextMatrix(1, 10) = "LicenseNo"
     
     .ColWidth(0) = 1500
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     .ColWidth(3) = 1500
     .ColWidth(4) = 1500
     .ColWidth(5) = 1500
     .ColWidth(6) = 1500
     .ColWidth(7) = 1500
     .ColWidth(8) = 1500
     .ColWidth(9) = 1500
     .ColWidth(10) = 1500
End With
End Sub
Public Sub Grid_Head3()
With MSHFlexGrid1
     .Clear
     
     .Rows = 3
     .FixedRows = 2
     .FixedCols = 0
     .WordWrap = True
     .Cols = 9
      
     .TextMatrix(1, 0) = "Godown No"
     .TextMatrix(1, 1) = "Godown ID"
     .TextMatrix(1, 2) = "CMP Name"
     .TextMatrix(1, 3) = "Location Name"
     .TextMatrix(1, 4) = "State Name"
     .TextMatrix(1, 5) = "Godown Capacity"
     .TextMatrix(1, 6) = "Start Date"
     .TextMatrix(1, 7) = "Agreement End Date"
     .TextMatrix(1, 8) = "Unit Type"
     
     .ColWidth(0) = 1500
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     .ColWidth(3) = 1500
     .ColWidth(4) = 1500
     .ColWidth(5) = 1500
     .ColWidth(6) = 1500
     .ColWidth(7) = 1500
     .ColWidth(8) = 1500

End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "License Wise Godown Detail.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "License Wise Godown Detail.xls")
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
ProgressBar1.Value = 0
wc = ""
Wc1 = ""
   
If Option1.Value = True Then
   Call Grid_Head1
      
   MsgBox "Wait till next message!!!"
      
   If SFrom.Value > STo.Value Then
      MsgBox "'From Date' must be less than 'To Date'!!! "
      SFrom.SetFocus
      Exit Sub
   End If

   tmp = " Select ML.LicenseID,MLT.LicenseTypeName,ML.LicenseNo,ML.LicenseDate,ML.LicenseExpiryDate,ML.Capacity,ML.LicenseDesc,"
   tmp = tmp & " ML.IssuingAuthority,ME.EmployeeName + '~'+ME.EmployeeNo As EmployeeName,"
   tmp = tmp & " MC.CMPName , MG.GodownNo, MLD.LocationName, ms.StateName, MG.GodownCapacity, MG.StartDate, MG.AgreementEndDate, MU.UnitType"
   tmp = tmp & " From Mst_License ML"
   tmp = tmp & " Inner Join Mst_Employee ME on ML.EmployeeID=ME.EmployeeID"
   tmp = tmp & " Inner join Mst_LicenseType MLT on ML.LicenseTypeID=MLT.LicenseTypeID"
   tmp = tmp & " Inner Join Txn_LicenseGodownDetail TQMD on ML.LicenseID=TQMD.LicenseID"
   tmp = tmp & " Inner Join Mst_Godown MG on TQMD.GodownID=MG.GodownID"
   tmp = tmp & " Inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
   tmp = tmp & " Inner join Mst_Location MLD on MC.LocationID=MLD.LocationID"
   tmp = tmp & " Inner join Mst_State MS on MLD.StateID=MS.StateID"
   tmp = tmp & " Inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
   
   If Gen_DBType = "MDB" Then
      wc = " Where ML.LicenseDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
   Else
      wc = " Where ML.LicenseDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
   End If
   
   If CmbLicenseType.Text <> "" Then
      Wc1 = Wc1 & " And MLT.LicenseTypeName='" & CmbLicenseType.Text & "'"
   End If
   
   If CmbLicenseNo.Text <> "" Then
      Wc1 = Wc1 & " And ML.LicenseNo='" & CmbLicenseNo.Text & "'"
   End If
    
   If CmbState.Text <> "" Then
      Wc1 = Wc1 & " And MS.StateName='" & CmbState.Text & "'"
   End If
   
   If CmbSBaseLoc.Text <> "" Then
      Wc1 = Wc1 & " And MLD.LocationName='" & CmbSBaseLoc.Text & "'"
   End If
   
   If CmbCMP.Text <> "" Then
      Wc1 = Wc1 & " And MC.CMPNAME='" & CmbCMP.Text & "'"
   End If
   
   If CmbGodown.Text <> "" Then
      Wc1 = Wc1 & " And MG.GodownNo='" & CmbGodown.Text & "'"
   End If
   
   If TxtSID.Text <> "" Then
      Wc1 = Wc1 & " And ML.LicenseID='" & Val(TxtSID.Text) & "'"
   End If
   
   If TxtExpiryOn <> "" Then
      Wc1 = Wc1 & " And ML.LicenseExpiryDate >= '" & Date & "'  And ML.LicenseExpiryDate <= dateadd(day," & Val(TxtExpiryOn) & ",'" & Date & "')" 'ML.LicenseExpiryDate Val (TxtExpiryOn)"
   End If
   
   tmp = tmp & " " & wc & " " & Wc1
    
   tmp = tmp & " order By ML.LicenseID"

ElseIf Option2.Value = True Then
   Call Grid_Head2
   
   MsgBox "Wait till next message!!!"
   
   tmp = " Select MG.GodownNo,MG.GodownId,MC.CMPName,MLD.LocationName,ms.StateName,MG.GodownCapacity,MG.StartDate,MG.AgreementEndDate , MU.UnitType, MLT.LicenseTypeName, ML.LicenseNo"
   tmp = tmp & " From Mst_License ML"
   tmp = tmp & " Inner Join Mst_Employee ME on ML.EmployeeID=ME.EmployeeID"
   tmp = tmp & " Inner join Mst_LicenseType MLT on ML.LicenseTypeID=MLT.LicenseTypeID"
   tmp = tmp & " Inner Join Txn_LicenseGodownDetail TQMD on ML.LicenseID=TQMD.LicenseID"
   tmp = tmp & " Inner Join Mst_Godown MG on TQMD.GodownID=MG.GodownID"
   tmp = tmp & " Inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
   tmp = tmp & " Inner join Mst_Location MLD on MC.LocationID=MLD.LocationID"
   tmp = tmp & " Inner join Mst_State MS on MLD.StateID=MS.StateID"
   tmp = tmp & " Inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"

   If Gen_DBType = "MDB" Then
      wc = " Where ML.LicenseDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
   Else
      wc = " Where ML.LicenseDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
   End If
   
   If CmbLicenseType.Text <> "" Then
      Wc1 = Wc1 & " And MLT.LicenseTypeName='" & CmbLicenseType.Text & "'"
   End If
   
   If CmbLicenseNo.Text <> "" Then
      Wc1 = Wc1 & " And ML.LicenseNo='" & CmbLicenseNo.Text & "'"
   End If
    
   If CmbState.Text <> "" Then
      Wc1 = Wc1 & " And MS.StateName='" & CmbState.Text & "'"
   End If
   
   If CmbSBaseLoc.Text <> "" Then
      Wc1 = Wc1 & " And MLD.LocationName='" & CmbSBaseLoc.Text & "'"
   End If
   
   If CmbCMP.Text <> "" Then
      Wc1 = Wc1 & " And MC.CMPNAME='" & CmbCMP.Text & "'"
   End If
   
   If CmbGodown.Text <> "" Then
      Wc1 = Wc1 & " And MG.GodownNo='" & CmbGodown.Text & "'"
   End If
   
   If TxtSID.Text <> "" Then
      Wc1 = Wc1 & " And ML.LicenseID='" & Val(TxtSID.Text) & "'"
   End If
   
   If TxtExpiryOn <> "" Then
      Wc1 = Wc1 & " And ML.LicenseExpiryDate >= '" & Date & "'  And ML.LicenseExpiryDate <= dateadd(day," & Val(TxtExpiryOn) & ",'" & Date & "')" 'ML.LicenseExpiryDate Val (TxtExpiryOn)"
   End If
   
   tmp = tmp & " " & wc & " " & Wc1
    
   tmp = tmp & " order By ML.LicenseID"
   
ElseIf Option3.Value = True Then
   Call Grid_Head3
   MsgBox "Wait till next message!!!"
   Wc1 = ""
   
   tmp = " Select MG.GodownNo,MG.GodownId,MC.CMPName,MLD.LocationName,ms.StateName,MG.GodownCapacity,MG.StartDate,"
   tmp = tmp & " MG.AgreementEndDate , MU.UnitType"
   tmp = tmp & " From Mst_Godown MG"
   tmp = tmp & " Inner join Mst_Cmp MC on MG.CmpID=MC.CmpID"
   tmp = tmp & " Inner join Mst_Location MLD on MC.LocationID=MLD.LocationID"
   tmp = tmp & " Inner join Mst_State MS on MLD.StateID=MS.StateID"
   tmp = tmp & " Inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID"
   tmp = tmp & " Where MG.GodownId Not In (Select GodownId From Txn_LicenseGodownDetail)"

   If CmbState.Text <> "" Then
      Wc1 = Wc1 & " And MS.StateName='" & CmbState.Text & "'"
   End If
   
   If CmbSBaseLoc.Text <> "" Then
      Wc1 = Wc1 & " And MLD.LocationName='" & CmbSBaseLoc.Text & "'"
   End If
   
   If CmbCMP.Text <> "" Then
      Wc1 = Wc1 & " And MC.CMPNAME='" & CmbCMP.Text & "'"
   End If
   
   If CmbGodown.Text <> "" Then
      Wc1 = Wc1 & " And MG.GodownNo='" & CmbGodown.Text & "'"
   End If
   tmp = tmp & " " & wc & " " & Wc1
   tmp = tmp & " order By MG.GodownNo"
End If
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs1.RecordCount + 2
   CmdExcel.Enabled = True
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 3
   For I = 2 To TmpRs1.RecordCount + 1
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       If Option1.Value = True Then
          MSHFlexGrid1.TextMatrix(I, 3) = IIf(MSHFlexGrid1.TextMatrix(I, 3) <> "", Format(MSHFlexGrid1.TextMatrix(I, 3), "dd-MMM-yyyy"), "")
          MSHFlexGrid1.TextMatrix(I, 4) = IIf(MSHFlexGrid1.TextMatrix(I, 4) <> "", Format(MSHFlexGrid1.TextMatrix(I, 4), "dd-MMM-yyyy"), "")
          MSHFlexGrid1.TextMatrix(I, 14) = IIf(MSHFlexGrid1.TextMatrix(I, 14) <> "", Format(MSHFlexGrid1.TextMatrix(I, 14), "dd-MMM-yyyy"), "")
          MSHFlexGrid1.TextMatrix(I, 15) = IIf(MSHFlexGrid1.TextMatrix(I, 15) <> "", Format(MSHFlexGrid1.TextMatrix(I, 15), "dd-MMM-yyyy"), "")
       ElseIf Option2.Value = True Or Option3.Value = True Then
          MSHFlexGrid1.TextMatrix(I, 6) = IIf(MSHFlexGrid1.TextMatrix(I, 6) <> "", Format(MSHFlexGrid1.TextMatrix(I, 6), "dd-MMM-yyyy"), "")
          MSHFlexGrid1.TextMatrix(I, 7) = IIf(MSHFlexGrid1.TextMatrix(I, 7) <> "", Format(MSHFlexGrid1.TextMatrix(I, 7), "dd-MMM-yyyy"), "")
       End If
       TmpRs1.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
ProgressBar1.Value = ProgressBar1.Max
End If
LblTotalRecords.Refresh
LblTotalRecords.Caption = MSHFlexGrid1.Rows - 3
MsgBox "Done!!!"
End Sub
Public Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 17

MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.FixedCols = 0
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.MergeCells = flexMergeRestrictRows

MSHFlexGrid1.TextMatrix(0, 0) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 1) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 2) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 3) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 4) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 5) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 6) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 7) = "License Detail"
MSHFlexGrid1.TextMatrix(0, 8) = "License Detail"

MSHFlexGrid1.TextMatrix(0, 9) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 10) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 11) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 12) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 13) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 14) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 15) = "Godown Detail"
MSHFlexGrid1.TextMatrix(0, 16) = "Godown Detail"

MSHFlexGrid1.TextMatrix(1, 0) = "License Map ID"
MSHFlexGrid1.TextMatrix(1, 1) = "License Type"
MSHFlexGrid1.TextMatrix(1, 2) = "License No"
MSHFlexGrid1.TextMatrix(1, 3) = "License Date"
MSHFlexGrid1.TextMatrix(1, 4) = "License Expiry Date"
MSHFlexGrid1.TextMatrix(1, 5) = "Capacity for which License Issued"
MSHFlexGrid1.TextMatrix(1, 6) = "License Description"
MSHFlexGrid1.TextMatrix(1, 7) = "License Issuing Authority"
MSHFlexGrid1.TextMatrix(1, 8) = "Name of the Person holding license copy"

MSHFlexGrid1.TextMatrix(1, 9) = "CMP"
MSHFlexGrid1.TextMatrix(1, 10) = "Godown No"
MSHFlexGrid1.TextMatrix(1, 11) = "Location"
MSHFlexGrid1.TextMatrix(1, 12) = "State"
MSHFlexGrid1.TextMatrix(1, 13) = "Capacity"
MSHFlexGrid1.TextMatrix(1, 14) = "Start Date"
MSHFlexGrid1.TextMatrix(1, 15) = "Agreement End Date"
MSHFlexGrid1.TextMatrix(1, 16) = "Unit Type"

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
MSHFlexGrid1.ColWidth(14) = 1200
MSHFlexGrid1.ColWidth(15) = 1200
MSHFlexGrid1.ColWidth(16) = 1200

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
MSHFlexGrid1.MergeCol(16) = True

MSHFlexGrid1.RowHeight(1) = 800

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
CmbLicenseType.Enabled = True
CmbLicenseNo.Enabled = True
TxtSID.Enabled = True
End Sub

Private Sub Option2_Click()
CmbLicenseType.Enabled = True
CmbLicenseNo.Enabled = True
TxtSID.Enabled = True
End Sub

Private Sub Option3_Click()
CmbLicenseType.Enabled = False
CmbLicenseNo.Enabled = False
TxtSID.Enabled = False
End Sub


Private Sub CmbLicenseType_LostFocus()

If CmbLicenseType.Text = "" Then
   CmbLicenseNo.Text = ""
   mLicenseTypeID = Null
   Exit Sub
End If

If RsMst_LicenseType.RecordCount > 0 Then RsMst_LicenseType.MoveFirst
RsMst_LicenseType.Find "LicenseTypeName = '" & CmbLicenseType.Text & "'"

If RsMst_LicenseType.EOF Then
   MsgBox "Invalid Selection "
   CmbLicenseType.SetFocus
   Exit Sub
End If

mLicenseTypeID = RsMst_LicenseType!LicenseTypeID
End Sub





Private Sub CmbCMP_GotFocus()

If CmbSBaseLoc.Text = "" Then
   MsgBox "Please Select Location"
   CmbSBaseLoc.SetFocus
   Exit Sub
End If

tmp = CmbCMP.Text

Call TableMst_CMP(" Where LocationId=" & mLocationID & "")

CmbCMP.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No CMP found"
   CmbCMP.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMP.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMP.Text = tmp
End Sub

Private Sub CmbCMP_LostFocus()

If CmbCMP.Text = "" Then
   CmbGodown.Text = ""
   Exit Sub
End If

If RsMst_CMP.RecordCount > 0 Then RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPNAME = '" & CmbCMP.Text & "'"

If RsMst_CMP.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbCMP.SetFocus
   Exit Sub
End If

If CmbCMP.Text <> tmp Then
   CmbGodown.Text = ""
End If

mCMPID = RsMst_CMP!CMPID

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
   CmbCMP.Text = ""
   CmbGodown.Text = ""
   Exit Sub
End If
If RsMst_Location.RecordCount > 0 Then RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSBaseLoc.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbSBaseLoc.SetFocus
   Exit Sub
End If

If CmbSBaseLoc.Text <> tmp Then
   CmbCMP.Text = ""
   CmbGodown.Text = ""
End If


mLocationID = RsMst_Location.Fields("LocationID")

End Sub


Private Sub TxtExpiryOn_LostFocus()

If TxtExpiryOn = "" Then Exit Sub
If IsNumeric(TxtExpiryOn) = False Or Val(TxtExpiryOn.Text) < 0 Then
   MsgBox "Invalid numeric format !!!!"
   TxtExpiryOn.SetFocus
   Exit Sub
End If
End Sub
