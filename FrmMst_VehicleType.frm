VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_VehicleType 
   Caption         =   "Vehicle Type"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   11055
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   885
      Left            =   240
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   8520
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
         Height          =   345
         Left            =   6090
         TabIndex        =   27
         Top             =   195
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchVehicleType 
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
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   165
         Width           =   3405
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
         Left            =   7095
         TabIndex        =   25
         Top             =   195
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   28
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
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
         Left            =   11805
         TabIndex        =   31
         Top             =   210
         Width           =   75
      End
      Begin VB.Label Label23 
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
         Left            =   9405
         TabIndex        =   30
         Top             =   210
         Width           =   2355
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Vehicle Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   29
         Top             =   210
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3900
      Left            =   120
      TabIndex        =   15
      Top             =   45
      Width           =   6825
      Begin VB.TextBox TxtMaximum 
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
         Left            =   4920
         MaxLength       =   75
         TabIndex        =   3
         Top             =   1140
         Width           =   1725
      End
      Begin VB.TextBox TxtMinimum 
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
         Left            =   1950
         MaxLength       =   75
         TabIndex        =   2
         Top             =   1140
         Width           =   1200
      End
      Begin VB.TextBox TxtUpdated 
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3405
         Width           =   4695
      End
      Begin VB.TextBox TxtCreated 
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2940
         Width           =   4695
      End
      Begin VB.TextBox TxtVehicleTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
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
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox TxtVehicleType 
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
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   1
         Top             =   690
         Width           =   4695
      End
      Begin VB.TextBox TxtRemark 
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
         Height          =   1230
         Left            =   1950
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1650
         Width           =   4665
      End
      Begin VB.Label Label1 
         Caption         =   "Maximum Tare Weight in MT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3270
         TabIndex        =   33
         Top             =   1110
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "Minimum Tare Weight in MT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   120
         TabIndex        =   32
         Top             =   1065
         Width           =   1740
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
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
         TabIndex        =   23
         Top             =   3465
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
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
         TabIndex        =   22
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label LblVehicleTypeID 
         Caption         =   "Vehicle Model ID"
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
         Left            =   120
         TabIndex        =   19
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label LblVehicleType 
         Caption         =   "Vehicle Model"
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
         TabIndex        =   18
         Top             =   765
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Remark"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2153
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   6825
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5385
         TabIndex        =   6
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5385
         TabIndex        =   7
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4065
         TabIndex        =   14
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4065
         TabIndex        =   13
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2745
         TabIndex        =   12
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2745
         TabIndex        =   11
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1425
         TabIndex        =   10
         Top             =   645
         Width           =   1320
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
         Height          =   400
         Left            =   1425
         TabIndex        =   5
         Top             =   240
         Width           =   1320
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   105
         TabIndex        =   9
         Top             =   645
         Width           =   1320
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   0
         Top             =   240
         Width           =   1320
      End
   End
End
Attribute VB_Name = "FrmMst_VehicleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub CmdGo_Click()
tmp = "Select VehicleTypeId,VehicleType,MinTareWt,MaxTareWt,Remarks from Mst_VehicleType "

Dim Wc As Variant
Wc = ""
If TxtSearchVehicleType <> "" Then
   Wc = " Where VehicleType Like '%" & TxtSearchVehicleType & "%' "
End If

tmp = tmp & Wc & " Order By VehicleTypeID "

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
TxtSearchVehicleType = ""
Label22.Caption = ""

Call TableMst_VehicleType
If RsMst_VehicleType.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_VehicleModel.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_VehicleModel.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_VehicleType.MoveNext
If RsMst_VehicleType.EOF Then
   RsMst_VehicleType.MoveLast
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_VehicleType.MovePrevious
If RsMst_VehicleType.BOF Then
   RsMst_VehicleType.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_VehicleType.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_VehicleType.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_VehicleType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 150
Frame0.Visible = True
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_VehicleType.MoveFirst

'Call Grid_Head
'TxtSearchVehicleType = ""
'Label22.Caption = ""

'MSHFlexGrid1.Rows = RsMst_VehicleType.RecordCount + 1
'MSHFlexGrid1.ColWidth(0) = 1000
'For i = 1 To RsMst_VehicleType.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_VehicleType.Fields(c)), "", RsMst_VehicleType.Fields(c))
'    Next
'    RsMst_VehicleType.MoveNext
'Next
End Sub
Public Sub Indata()

TxtVehicleTypeID.Locked = True
TxtVehicleType.Locked = True
TxtMinimum.Locked = True
TxtMaximum.Locked = True
TxtRemark.Locked = True

TxtVehicleTypeID = IIf(IsNull(RsMst_VehicleType!VehicleTypeID), "", RsMst_VehicleType!VehicleTypeID)
TxtVehicleType = IIf(IsNull(RsMst_VehicleType!VehicleType), "", RsMst_VehicleType!VehicleType)

TxtMinimum = IIf(IsNull(RsMst_VehicleType!MinTareWt), 0, RsMst_VehicleType!MinTareWt)
TxtMaximum = IIf(IsNull(RsMst_VehicleType!MaxTareWt), 0, RsMst_VehicleType!MaxTareWt)
TxtRemark = IIf(IsNull(RsMst_VehicleType!Remarks), "", RsMst_VehicleType!Remarks)

TxtCreated = IIf(IsNull(RsMst_VehicleType!CreatedBy), "", RsMst_VehicleType!CreatedBy) & " :- " & IIf(IsNull(RsMst_VehicleType!CreatedDate), "", RsMst_VehicleType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_VehicleType!UpdatedBy), "", RsMst_VehicleType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_VehicleType!UpdatedDate), "", RsMst_VehicleType!UpdatedDate)

If RsMst_VehicleType.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtFlag")
End If
End Sub

Private Sub TxtMaximum_LostFocus()

If TxtMaximum = "" Then Exit Sub
If TxtMaximum.Locked = True Then Exit Sub
If Not IsNumeric(TxtMaximum) Or Val(TxtMaximum) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtMaximum.SetFocus
End If

End Sub

Private Sub TxtMinimum_LostFocus()

If TxtMinimum = "" Then Exit Sub
If TxtMinimum.Locked = True Then Exit Sub
If Not IsNumeric(TxtMinimum) Or Val(TxtMinimum) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtMinimum.SetFocus
End If

End Sub

Private Sub TxtVehicleType_LostFocus()
Dim ans As Boolean
'ans = IsClean(TxtVehicleType)
'If ans = False Then
'   MsgBox "Special character not Allowed  !!!!!!!! "
'   TxtVehicleType.SetFocus
'End If
TxtVehicleType.Text = Trim(TxtVehicleType.Text)
End Sub

Private Sub TxtVehicleTypeID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtVehicleTypeID")
End If

End Sub

Private Sub TxtVehicleTypeID_LostFocus()
If TxtVehicleTypeID = "" Then Exit Sub
If IsNumeric(TxtVehicleTypeID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtVehicleTypeID.SetFocus
   Exit Sub
End If
TxtVehicleTypeID = Val(TxtVehicleTypeID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtVehicleTypeID.Locked = True
TxtVehicleType.Locked = False
TxtMinimum.Locked = False
TxtMaximum.Locked = False
TxtRemark.Locked = False

TxtVehicleTypeID = ""
TxtVehicleType = ""
TxtMinimum = ""
TxtMaximum = ""
TxtRemark = ""
TxtCreated = ""
TxtUpdated = ""

If RsMst_VehicleType.RecordCount > 0 Then
   TxtVehicleType.SetFocus
End If
End Sub
Private Sub EditCmd_Click()

Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtVehicleTypeID.Locked = True
TxtVehicleType.Locked = False
TxtMinimum.Locked = False
TxtMaximum.Locked = False
TxtRemark.Locked = False

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   'RsMst_VehicleType
   RsMst_VehicleType.Delete
   RsMst_VehicleType.Update
   If RsMst_VehicleType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_VehicleType.MoveNext
   If RsMst_VehicleType.EOF() Then
      RsMst_VehicleType.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_VehicleType.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If TxtVehicleType.Text = "" Then
   MsgBox "Blank Vehicle Model  Not Allowed !!! "
   TxtVehicleType.SetFocus
   Exit Sub
End If

If Val(TxtMinimum) = 0 Then
   MsgBox "Blank Minimum tare weight not Allowed or Minimum tare weight should be greater than zero!!!"
   TxtMinimum.SetFocus
   Exit Sub
End If


If Val(TxtMaximum) = 0 Then
   MsgBox "Blank Maximum tare weight not Allowed or Maximum tare weight should be greater than zero!!!"
   TxtMaximum.SetFocus
   Exit Sub
End If

If Val(TxtMinimum) > Val(TxtMaximum) Then
   MsgBox "Minimum tare weight should always be less than maximum tare weight !!!"
   TxtMaximum.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsMst_VehicleType.RecordCount > 0 Then
      RsMst_VehicleType.MoveFirst
      RsMst_VehicleType.Find "VehicleType= '" & TxtVehicleType.Text & "'"
      If Not RsMst_VehicleType.EOF Then
        MsgBox "Duplicate Vehicle Model  Not Allowed !!! "
         TxtVehicleType.SetFocus
         Exit Sub
      End If
   End If
      RsMst_VehicleType.AddNew
      RsMst_VehicleType!VehicleTypeID = GetVehicleTypeID
      RsMst_VehicleType!G_UID = GetGUID
      TxtVehicleTypeID.Text = RsMst_VehicleType!VehicleTypeID
Else
    RsMst_VehicleType.MoveFirst
    RsMst_VehicleType.Find "VehicleType= '" & TxtVehicleType.Text & "'"
    If Not RsMst_VehicleType.EOF Then
       If RsMst_VehicleType!VehicleTypeID <> TxtVehicleTypeID.Text Then
          MsgBox "Duplicate Vehicle Model Not Allowed !!! "
          TxtVehicleType.SetFocus
          Exit Sub
       End If
    End If
    RsMst_VehicleType.MoveFirst
    RsMst_VehicleType.Find " VehicleTypeID= " & TxtVehicleTypeID.Text
End If

RsMst_VehicleType!VehicleTypeID = TxtVehicleTypeID.Text
RsMst_VehicleType!VehicleType = TxtVehicleType.Text
RsMst_VehicleType!MinTareWt = Val(TxtMinimum)
RsMst_VehicleType!MaxTareWt = Val(TxtMaximum)
RsMst_VehicleType!Remarks = Trim(TxtRemark.Text)

If IsNull(RsMst_VehicleType!CreatedBy) = True Or RsMst_VehicleType!CreatedBy = "" Then
   RsMst_VehicleType!CreatedBy = Gen_UserLoginID
   RsMst_VehicleType!CreatedDate = Now
Else
   RsMst_VehicleType!UpdatedBy = Gen_UserLoginID
   RsMst_VehicleType!UpdatedDate = Now
End If


RsMst_VehicleType.Update
Call Indata
End Sub
Private Function GetVehicleTypeID() As Double
Dim Retval As Double
tmp = "Select max(VehicleTypeID) from Mst_VehicleType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetVehicleTypeID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.TextMatrix(0, 0) = "Vehicle Type ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Vehicle Type "
MSHFlexGrid1.TextMatrix(0, 2) = "Minimum Tare Weight"
MSHFlexGrid1.TextMatrix(0, 3) = "Maximum Tare Weight"
MSHFlexGrid1.TextMatrix(0, 4) = "Remarks"

MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 600

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_VehicleType.MoveFirst
RsMst_VehicleType.Find "VehicleTypeID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_VehicleType.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtVehicleType_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtVehicleType")
End If
End Sub


