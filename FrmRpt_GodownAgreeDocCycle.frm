VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_GodownAgreeDocCycle 
   Caption         =   "Godown Agreement Document Cycle Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbForwardedTo 
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
         TabIndex        =   11
         Top             =   1950
         Width           =   6420
      End
      Begin VB.ComboBox CmbStatus 
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
         Left            =   9780
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1275
         Width           =   5100
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   14775
         Begin VB.TextBox TxtSAgreeID 
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
            Left            =   13080
            MaxLength       =   12
            TabIndex        =   6
            Top             =   240
            Width           =   1530
         End
         Begin VB.OptionButton OptS 
            Caption         =   "Agreement Start Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   120
            TabIndex        =   0
            Top             =   135
            Width           =   1335
         End
         Begin VB.OptionButton OptN 
            Caption         =   "Doc Lifecycle Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   4740
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.OptionButton OptE 
            Caption         =   "Agreement End Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1515
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
         Begin VB.OptionButton OptA 
            Caption         =   "Agreement Effect Date"
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
            Left            =   2865
            TabIndex        =   2
            Top             =   210
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker SSTo 
            Height          =   375
            Left            =   9870
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   180
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
            Format          =   111542273
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SSFrom 
            Height          =   375
            Left            =   7485
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   180
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
            Format          =   111542273
            CurrentDate     =   36494
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Doc Lifycycle ID"
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
            Left            =   11460
            TabIndex        =   27
            Top             =   285
            Width           =   1440
         End
         Begin VB.Label Label7 
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
            Left            =   9030
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
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
            Left            =   6435
            TabIndex        =   25
            Top             =   240
            Width           =   975
         End
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
         Left            =   11610
         TabIndex        =   12
         Top             =   2055
         Width           =   1185
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
         Left            =   12795
         TabIndex        =   14
         Top             =   2055
         Width           =   1095
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   6555
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1275
         Width           =   3195
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
         Left            =   3330
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1275
         Width           =   3195
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
         Left            =   13890
         TabIndex        =   15
         Top             =   2055
         Width           =   1050
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
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1275
         Width           =   3195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forwarded To"
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
         Left            =   2603
         TabIndex        =   29
         Top             =   1695
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   12053
         TabIndex        =   28
         Top             =   1005
         Width           =   555
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
         Left            =   8535
         TabIndex        =   23
         Top             =   2085
         Width           =   2295
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
         Left            =   10935
         TabIndex        =   22
         Top             =   2085
         Width           =   90
      End
      Begin VB.Label LblLocationID 
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
         Left            =   4545
         TabIndex        =   21
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   7935
         TabIndex        =   20
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   1140
         TabIndex        =   19
         Top             =   1020
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7170
      Left            =   120
      TabIndex        =   16
      Top             =   2655
      Width           =   15060
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6540
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   11536
         _Version        =   393216
         FixedCols       =   0
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
         TabIndex        =   17
         Top             =   6795
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownAgreeDocCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wc, Wc1
Dim mStateID As Variant, mSLocationID As Variant, mSSMPrefix As Variant, mSEmployeeID As Variant
Dim mFlag As Variant



Private Sub CmbCMPNameName_GotFocus()
If mSLocationID = "" Then
   MsgBox "Please Select Location !!!"
   If CmbLocation.Enabled = True And CmbLocation.Visible = True Then CmbLocation.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbForwardedTo_GotFocus()
tmp = CmbForwardedTo.Text
If Gen_WcEmployeeID <> "" Then
   Call TableMst_Employee("Where " & Gen_WcEmployeeID)
Else
   Call TableMst_Employee
End If
CmbForwardedTo.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbForwardedTo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbForwardedTo.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next

CmbForwardedTo.Text = tmp
End Sub

Private Sub CmbForwardedTo_LostFocus()

Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbForwardedTo.Text = "" Then
   mSEmployeeID = Null
   Exit Sub
End If


strSplitString = Split(CmbForwardedTo, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection "
   CmbForwardedTo.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbForwardedTo.SetFocus
   Exit Sub
End If

mSEmployeeID = RsMst_Employee!EmployeeID
End Sub

Private Sub CmbCMPName_GotFocus()
tmp = CmbCMPName.Text

If CmbLocation.Text = "" Then
   MsgBox "Please Select Location First !!!"
   CmbLocation.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
Else
   tmp1 = "  LocationID = " & mSLocationID & ""
End If

Call TableMst_CMP(" where  " & tmp1)

CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp

End Sub


Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
End Sub


Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

If CmbState.Text = "" Then
   MsgBox "Please Select State First !!!"
   CmbState.SetFocus
   Exit Sub
End If

If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where Stateid = " & mStateID & "")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
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

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   CmbCMPName.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
CmbCMPName.Clear
mSLocationID = RsMst_Location.Fields("LocationID")
mSSMPrefix = RsMst_Location.Fields("SM_Prefix")
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
   CmbLocation.Text = ""
   CmbCMPName.Text = ""
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
   CmbCMPName.Text = ""
End If

mStateID = RsMst_State!StateID
End Sub




Private Sub CmbStatus_GotFocus()
CmbStatus.Clear
CmbStatus.AddItem "Forwarded To CMPM"
CmbStatus.AddItem "Forwarded to CMG"
CmbStatus.AddItem "Forwarded to AM"
CmbStatus.AddItem "Forwarded To Legal"
CmbStatus.AddItem "Approved"
End Sub

Private Sub CmbStatus_LostFocus()
If CmbStatus.Text = "" Then Exit Sub

If LCase(CmbStatus.Text) <> LCase("Forwarded To CMPM") And LCase(CmbStatus) <> LCase("Forwarded to CMG") And LCase(CmbStatus) <> LCase("Forwarded to AM") And LCase(CmbStatus) <> LCase("Forwarded To Legal") And LCase(CmbStatus.Text) <> LCase("Approved") Then
   MsgBox "Invalid Selection"
   CmbStatus.SetFocus
   Exit Sub
End If

If LCase(CmbStatus.Text) = "forwarded to cmg" Then
   mFlag = "g"
ElseIf LCase(CmbStatus.Text) = "forwarded to cmpm" Then
   mFlag = "p"
ElseIf LCase(CmbStatus.Text) = "forwarded to fag" Then
   mFlag = "f"
ElseIf LCase(CmbStatus.Text) = "forwarded to legal" Then
   mFlag = "l"
ElseIf LCase(CmbStatus.Text) = "approved" Then
   mFlag = "a"
ElseIf LCase(CmbStatus.Text) = "forwarded to am" Then
   mFlag = "m"
End If

End Sub

Private Sub CmdClear_Click()
OptN.Value = True
SSFrom = Date - 30
SSTo = Date
TxtSAgreeID = ""
CmbState = ""
CmbLocation = ""
CmbCMPName = ""
CmbForwardedTo = ""
CmbStatus = ""
Call GridHead
ProgressBar1.Value = 0
CmdExcel.Enabled = False
Label2.Caption = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "GodownAgreementDocumentCycleDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "GodownAgreementDocumentCycleDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
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
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub CmdGetReport_Click()

ProgressBar1.Value = 0

Call GridHead
CmdExcel.Enabled = False
Label2.Caption = ""

If SSFrom > SSTo Then
   MsgBox "From Date must be less than To Date!!!"
   SSFrom.SetFocus
   Exit Sub
End If
MsgBox "Please Wait Till Next Message !!!"

If OptS.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where TGA.StartDate Between #" & Format(SSFrom.Value, Gen_DatFormat) & "# And #" & Format(SSTo.Value, Gen_DatFormat) & "# "
   Else
      wc = " Where TGA.StartDate Between '" & Format(SSFrom.Value, Gen_DatFormat) & "' And '" & Format(SSTo.Value, Gen_DatFormat) & "' "
   End If
ElseIf OptA.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where TGA.AgreementActualStartDate Between #" & Format(SSFrom.Value, Gen_DatFormat) & "# And #" & Format(SSTo.Value, Gen_DatFormat) & "# "
   Else
      wc = " Where TGA.AgreementActualStartDate Between '" & Format(SSFrom.Value, Gen_DatFormat) & "' And '" & Format(SSTo.Value, Gen_DatFormat) & "' "
   End If
ElseIf OptE.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where TGA.ExpiryDate Between #" & Format(SSFrom.Value, Gen_DatFormat) & "# And #" & Format(SSTo.Value, Gen_DatFormat) & "# "
   Else
      wc = " Where TGA.ExpiryDate Between '" & Format(SSFrom.Value, Gen_DatFormat) & "' And '" & Format(SSTo.Value, Gen_DatFormat) & "' "
   End If
ElseIf OptN.Value = True Then
   If Gen_DBType = "MDB" Then
      wc = " Where TGA.EntryDate Between #" & Format(SSFrom.Value, Gen_DatFormat) & "# And #" & Format(SSTo.Value, Gen_DatFormat) & "# "
   Else
      wc = " Where TGA.EntryDate Between '" & Format(SSFrom.Value, Gen_DatFormat) & "' And '" & Format(SSTo.Value, Gen_DatFormat) & "' "
   End If
End If

If TxtSAgreeID.Text <> "" Then
   wc = wc & " And TGA.AgreementID=" & Val(TxtSAgreeID.Text) & ""
End If

If CmbState.Text <> "" Then
   wc = wc & " And MS.StateName='" & Trim(CmbState.Text) & "'"
End If

If CmbLocation.Text <> "" Then
   wc = wc & " And ML.LocationName='" & Trim(CmbLocation.Text) & "'"
End If



tmp = " Select Distinct TGA.AgreementID,TGA.EntryDate,MS.StateName,ML.LocationName,MUT.UnitType,"
tmp = tmp & " TGA.StartDate,TGA.ExpiryDate,TGA.AgreementActualStartDate"
tmp = tmp & " From Txn_GodownAgreeDocCycle TGA"
tmp = tmp & " Inner Join Mst_Location ML on TGA.LocationId=ML.LocationId"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join Mst_UnitType MUT on TGA.UnitTypeID=MUT.UnitTypeID"
tmp = tmp & " Inner Join Txn_GADCGodownDetail TGCD on TGA.AgreementID=TGCD.AgreementID"
tmp = tmp & " Inner Join Mst_GodownInspection MGI on TGCD.InspectionID=MGI.InspectionID"
tmp = tmp & " Inner Join Mst_CMP MC on MGI.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Txn_GADCDocumentAttachDetail TGDAD on TGA.AgreementID=TGDAD.AgreementID"
tmp = tmp & " Inner Join Mst_DocumentType MDT on TGDAD.DocumentTypeID=MDT.DocumentTypeID"
tmp = tmp & " Inner Join Mst_Employee ME on TGDAD.ForwardID=ME.EmployeeId"
tmp = tmp & wc
If CmbCMPName.Text <> "" Then
   tmp = tmp & "  And Mc.CMPName='" & Trim(CmbCMPName.Text) & "'"
End If
If CmbStatus.Text <> "" Then
   tmp = tmp & " And TGDAD.Flag='" & Trim(mFlag) & "'"
End If
If CmbForwardedTo.Text <> "" Then
   tmp = tmp & " And ME.EmployeeID=" & Val(mSEmployeeID) & ""
End If
tmp = tmp & " order by TGA.AgreementID"

Call Tmp1Table

tmp = " Select TGA.AgreementID,MDT.DocumentType,TGDAD.Remarks,"
tmp = tmp & " Case TGDAD.Flag"
tmp = tmp & " WHEN 'G' THEN 'Forwarded TO CMG'"
tmp = tmp & " WHEN 'L' THEN 'Forwarded to Legal'"
tmp = tmp & " WHEN 'M' THEN 'Forwarded to AM'"
tmp = tmp & " WHEN 'P' THEN 'Forwarded to CMPM'"
tmp = tmp & " WHEN 'F' THEN 'Forwarded to FAG'"
tmp = tmp & " WHEN 'A' THEN 'Approved'"
tmp = tmp & " ELSE Null"
tmp = tmp & " END As 'Status',Me.EmployeeName+'~'+Me.EmployeeNo As 'ForwardedTo'"
tmp = tmp & " From Txn_GodownAgreeDocCycle TGA"
tmp = tmp & " Inner Join Mst_Location ML on TGA.LocationId=ML.LocationId"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join Mst_UnitType MUT on TGA.UnitTypeID=MUT.UnitTypeID"
tmp = tmp & " Inner Join Txn_GADCDocumentAttachDetail TGDAD on TGA.AgreementID=TGDAD.AgreementID"
tmp = tmp & " Inner Join Mst_DocumentType MDT on TGDAD.DocumentTypeID=MDT.DocumentTypeID"
tmp = tmp & " Inner Join Mst_Employee ME on TGDAD.ForwardID=ME.EmployeeId"
tmp = tmp & wc
If CmbStatus.Text <> "" Then
   tmp = tmp & " And TGDAD.Flag='" & Trim(mFlag) & "'"
End If
If CmbForwardedTo.Text <> "" Then
   tmp = tmp & " And ME.EmployeeID=" & Val(mSEmployeeID) & ""
End If
tmp = tmp & "  order by TGA.AgreementID"

Call Tmp2Table

tmp = " Select TGA.AgreementID,Mc.CMPName,TGCD.GodownNo,TGCD.Remarks,TGCD.InspectionID"
tmp = tmp & " From Txn_GodownAgreeDocCycle TGA"
tmp = tmp & " Inner Join Txn_GADCGodownDetail TGCD on TGA.AgreementID=TGCD.AgreementID"
tmp = tmp & " Inner Join Mst_UnitType MUT on TGA.UnitTypeID=MUT.UnitTypeID"
tmp = tmp & " Inner Join Mst_GodownInspection MGI on TGCD.InspectionID=MGI.InspectionID"
tmp = tmp & " Inner Join Mst_CMP MC on MGI.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on TGA.LocationId=ML.LocationId"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & wc
If CmbCMPName.Text <> "" Then
   tmp = tmp & "  And Mc.CMPName='" & Trim(CmbCMPName.Text) & "'"
End If
tmp = tmp & "  order by TGA.AgreementID"
Call Tmp3Table
If TmpRs1.RecordCount > 0 Then
   CmdExcel.Enabled = True
   Call dispdata
Else
   MsgBox "No record found !!!"
   Exit Sub
End If

MsgBox "Done !!!"
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

Call CmdClear_Click

End Sub

Public Sub GridHead()
With MSHFlexGrid1
     MSHFlexGrid1.Clear
     MSHFlexGrid1.Rows = 2
     MSHFlexGrid1.Cols = 16
     MSHFlexGrid1.WordWrap = True
     
     MSHFlexGrid1.TextMatrix(0, 0) = "Doc Lifecycle ID"
     MSHFlexGrid1.TextMatrix(0, 1) = "Doc Lifecycle Date"
     MSHFlexGrid1.TextMatrix(0, 2) = "State"
     MSHFlexGrid1.TextMatrix(0, 3) = "Location"
     MSHFlexGrid1.TextMatrix(0, 4) = "Agreement  Type"
     MSHFlexGrid1.TextMatrix(0, 5) = "Agreement Start Date"
     MSHFlexGrid1.TextMatrix(0, 6) = "Agreement End Date"
     MSHFlexGrid1.TextMatrix(0, 7) = "Agreement Effect Date"
     MSHFlexGrid1.TextMatrix(0, 8) = "Document Type Name"
     MSHFlexGrid1.TextMatrix(0, 9) = "Remarks"
     MSHFlexGrid1.TextMatrix(0, 10) = "Status"
     MSHFlexGrid1.TextMatrix(0, 11) = "Forwarded To"
     MSHFlexGrid1.TextMatrix(0, 12) = "CMP Name"
     MSHFlexGrid1.TextMatrix(0, 13) = "Godown No"
     MSHFlexGrid1.TextMatrix(0, 14) = "Godown Address"
     MSHFlexGrid1.TextMatrix(0, 15) = "Godown Viability"
     
     .RowHeight(0) = 800
End With
End Sub

Public Sub dispdata()
Dim j, k, I, m
j = 1: I = 1: k = 1: m = 1

Call GridHead

ProgressBar1.Max = TmpRs1.RecordCount + 2

If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
For iCtr = 1 To TmpRs1.RecordCount
    If CmbCMPName.Text = "" And CmbStatus.Text = "" And CmbForwardedTo.Text = "" Then
       For C = 0 To 7
           MSHFlexGrid1.TextMatrix(j, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
           If C = 1 Or C = 5 Or C = 6 Or C = 7 Then
              MSHFlexGrid1.TextMatrix(j, C) = Format(MSHFlexGrid1.TextMatrix(j, C), "dd-MMM-yyyy")
           End If
       Next
    End If
    If TmpRs2.RecordCount > 0 Then TmpRs2.MoveFirst
    For iCtr1 = 1 To TmpRs2.RecordCount
        If TmpRs2!AgreementID = TmpRs1!AgreementID Then
           For C = 0 To 7
               MSHFlexGrid1.TextMatrix(j, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
               If C = 1 Or C = 5 Or C = 6 Or C = 7 Then
                  MSHFlexGrid1.TextMatrix(j, C) = Format(MSHFlexGrid1.TextMatrix(j, C), "dd-MMM-yyyy")
               End If
           Next
           For C = 8 To 11
               MSHFlexGrid1.TextMatrix(j, C) = IIf(IsNull(TmpRs2.Fields(C - 7)), "", TmpRs2.Fields(C - 7))
           Next
           If j = MSHFlexGrid1.Rows - 1 Then MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           j = j + 1
           
        End If
        TmpRs2.MoveNext
    Next
    k = j
    
    If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
    j = I
    For iCtr2 = 1 To TmpRs3.RecordCount
        If TmpRs3!AgreementID = TmpRs1!AgreementID Then
           For C = 0 To 7
               MSHFlexGrid1.TextMatrix(j, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
               If C = 1 Or C = 5 Or C = 6 Or C = 7 Then
                  MSHFlexGrid1.TextMatrix(j, C) = Format(MSHFlexGrid1.TextMatrix(j, C), "dd-MMM-yyyy")
               End If
           Next
           For C = 12 To 15
               MSHFlexGrid1.TextMatrix(j, C) = IIf(IsNull(TmpRs3.Fields(C - 11)), "", TmpRs3.Fields(C - 11))
           Next
           If j = MSHFlexGrid1.Rows - 1 Then MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           j = j + 1
        End If
        TmpRs3.MoveNext
     Next
     m = j

     I = MSHFlexGrid1.Rows - 1
     j = MSHFlexGrid1.Rows - 1
     k = MSHFlexGrid1.Rows - 1
     m = MSHFlexGrid1.Rows - 1
     
ProgressBar1.Value = ProgressBar1.Value + 1
TmpRs1.MoveNext
Next
Label2.Caption = TmpRs1.RecordCount
ProgressBar1.Value = ProgressBar1.Max
End Sub
