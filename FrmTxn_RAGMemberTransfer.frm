VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_RAGMemberTransfer 
   Caption         =   "Transfer Of RAG Member"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1575
      Left            =   2325
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   7800
      Begin VB.ComboBox CmbSFromRAG 
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
         Left            =   3195
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   375
         Width           =   5865
      End
      Begin VB.ComboBox CmbSToRAG 
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
         Left            =   9165
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   375
         Width           =   5625
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   11040
         TabIndex        =   13
         Top             =   840
         Width           =   1245
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   12360
         TabIndex        =   14
         Top             =   840
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   16
         Top             =   1740
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   360
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
         Format          =   61800449
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   1650
         TabIndex        =   43
         Top             =   360
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
         Format          =   61800449
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label21 
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
         Left            =   352
         TabIndex        =   45
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label19 
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
         Left            =   1980
         TabIndex        =   44
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "From RAG Employee"
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
         Left            =   5025
         TabIndex        =   41
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label Label4 
         Caption         =   "To RAG Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10995
         TabIndex        =   40
         Top             =   150
         Width           =   1965
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
         Left            =   6840
         TabIndex        =   18
         Top             =   840
         Width           =   2340
      End
      Begin VB.Label Label22 
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
         Left            =   9225
         TabIndex        =   17
         Top             =   847
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   12720
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
         Height          =   975
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1635
         Width           =   9975
      End
      Begin VB.CommandButton CmdToRAG 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1188
         Width           =   375
      End
      Begin VB.ComboBox CmbToRAG 
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
         Left            =   2145
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1165
         Width           =   9780
      End
      Begin VB.TextBox TxtTransferID 
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
         Height          =   375
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   210
         Width           =   2130
      End
      Begin VB.ComboBox CmbFromRAG 
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
         Left            =   2145
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   695
         Width           =   9780
      End
      Begin VB.TextBox TxtFromRAG 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   690
         Width           =   9735
      End
      Begin MSComCtl2.DTPicker TxtTransferDate 
         Height          =   375
         Left            =   6300
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   210
         Width           =   1440
         _ExtentX        =   2540
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
         Format          =   61800449
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtToRAG 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1165
         Width           =   9735
      End
      Begin VB.Label Label3 
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   30
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "To RAG Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         TabIndex        =   28
         Top             =   1105
         Width           =   1440
      End
      Begin VB.Label Label8 
         Caption         =   "Transfer Date"
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
         Left            =   4740
         TabIndex        =   26
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Transfer ID"
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
         Left            =   150
         TabIndex        =   24
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "From RAG Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   150
         TabIndex        =   23
         Top             =   605
         Width           =   2010
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1020
      Left            =   120
      TabIndex        =   19
      Top             =   8520
      Width           =   12720
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
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
         Left            =   165
         TabIndex        =   0
         Top             =   165
         Width           =   2415
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
         Height          =   375
         Left            =   2580
         TabIndex        =   4
         Top             =   165
         Width           =   2415
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
         Height          =   375
         Left            =   4995
         TabIndex        =   5
         Top             =   165
         Width           =   2535
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
         Height          =   375
         Left            =   7530
         TabIndex        =   6
         Top             =   165
         Width           =   2415
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
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
         Left            =   9945
         TabIndex        =   7
         Top             =   165
         Width           =   2415
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
         Height          =   375
         Left            =   165
         TabIndex        =   12
         Top             =   540
         Width           =   2415
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
         Height          =   375
         Left            =   2580
         TabIndex        =   11
         Top             =   540
         Width           =   2415
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
         Height          =   375
         Left            =   4995
         TabIndex        =   10
         Top             =   540
         Width           =   2535
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
         Height          =   375
         Left            =   7530
         TabIndex        =   9
         Top             =   540
         Width           =   2415
      End
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
         Height          =   375
         Left            =   9945
         TabIndex        =   8
         Top             =   540
         Width           =   2415
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5685
      Left            =   120
      TabIndex        =   31
      Top             =   2760
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   10028
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CMP Details"
      TabPicture(0)   =   "FrmTxn_RAGMemberTransfer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ChkCheckAll"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Log History"
      TabPicture(1)   =   "FrmTxn_RAGMemberTransfer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "TxtUpdated"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtCreated"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
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
         Height          =   375
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   840
         Width           =   9495
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
         Height          =   375
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1395
         Width           =   9495
      End
      Begin VB.CheckBox ChkCheckAll 
         Caption         =   "Select All"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   5100
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   12450
         _ExtentX        =   21960
         _ExtentY        =   8996
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         Height          =   240
         Left            =   -74520
         TabIndex        =   37
         Top             =   900
         Width           =   1695
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
         Left            =   -74550
         TabIndex        =   36
         Top             =   1425
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmTxn_RAGMemberTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Dim Edit_Flg, mFromEmployeeID, mToEmployeeID As Variant
Dim TempID  As Variant

Private Sub CmbFromRAG_GotFocus()
tmp = CmbFromRAG.Text
If Gen_WcEmployeeID = "" Then
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7")
Else
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7 And " & Gen_WcEmployeeID)
End If
CmbFromRAG.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbFromRAG.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbFromRAG.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbFromRAG.Text = tmp
End Sub

Private Sub CmbFromRAG_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbFromRAG.Text = "" Then
   mFromEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbFromRAG, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection."
   CmbFromRAG.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbFromRAG.SetFocus
   Exit Sub
End If

mFromEmployeeID = RsMst_Employee!EmployeeID
Call FillCMPMDetails(mFromEmployeeID)

End Sub


Private Sub CmbSFromRAG_GotFocus()
tmp = CmbSFromRAG.Text
If Gen_WcEmployeeID = "" Then
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7 ")
Else
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7 And " & Gen_WcEmployeeID)
End If
CmbSFromRAG.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSFromRAG.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSFromRAG.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSFromRAG.Text = tmp
End Sub

Private Sub CmbSFromRAG_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSFromRAG.Text = "" Then
   mSFromEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSFromRAG, "~")
If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection."
   CmbSFromRAG.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSFromRAG.SetFocus
   Exit Sub
End If
mSFromEmployeeID = RsMst_Employee!EmployeeID

End Sub

Private Sub CmbSToRAG_GotFocus()
tmp = CmbSToRAG.Text
If Gen_WcEmployeeID = "" Then
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7")
Else
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7 And " & Gen_WcEmployeeID)
End If
CmbSToRAG.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbSToRAG.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbSToRAG.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbSToRAG.Text = tmp

End Sub

Private Sub CmbSToRAG_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbSToRAG.Text = "" Then
   mSToEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbSToRAG, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection."
   CmbSToRAG.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbSToRAG.SetFocus
   Exit Sub
End If
mSToEmployeeID = RsMst_Employee!EmployeeID
End Sub

Private Sub CmbToRAG_GotFocus()
tmp = CmbToRAG.Text
If Gen_WcEmployeeID = "" Then
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7")
Else
   Call TableMst_Employee("Where Flag = 'Y' and DesignationID=7 And " & Gen_WcEmployeeID)
End If
CmbToRAG.Clear
If RsMst_Employee.RecordCount = 0 Then
   MsgBox "No Employee Found!!!"
   CmbToRAG.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Employee.RecordCount
    CmbToRAG.AddItem RsMst_Employee!EmployeeName & "~" & RsMst_Employee!EmployeeNo
    RsMst_Employee.MoveNext
Next
CmbToRAG.Text = tmp

End Sub

Private Sub CmbToRAG_LostFocus()
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

If CmbToRAG.Text = "" Then
   mToEmployeeID = Null
   Exit Sub
End If

strSplitString = Split(CmbToRAG, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection."
   CmbToRAG.SetFocus
   Exit Sub
End If

EmpName_ = Trim(strSplitString(0))
EmpNo_ = Trim(strSplitString(1))

RsMst_Employee.MoveFirst
RsMst_Employee.Find "EmployeeNo = '" & EmpNo_ & "'"

If RsMst_Employee.EOF Then
   MsgBox "Invalid Selection "
   CmbToRAG.SetFocus
   Exit Sub
End If
mToEmployeeID = RsMst_Employee!EmployeeID
End Sub
Private Sub FillCMPMDetails(mFromEmployeeID_ As Variant)

Call Grid_HeadDetails
'tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID,ME.EmployeeID"
'tmp = tmp & " from Mst_EmpCMPMapping ME"
'tmp = tmp & " Inner Join Mst_CMP MC On ME.CMPID=MC.CMPID"
'tmp = tmp & " Inner Join Mst_Godown MG On MC.CMPID=MG.CMPID"
'tmp = tmp & " Inner Join Mst_UnitType MU On MG.UnitTypeID=MU.UnitTypeID"
'tmp = tmp & " Inner Join Mst_Employee ME1 On ME1.EmployeeID=ME.EmployeeID"
'tmp = tmp & " Where ME1.ConnectedToRAG = " & mFromEmployeeID_ & ""
'tmp = tmp & " Order By MC.CMPName,MG.GodownNo"


'tmp = "Select Distinct ML.LocationName,MC.CMPName,MC.Address,MC.CMPID"
'tmp = tmp & " from Mst_EmpCMPMapping ME"
'tmp = tmp & " Inner Join Mst_CMP MC On ME.CMPID=MC.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
'tmp = tmp & " Inner Join Mst_Employee ME1 On ME1.EmployeeID=ME.EmployeeID"
'tmp = tmp & " Where ME1.ConnectedToRAG = " & mFromEmployeeID_ & ""
'tmp = tmp & " Order By MC.CMPName"

tmp = " Select Distinct ME1.EmployeeName +'~'+ME1.EmployeeNo as 'CMPM',ML.LocationName,MC.CMPName,MC.Address,ME.EmployeeID "
tmp = tmp & " from Mst_EmpCMPMapping ME "
tmp = tmp & " Inner Join Mst_CMP MC On ME.CMPID=MC.CMPID "
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID=ML.LocationID "
tmp = tmp & " Inner Join Mst_Employee ME1 On ME1.EmployeeID=ME.EmployeeID "
tmp = tmp & " Where ME1.ConnectedToRAG = " & mFromEmployeeID_ & " "
tmp = tmp & " Order By ML.LocationName "


Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid2.TextMatrix(I, c + 1) = TmpRs.Fields(c)
       Next
       MSHFlexGrid2.Row = I
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellFontName = "Wingdings"
       MSHFlexGrid2.CellFontSize = 12
       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       TmpRs.MoveNext
   Next
Else
   MsgBox "No CMPM Details Found !!"
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_RAGMemberTransfer.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_RAGMemberTransfer.xls")
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

Private Sub CmdGo_Click()
Dim Wc As Variant

Call Grid_Head

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "'From Date' must be less than 'To Date' always !!!"
   TxtFromDate.SetFocus
   Exit Sub
End If

tmp = ""

tmp = "Select A.TransferID,A.TransferDate,ME.EmployeeName + '~'+ME.EmployeeNo As FromID,ME1.EmployeeName + '~'+ ME1.EmployeeNo As ToID,A.Remark"
tmp = tmp & " From Txn_RAGMemberTransfer A"
tmp = tmp & " Inner Join Mst_Employee ME On ME.EmployeeID=A.FromEmployeeID"
tmp = tmp & " Inner Join Mst_Employee ME1 On ME1.EmployeeID=A.ToEmployeeID"

Wc = ""

If Gen_DBType = "MDB" Then
   Wc = " Where A.TransferDate Between #" & TxtFromDate.Value & "# And #" & TxtToDate.Value & "# "
Else
   Wc = " Where A.TransferDate Between '" & TxtFromDate.Value & "' And '" & TxtToDate.Value & "' "
End If

If CmbSFromRAG.Text <> "" Then
   Wc = Wc & " And ME.EmployeeName + '~'+ME.EmployeeNo = '" & CmbSFromRAG.Text & "'"
Else
   If Gen_WcEmployeeID <> "" Then
      Wc = Wc & " And ME." & Gen_WcEmployeeID
   End If
End If

If CmbSToRAG.Text <> "" Then
   Wc = Wc & " And ME1.EmployeeName + '~'+ME1.EmployeeNo = '" & CmbSToRAG.Text & "'"
Else
   If Gen_WcEmployeeID <> "" Then
      Wc = Wc & " And ME." & Gen_WcEmployeeID
   End If
End If

tmp = tmp & Wc

Call TmpTable

For I = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!TransferID
    MSHFlexGrid1.TextMatrix(I, 1) = Format(TmpRs!TransferDate, "dd-mmm-yyyy")
    MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!FromID
    MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!ToID
    MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!Remark
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    TmpRs.MoveNext
Next
Label22.Caption = TmpRs.RecordCount
End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 5
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .TextMatrix(0, 0) = "Transfer ID"
     .TextMatrix(0, 1) = "Transfer Date"
     .TextMatrix(0, 2) = "From RAG Employee"
     .TextMatrix(0, 3) = "To RAG Employee"
     .TextMatrix(0, 4) = "Remark"
     .ColWidth(0) = 1200
     .ColWidth(1) = 1400
     .ColWidth(2) = 3500
     .ColWidth(3) = 3500
     .ColWidth(4) = 5000
     .RowHeight(0) = 400
End With

End Sub



Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_RAGMemberTransfer.MoveFirst
RsTxn_RAGMemberTransfer.Find "TransferID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsTxn_RAGMemberTransfer.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub CmdToRAG_Click()
FrmMst_Employee.Show
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub



Private Sub Form_Load()
Call Grid_HeadDetails
MSHFlexGrid2.Row = 0
MSHFlexGrid2.Col = 0

With ChkCheckAll
     .Font.Size = 10
     .Left = MSHFlexGrid2.Left + MSHFlexGrid2.CellLeft
     .Top = MSHFlexGrid2.Top + MSHFlexGrid2.CellTop
     .Width = MSHFlexGrid2.CellWidth - 10
     .Height = MSHFlexGrid2.CellHeight - 10
     .ZOrder
End With

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True

Call TableTxn_RAGMemberTransfer(" Order by TransferID ")

If RsTxn_RAGMemberTransfer.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Private Sub Grid_HeadDetails()
With MSHFlexGrid2
     .Clear
     .Rows = 2
     .Cols = 6
     .FixedCols = 0
     .Font.Size = 10
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(0, 0) = "Select"
'     .TextMatrix(0, 1) = "Location"
'     .TextMatrix(0, 2) = "CMP Name"
''     .TextMatrix(0, 2) = "Godown No"
'     .TextMatrix(0, 3) = "Address"
''     .TextMatrix(0, 4) = "Type"
''     .TextMatrix(0, 4) = "Area"
''     .TextMatrix(0, 6) = "GodownID"
'
'     .TextMatrix(0, 4) = "CMPID"
     
     .TextMatrix(0, 1) = "CMPM Name"
     .TextMatrix(0, 2) = "Location"
     .TextMatrix(0, 3) = "CMP Name"
     .TextMatrix(0, 4) = "Address"
     .TextMatrix(0, 5) = "EmployeeID"
     
     
'     .TextMatrix(0, 6) = "EmployeeID"
     .ColWidth(0) = 1000
     .ColWidth(1) = 3000
     .ColWidth(2) = 1800
     .ColWidth(3) = 2000
     .ColWidth(4) = 4200 '1600
     .ColWidth(5) = 0
'     .ColWidth(6) = 0
'     .ColWidth(7) = 0
     .RowHeight(0) = 500
End With
End Sub
Private Sub ChkCheckAll_Click()
If SaveCmd.Enabled = False Then
   ChkCheckAll.Value = 0
   Exit Sub
End If
If MSHFlexGrid2.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If
MSHFlexGrid2.Col = 0
If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid2.Rows - 2
       'MSHFlexGrid2.Row = i
       If ChkPrevselection(I) = False Then
          If MSHFlexGrid2.TextMatrix(I, 0) <> strChecked Then
             MSHFlexGrid2.TextMatrix(I, 0) = strChecked
          End If
       End If
'       If MSHFlexGrid2.TextMatrix(i, 0) <> strChecked Then
'          MSHFlexGrid2.TextMatrix(i, 0) = strChecked
'       End If
   Next
Else
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.Row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
   Next
End If

End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim x As Variant
x = MSHFlexGrid2.RowSel
If MSHFlexGrid2.TextMatrix(x, 1) = "" Then Exit Sub

If MSHFlexGrid2.TextMatrix(x, 0) = strChecked Then
   MSHFlexGrid2.TextMatrix(x, 0) = strUnChecked
Else
   If ChkPrevselection(x) = False Then
      MSHFlexGrid2.TextMatrix(x, 0) = strChecked
   End If
End If

End Sub
Private Function ChkPrevselection(x_ As Variant) As Boolean
Dim Retval As Boolean
Retval = False

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(x_, 5) = MSHFlexGrid2.TextMatrix(I, 5) Then
       If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
          Retval = True
          Exit For
       End If
    End If
Next

ChkPrevselection = Retval
End Function
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtTransferID.Text = ""
TxtTransferDate.Value = Date

CmbFromRAG.Visible = True
CmbToRAG.Visible = True

CmbFromRAG.Text = ""
CmbToRAG.Text = ""

TxtFromRAG.Visible = False
TxtToRAG.Visible = False
CmdToRAG.Visible = True
TxtRemark.Text = ""
TxtCreated = ""
TxtUpdated = ""
Frame1.Enabled = True
Call Grid_HeadDetails
If RsTxn_RAGMemberTransfer.RecordCount > 0 Then
   CmbFromRAG.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub DeleteCmd_Click()

'Dim ans As Variant
'On Error GoTo msgError
'
'
'
'tmp = "Select Flag from Mst_Employee Where EmployeeID=" & mFromEmployeeID & ""
'Call TmpTable
'If LCase(TmpRs!Flag) = "n" Then
'   MsgBox "You Can't Delete This Transfer because 'From RAG Employee' is Deactive !!"
'   Exit Sub
'End If
'
'ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
'If ans = vbYes Then
'
'   TempID = ""
'   For i = 1 To MSHFlexGrid2.Rows - 1
'       If MSHFlexGrid2.TextMatrix(i, 0) = strChecked Then
'          If TempID = "" Then
'             TempID = "(" & MSHFlexGrid2.TextMatrix(i, 5)
'          Else
'             TempID = TempID & "," & MSHFlexGrid2.TextMatrix(i, 5)
'          End If
'       End If
'   Next
'   TempID = TempID & ")"
'
'   tmp = "Update Mst_Employee Set ConnectedToRAG=" & mFromEmployeeID & " Where ConnectedToRAG= " & mToEmployeeID & "  And EmployeeID In " & TempID & ""
'   Call TmpTable
'
'   tmp = "Delete Txn_RAGMemberTransferDetail Where TransferID=" & Val(TxtTransferID.Text) & ""
'   Call TmpTable
'
'   RsTxn_RAGMemberTransfer.Delete
'   RsTxn_RAGMemberTransfer.Update
'   If RsTxn_RAGMemberTransfer.RecordCount = 0 Then
'      Call AddCmd_Click
'      Exit Sub
'   End If
'   RsTxn_RAGMemberTransfer.MoveNext
'   If RsTxn_RAGMemberTransfer.EOF() Then
'      RsTxn_RAGMemberTransfer.MoveLast
'   End If
'   Call InData
'End If
'
'Exit Sub
'msgError:
'MsgBox "Child Record Present "
'RsTxn_RAGMemberTransfer.CancelUpdate
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
'Call GButtonLock(Me, False)
'
'TxtFromRAG.Visible = False
'CmbFromRAG.Visible = True
'CmbFromRAG.Text = TxtFromRAG.Text
'
'TxtToRAG.Visible = False
'CmbToRAG.Visible = True
'CmbToRAG.Text = TxtToRAG.Text
'CmdToRAG.Visible = True
'Frame1.Enabled = True
End Sub

Private Sub Indata()
TxtTransferID.Text = ""
CmbFromRAG.Visible = False
CmbToRAG.Visible = False
TxtFromRAG.Visible = True
TxtToRAG.Visible = True
CmdToRAG.Visible = False
TxtFromRAG.Text = ""
TxtToRAG.Text = ""
TxtRemark.Text = ""
Call Grid_HeadDetails
Frame1.Enabled = False

TxtTransferID.Text = RsTxn_RAGMemberTransfer!TransferID
TxtTransferDate.Value = RsTxn_RAGMemberTransfer!TransferDate
mFromEmployeeID = RsTxn_RAGMemberTransfer!FromEmployeeID
TxtFromRAG.Text = GetEmployeeNonName(RsTxn_RAGMemberTransfer!FromEmployeeID)
mToEmployeeID = RsTxn_RAGMemberTransfer!ToEmployeeID
TxtToRAG.Text = GetEmployeeNonName(RsTxn_RAGMemberTransfer!ToEmployeeID)
TxtRemark.Text = RsTxn_RAGMemberTransfer!Remark
TxtCreated = IIf(IsNull(RsTxn_RAGMemberTransfer!CreatedBy), "", RsTxn_RAGMemberTransfer!CreatedBy) & " :- " & IIf(IsNull(RsTxn_RAGMemberTransfer!CreatedDate), "", RsTxn_RAGMemberTransfer!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_RAGMemberTransfer!UpdatedBy), "", RsTxn_RAGMemberTransfer!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_RAGMemberTransfer!UpdatedDate), "", RsTxn_RAGMemberTransfer!UpdatedDate)
Call InDataDetail
Call GButtonLock(Me, True)
End Sub
Private Sub InDataDetail()

'tmp = "Select MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.Area,MG.GodownID"
'tmp = tmp & " from Txn_RAGMemberTransferDetail TR"
'tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TR.GodownID"
'tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
'tmp = tmp & " Inner Join Mst_UnitType MU On MG.UnitTypeID=MU.UnitTypeID"
'tmp = tmp & " Where TR.TransferID = " & Val(TxtTransferID.Text) & ""
'tmp = tmp & " Order By MC.CMPName,MG.GodownNo"

Call Grid_HeadDetails


'tmp = "Select ML.LocationName,MC.CMPName,MC.Address,MC.CMPID"
'tmp = tmp & " from Txn_RAGMemberTransferDetail TR"
'tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=TR.GodownID"
'tmp = tmp & " Inner Join Mst_CMP MC On TR.CMPID=MC.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID=ML.LocationID"
'tmp = tmp & " Where TR.TransferID = " & Val(TxtTransferID.Text) & ""
'tmp = tmp & " Order By MC.CMPName"


tmp = " Select Distinct ME1.EmployeeName +'~'+ME1.EmployeeNo as 'CMPM',ML.LocationName,MC.CMPName,MC.Address,ME.EmployeeID "
tmp = tmp & " from Txn_RAGMemberTransferDetail TR "
tmp = tmp & " Inner Join Mst_EmpCMPMapping ME On TR.EmployeeID = ME.EmployeeID "
tmp = tmp & " Inner Join Mst_Employee ME1 On ME1.EmployeeID=ME.EmployeeID "
tmp = tmp & " Inner Join Mst_CMP MC On ME.CMPID=MC.CMPID "
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID=ML.LocationID "
tmp = tmp & " Where TR.TransferID =  " & Val(TxtTransferID.Text) & " "
tmp = tmp & " Order By ML.LocationName "


Call TmpTable


For I = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid2.TextMatrix(I, c + 1) = TmpRs.Fields(c)
    Next
    MSHFlexGrid2.Row = I
    MSHFlexGrid2.Col = 0
    MSHFlexGrid2.CellFontName = "Wingdings"
    MSHFlexGrid2.CellFontSize = 12
    MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
    If ChkPrevselection(I) = False Then
       MSHFlexGrid2.TextMatrix(I, 0) = strChecked
    Else
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
    End If
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    TmpRs.MoveNext
Next
End Sub
Private Sub SaveCmd_Click()

If TxtTransferDate > Date Then
   MsgBox "Future Transfer date not allowed !!"
   TxtTransferDate.SetFocus
   Exit Sub
End If

If CmbFromRAG.Text = "" Then
   MsgBox "Blank 'From RAG Employee' Name Not Allowed !!"
   CmbFromRAG.SetFocus
   Exit Sub
End If

If CmbToRAG.Text = "" Then
   MsgBox "Blank 'To RAG Employee' Name Not Allowed !!!"
   CmbToRAG.SetFocus
   Exit Sub
End If

If CmbFromRAG.Text = CmbToRAG.Text Then
   MsgBox "Transfer To Same Employee Not Allowed !!!"
   CmbToRAG.SetFocus
   Exit Sub
End If

If TxtRemark = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If CheckGrid = False Then
   MsgBox "Please Select CMP Details From List!!!"
   Exit Sub
End If



If Edit_Flg = "A" Then
'   If RsTxn_RAGMemberTransfer.RecordCount > 0 Then
'      RsTxn_RAGMemberTransfer.MoveFirst
'   End If
'   RsTxn_RAGMemberTransfer.Filter = "FromEmployeeID=" & mFromEmployeeID & " And ToEmployeeID =" & mToEmployeeID & ""
'   If RsTxn_RAGMemberTransfer.RecordCount > 0 Then
'      MsgBox "Duplicate Transfer Not Allowed !!"
'      RsTxn_RAGMemberTransfer.Filter = ""
'      Exit Sub
'   End If
'   RsTxn_RAGMemberTransfer.Filter = ""
   RsTxn_RAGMemberTransfer.AddNew
   RsTxn_RAGMemberTransfer!TransferID = GetMaxID("TransferID", "Txn_RAGMemberTransfer")
   RsTxn_RAGMemberTransfer!G_UID = GetGUID
   TxtTransferID.Text = RsTxn_RAGMemberTransfer!TransferID
Else
'   RsTxn_RAGMemberTransfer.MoveFirst
'   RsTxn_RAGMemberTransfer.Find "TransferID=" & Val(TxtTransferID.Text) & ""
End If
RsTxn_RAGMemberTransfer!TransferDate = TxtTransferDate.Value
RsTxn_RAGMemberTransfer!FromEmployeeID = mFromEmployeeID
RsTxn_RAGMemberTransfer!ToEmployeeID = mToEmployeeID
RsTxn_RAGMemberTransfer!Remark = Trim(TxtRemark.Text)

If IsNull(RsTxn_RAGMemberTransfer!CreatedBy) = True Or RsTxn_RAGMemberTransfer!CreatedBy = "" Then
   RsTxn_RAGMemberTransfer!CreatedBy = Gen_UserLoginID
   RsTxn_RAGMemberTransfer!CreatedDate = Now
Else
   RsTxn_RAGMemberTransfer!UpdatedBy = Gen_UserLoginID
   RsTxn_RAGMemberTransfer!UpdatedDate = Now
End If

RsTxn_RAGMemberTransfer.Update
Call SaveDetails
Call UpdateEmployeeMaster

Call Indata
End Sub
Private Function GetMaxID(mCol_ As Variant, mTbl As Variant)
tmp = "Select max(" & mCol_ & ")As ID from " & mTbl & ""
Call TmpTable
GetMaxID = IIf(IsNull(TmpRs!ID), 1, TmpRs!ID + 1)
End Function
Private Sub SaveDetails()
tmp = "Delete Txn_RAGMemberTransferDetail Where TransferID=" & Val(TxtTransferID.Text) & ""
Call TmpTable

Call TableTxn_RAGMemberTransferDetail
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       RsTxn_RAGMemberTransferDetail.AddNew
       RsTxn_RAGMemberTransferDetail!RAGMemberTransferDetailID = GetMaxID("RAGMemberTransferDetailID", "Txn_RAGMemberTransferDetail")
       RsTxn_RAGMemberTransferDetail!TransferID = Val(TxtTransferID.Text)
       RsTxn_RAGMemberTransferDetail!EmployeeID = Val(MSHFlexGrid2.TextMatrix(I, 5))
       RsTxn_RAGMemberTransferDetail!G_UID = GetGUID
       RsTxn_RAGMemberTransferDetail.Update
    End If
Next
End Sub
Private Sub UpdateEmployeeMaster()
TempID = ""
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       If TempID = "" Then
          TempID = "(" & MSHFlexGrid2.TextMatrix(I, 5)
       Else
          TempID = TempID & "," & MSHFlexGrid2.TextMatrix(I, 5)
       End If
    End If
Next
TempID = TempID & ")"

tmp = "Update Mst_Employee Set ConnectedToRAG=" & mToEmployeeID & " Where ConnectedToRAG= " & mFromEmployeeID & "  And EmployeeID In " & TempID & ""
Call TmpTable

End Sub

Private Function CheckGrid() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckGrid = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_RAGMemberTransfer.MoveNext
If RsTxn_RAGMemberTransfer.EOF Then
   RsTxn_RAGMemberTransfer.MoveLast
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
RsTxn_RAGMemberTransfer.MovePrevious
If RsTxn_RAGMemberTransfer.BOF Then
   RsTxn_RAGMemberTransfer.MoveFirst
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
RsTxn_RAGMemberTransfer.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_RAGMemberTransfer.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_RAGMemberTransfer.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_RAGMemberTransfer.MoveFirst
   RsTxn_RAGMemberTransfer.Find "TransferID = " & Val(TxtTransferID.Text) & ""
   If RsTxn_RAGMemberTransfer.EOF Then
      RsTxn_RAGMemberTransfer.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200

RsTxn_RAGMemberTransfer.MoveFirst
Call Grid_Head
TxtFromDate.Value = Date - 365
TxtToDate.Value = Date
CmbSFromRAG.Text = ""
CmbSToRAG.Text = ""

Label22.Caption = ""


End Sub


