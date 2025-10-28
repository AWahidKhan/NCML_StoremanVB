VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_StateDRCMapping 
   Caption         =   "State DRC Mapping"
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
      Height          =   885
      Left            =   120
      TabIndex        =   10
      Top             =   8280
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   7095
         TabIndex        =   13
         Top             =   195
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchState 
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
         Left            =   2280
         TabIndex        =   12
         Top             =   165
         Width           =   3645
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   6090
         TabIndex        =   11
         Top             =   195
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   14
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "Search By State Name"
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
         TabIndex        =   17
         Top             =   210
         Width           =   2895
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
         TabIndex        =   16
         Top             =   210
         Width           =   2355
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
         TabIndex        =   15
         Top             =   210
         Width           =   75
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1305
      Left            =   120
      TabIndex        =   35
      Top             =   1800
      Width           =   9120
      Begin VB.ComboBox CmbFlag 
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
         ItemData        =   "FrmMst_StateDRCMapping.frx":0000
         Left            =   840
         List            =   "FrmMst_StateDRCMapping.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   2625
      End
      Begin VB.TextBox TxtEmailID 
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
         Left            =   4800
         MaxLength       =   100
         TabIndex        =   4
         Top             =   360
         Width           =   4170
      End
      Begin VB.ComboBox CmbEmployeeID 
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
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   4635
      End
      Begin VB.CommandButton CmdAdd2List 
         Caption         =   "Add to List"
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
         Left            =   5535
         TabIndex        =   6
         Top             =   855
         Width           =   1575
      End
      Begin VB.CommandButton CmdDeleteList 
         Caption         =   "Delete from List"
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
         Left            =   7110
         TabIndex        =   36
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label Label11 
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
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   870
         Width           =   735
      End
      Begin VB.Label LblBagSizeFrom 
         Caption         =   "Email ID"
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
         Left            =   6488
         TabIndex        =   39
         Top             =   135
         Width           =   795
      End
      Begin VB.Label lblRM 
         Caption         =   "DRC Name"
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
         Left            =   1875
         TabIndex        =   38
         Top             =   120
         Width           =   1065
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   27
      Top             =   7005
      Width           =   9090
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
         Left            =   345
         TabIndex        =   0
         Top             =   240
         Width           =   1680
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
         Left            =   345
         TabIndex        =   33
         Top             =   645
         Width           =   1680
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
         Left            =   2025
         TabIndex        =   7
         Top             =   240
         Width           =   1680
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
         Left            =   2025
         TabIndex        =   32
         Top             =   645
         Width           =   1680
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
         Left            =   3705
         TabIndex        =   31
         Top             =   240
         Width           =   1680
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
         Left            =   3705
         TabIndex        =   30
         Top             =   645
         Width           =   1680
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
         Left            =   5385
         TabIndex        =   29
         Top             =   240
         Width           =   1680
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
         Left            =   5385
         TabIndex        =   28
         Top             =   645
         Width           =   1680
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
         Left            =   7065
         TabIndex        =   9
         Top             =   240
         Width           =   1680
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
         Height          =   400
         Left            =   7065
         TabIndex        =   8
         Top             =   645
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   9120
      Begin VB.TextBox TxtRemarks 
         Appearance      =   0  'Flat
         Height          =   570
         Left            =   1365
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1155
         Width           =   7170
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
         Left            =   1365
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   675
         Width           =   7170
      End
      Begin VB.CommandButton CmdStateID 
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
         Height          =   375
         Left            =   8595
         TabIndex        =   34
         Top             =   645
         Width           =   375
      End
      Begin VB.TextBox TxtStateName 
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
         Left            =   1365
         MaxLength       =   100
         TabIndex        =   22
         Top             =   675
         Width           =   7170
      End
      Begin VB.TextBox TxtStateID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1680
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
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2205
         Visible         =   0   'False
         Width           =   7155
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
         Height          =   480
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   7155
      End
      Begin VB.Label LblRemarks 
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   41
         Top             =   1290
         Width           =   705
      End
      Begin VB.Label LblStateName 
         Caption         =   "State Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   750
         Width           =   1800
      End
      Begin VB.Label LblStateID 
         Caption         =   "State ID"
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
         TabIndex        =   25
         Top             =   315
         Width           =   1800
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
         TabIndex        =   24
         Top             =   2265
         Visible         =   0   'False
         Width           =   975
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
         Left            =   120
         TabIndex        =   23
         Top             =   1980
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   3855
      Left            =   120
      TabIndex        =   37
      Top             =   3150
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   6800
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
Attribute VB_Name = "FrmMst_StateDRCMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mEmployeeID, mEmployeeEmailID, mStateID As Variant
Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.FixedCols = 1
MSHFlexGrid2.Cols = 5

MSHFlexGrid2.TextMatrix(0, 0) = "DRC Name and No"
MSHFlexGrid2.TextMatrix(0, 1) = "Email ID"
MSHFlexGrid2.TextMatrix(0, 2) = "Status "
MSHFlexGrid2.TextMatrix(0, 3) = "Employee ID"
MSHFlexGrid2.TextMatrix(0, 4) = "MappingID"

MSHFlexGrid2.ColWidth(0) = 5000
MSHFlexGrid2.ColWidth(1) = 2800
MSHFlexGrid2.ColWidth(2) = 1000
MSHFlexGrid2.ColWidth(3) = 0
MSHFlexGrid2.ColWidth(4) = 0

End Sub

Private Sub AddCmd_Click()

Edit_Flg = "A"
TxtStateID = ""

CmbStateID.Visible = True
TxtStateName.Visible = False
CmdStateID.Visible = True
CmbStateID.Text = ""

Call Grid_Head1
Call ClearDetail
Frame2.Enabled = True

Call GButtonLock(Me, False)

End Sub

Private Sub CmbFlag_LostFocus()

If CmbFlag.Text = "" Then Exit Sub


If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid Selection!!!"
   CmbFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
Call TableMst_State
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
'----
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
TxtStateID = RsMst_State!StateID
 
Call IndataDRCDetails

End Sub

Private Sub CmbEmployeeID_GotFocus()

tmp1 = CmbEmployeeID.Text

tmp = " Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,MD.Designation,ME.Flag,ME.EmployeeEmailID from Mst_Employee ME"
tmp = tmp & " Inner Join Mst_Designation MD On  ME.DesignationID = MD.DesignationID"
If CmdAdd2List.Caption = "Update List" Then
   tmp = tmp & " Where (Designation in ('DRC','Sub DRC','Associate DRC') And Flag = 'Y') Or ME.EmployeeID = " & mEmployeeID & " "
Else
   tmp = tmp & " Where Designation in ('DRC','Sub DRC','Associate DRC') And Flag = 'Y' "
End If
Call TmpTableCMPro

CmbEmployeeID.Clear
If TmpRsCMPro.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRsCMPro.RecordCount
    CmbEmployeeID.AddItem TmpRsCMPro!EmployeeName & "~" & TmpRsCMPro!EmployeeNo
    TmpRsCMPro.MoveNext
Next

CmbEmployeeID.Text = tmp1

End Sub

Private Sub CmbEmployeeID_LostFocus()

If CmbEmployeeID.Text = "" Then
   Exit Sub
End If

Dim Str_() As String
Str_ = Split(CmbEmployeeID.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbEmployeeID.SetFocus
   Exit Sub
End If
TmpRsCMPro.MoveFirst
TmpRsCMPro.Find "EmployeeNo='" & Str_(1) & "'"
If TmpRsCMPro.EOF Then
   MsgBox "Invalid selection "
   CmbEmployeeID.SetFocus
   Exit Sub
End If
mEmployeeID = TmpRsCMPro!EmployeeID
If TxtEmailID = "" Then
   TxtEmailID = IIf(IsNull(TmpRsCMPro!EmployeeEmailID), "", TmpRsCMPro!EmployeeEmailID)
End If
End Sub

Private Sub CmdAdd2List_Click()

If CmbEmployeeID.Text = "" Then
   MsgBox "Blank DRC Name not allowed!!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If

If TxtEmailID.Text = "" Then
   MsgBox "Blank Email ID not allowed!!!"
   TxtEmailID.SetFocus
   Exit Sub
End If

If CmbFlag.Text = "" Then
   MsgBox "Blank Status not allowed!!!"
   CmbFlag.SetFocus
   Exit Sub
End If

If CheckDuplicateEmployee = True Then
   MsgBox "Duplicate DRC in grid!!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If

If CmdAdd2List.Caption = "Update List" Then
   r = MSHFlexGrid2.RowSel
Else
   r = MSHFlexGrid2.Rows - 1
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
   MSHFlexGrid2.TextMatrix(r, 4) = ""
End If

MSHFlexGrid2.TextMatrix(r, 0) = CmbEmployeeID.Text
MSHFlexGrid2.TextMatrix(r, 1) = TxtEmailID
MSHFlexGrid2.TextMatrix(r, 2) = CmbFlag.Text
MSHFlexGrid2.TextMatrix(r, 3) = mEmployeeID

Call ClearDetail

CmbEmployeeID.SetFocus
End Sub
Function CheckDuplicateEmployee() As Boolean
Dim Retval As Boolean
Retval = False
If MSHFlexGrid2.Rows > 1 Then
   For u = 1 To MSHFlexGrid2.Rows - 1
       If CmbEmployeeID.Text = MSHFlexGrid2.TextMatrix(u, 0) Then
          If LCase(CmdAdd2List.Caption) = "add to list" Then
             Retval = True
             Exit For
          Else
             If u <> MSHFlexGrid2.RowSel Then
                Retval = True
             End If
          End If
       End If
   Next
End If

CheckDuplicateEmployee = Retval

End Function

Private Sub ClearDetail()

CmbEmployeeID.Text = ""
TxtEmailID = ""
CmbFlag.Text = ""
CmdAdd2List.Caption = "Add to List"

End Sub

Private Sub CmdDeleteList_Click()

If CmbEmployeeID.Text = "" Then Exit Sub
If LCase(CmdAdd2List.Caption) = "add to list" Then Exit Sub

I = MSHFlexGrid2.RowSel
If I > 0 Then
   If MSHFlexGrid2.TextMatrix(I, 1) <> "" Then
      For RO = I To MSHFlexGrid2.Rows - 2
          For c = 0 To MSHFlexGrid2.Cols - 1
              MSHFlexGrid2.TextMatrix(RO, c) = MSHFlexGrid2.TextMatrix(RO + 1, c)
          Next
      Next
      MSHFlexGrid2.Rows = MSHFlexGrid2.Rows - 1
   End If
End If

Call ClearDetail


End Sub

Private Sub EditCmd_Click()

Edit_Flg = "E"

'CmbStateID.Visible = True
'TxtStateID.Visible = False
'CmdStateID.Value = True

CmbStateID.Text = TxtStateName

Call ClearDetail
Frame2.Enabled = True

Call GButtonLock(Me, False)

End Sub

Private Sub ExitCmd_Click()
Unload Me
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
TxtSearchState = ""
Label22.Caption = ""

Call TableMst_State
If RsMst_State.RecordCount = 0 Then
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_StateDRCMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_StateDRCMapping.xls")
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


Private Sub MSHFlexGrid2_Click()

If SaveCmd.Enabled = False Then Exit Sub
Dim r As Variant
r = MSHFlexGrid2.RowSel
If r < 1 Or r = MSHFlexGrid2.Rows - 1 Then
   Call ClearDetail
   Exit Sub
End If
If MSHFlexGrid2.TextMatrix(r, 0) = "" Then
   Exit Sub
End If

CmbEmployeeID.Text = MSHFlexGrid2.TextMatrix(r, 0)
TxtEmailID = MSHFlexGrid2.TextMatrix(r, 1)
CmbFlag.Text = MSHFlexGrid2.TextMatrix(r, 2)
mEmployeeID = MSHFlexGrid2.TextMatrix(r, 3)

CmdAdd2List.Caption = "Update List"
CmbEmployeeID.SetFocus

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_State.MoveNext
If RsMst_State.EOF Then
   RsMst_State.MoveLast
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
RsMst_State.MovePrevious
If RsMst_State.BOF Then
   RsMst_State.MoveFirst
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
RsMst_State.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_State.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SaveCmd_Click()

If CmbStateID.Text = "" Then
   MsgBox "Blank State Not allowed!!!"
   CmbStateID.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 3 Then
   MsgBox "Map atleast one DRC to State!!!"
   CmbEmployeeID.SetFocus
   Exit Sub
End If


RsMst_State.MoveFirst
RsMst_State.Find "StateID = " & TxtStateID
RsMst_State!Remarks = Trim(TxtRemarks)
RsMst_State.Update


Call SaveDRCDetails

RsMst_State.MoveFirst
RsMst_State.Find "StateID = " & TxtStateID
 
Call GButtonLock(Me, True)

Call Indata


End Sub


Private Sub SaveDRCDetails()

Call DeleteDRCDetails

Call TableMst_StateDRCMapping(" Where StateID = " & TxtStateID & " ")

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 4) = "" Then
       RsMst_StateDRCMapping.AddNew
       RsMst_StateDRCMapping!MappingID = GetMaxMappingID
       RsMst_StateDRCMapping!G_UID = GetGUID
       RsMst_StateDRCMapping!StateID = TxtStateID
    Else
       RsMst_StateDRCMapping.MoveFirst
       RsMst_StateDRCMapping.Find "MappingID = " & MSHFlexGrid2.TextMatrix(I, 4) & ""
    End If
    RsMst_StateDRCMapping!EmployeeID = MSHFlexGrid2.TextMatrix(I, 3)
    RsMst_StateDRCMapping!EmailID = MSHFlexGrid2.TextMatrix(I, 1)
    RsMst_StateDRCMapping!Flag = Left(MSHFlexGrid2.TextMatrix(I, 2), 1)
    
    If IsNull(RsMst_StateDRCMapping!CreatedBy) = True Or RsMst_StateDRCMapping!CreatedBy = "" Then
       RsMst_StateDRCMapping!CreatedBy = Gen_UserLoginID
       RsMst_StateDRCMapping!CreatedDate = Now
    Else
       RsMst_StateDRCMapping!UpdatedBy = Gen_UserLoginID
       RsMst_StateDRCMapping!UpdatedDate = Now
    End If
    RsMst_StateDRCMapping.Update
    RsMst_StateDRCMapping.Requery
Next

End Sub
Private Function GetMaxMappingID()
Dim Retval As Double
tmp = "Select max(MappingID) from Mst_StateDRCMapping "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxMappingID = Retval
End Function
Private Sub DeleteDRCDetails()

mDeleteRow = ""

For I = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(I, 1) = "" Then Exit For
    If MSHFlexGrid2.TextMatrix(I, 4) <> "" Then
       If mDeleteRow = "" Then
          mDeleteRow = "(" & Val(MSHFlexGrid2.TextMatrix(I, 4))
       Else
          mDeleteRow = mDeleteRow & "," & Val(MSHFlexGrid2.TextMatrix(I, 4))
       End If
    End If
Next

If mDeleteRow <> "" Then
   mDeleteRow = mDeleteRow & ")"
   tmp = "Delete From Mst_StateDRCMapping Where MappingID not in " & mDeleteRow & " And StateID = " & TxtStateID & ""
   Call TmpTable
End If

End Sub

Public Sub Indata()

CmbStateID.Visible = False
CmdStateID.Visible = False
TxtStateName.Visible = True

TxtStateID = IIf(IsNull(RsMst_State!StateID), "", RsMst_State!StateID)
TxtStateName = IIf(IsNull(RsMst_State!StateName), "", RsMst_State!StateName)
TxtRemarks = IIf(IsNull(RsMst_State!Remarks), "", RsMst_State!Remarks)

Frame2.Enabled = False

Call IndataDRCDetails
Call GButtonLock(Me, True)

End Sub

Private Sub IndataDRCDetails()

Call Grid_Head1
'Call TableMst_StorageCardRateDetail(" where SRCID=" & Val(TxtSRCID.Text) & "")
tmp = "Select * from Mst_StateDRCMapping Where StateID = " & TxtStateID & ""
Call TmpTable

MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
For I = 1 To MSHFlexGrid2.Rows - 2
    tmp = " Select ME.EmployeeID,ME.EmployeeNo,ME.EmployeeName,MD.Designation,ME.Flag from Mst_Employee ME"
    tmp = tmp & " Inner Join Mst_Designation MD On  ME.DesignationID = MD.DesignationID"
    tmp = tmp & " Where EmployeeID = " & TmpRs!EmployeeID & " "
          
    Call TmpTableCMPro
    If TmpRsCMPro.RecordCount > 0 Then
       MSHFlexGrid2.TextMatrix(I, 0) = TmpRsCMPro!EmployeeName & "~" & TmpRsCMPro!EmployeeNo
    Else
       MSHFlexGrid2.TextMatrix(I, 0) = ""
    End If
    MSHFlexGrid2.TextMatrix(I, 1) = TmpRs!EmailID
    If TmpRs!Flag = "A" Then
       MSHFlexGrid2.TextMatrix(I, 2) = "Active"
    Else
       MSHFlexGrid2.TextMatrix(I, 2) = "De-Active"
    End If
    MSHFlexGrid2.TextMatrix(I, 3) = TmpRs!EmployeeID
    MSHFlexGrid2.TextMatrix(I, 4) = TmpRs!MappingID
    TmpRs.MoveNext
Next

End Sub



Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0, 0) = "StateID"
MSHFlexGrid1.TextMatrix(0, 1) = "State Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3200
MSHFlexGrid1.ColWidth(2) = 1000

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_State.MoveFirst
RsMst_State.Find "StateID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_State.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub CmdGo_Click()

tmp = "Select StateId,StateName,Flag from Mst_State "

Dim Wc As Variant
Wc = ""
If TxtSearchState <> "" Then
   Wc = " Where StateName Like '%" & TxtSearchState & "%' "
End If

tmp = tmp & Wc & " Order By StateName "

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

Private Sub SearchCmd_Click()

If SearchCmd.Caption = "Cancel" Then
   If RsMst_State.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Exit Sub
End If

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 600
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_State.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchState = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
